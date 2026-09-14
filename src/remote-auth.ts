import { createHash, randomBytes, randomUUID, timingSafeEqual } from "node:crypto"

import type { OAuthRegisteredClientsStore } from "@modelcontextprotocol/sdk/server/auth/clients.js"
import {
  AccessDeniedError,
  InvalidGrantError,
  InvalidRequestError,
  InvalidScopeError,
  InvalidTargetError,
  InvalidTokenError,
} from "@modelcontextprotocol/sdk/server/auth/errors.js"
import type { AuthorizationParams, OAuthServerProvider } from "@modelcontextprotocol/sdk/server/auth/provider.js"
import type { AuthInfo } from "@modelcontextprotocol/sdk/server/auth/types.js"
import type {
  OAuthClientInformationFull,
  OAuthTokenRevocationRequest,
  OAuthTokens,
} from "@modelcontextprotocol/sdk/shared/auth.js"
import type { Response } from "express"

export const MCP_SCOPE = "mcp:tools"

const AUTHORIZATION_REQUEST_TTL_MS = 10 * 60 * 1000
const AUTHORIZATION_CODE_TTL_MS = 5 * 60 * 1000
const ACCESS_TOKEN_TTL_MS = 60 * 60 * 1000
const REFRESH_TOKEN_TTL_MS = 30 * 24 * 60 * 60 * 1000
const MAX_REGISTERED_CLIENTS = 100

interface PendingAuthorization {
  client: OAuthClientInformationFull
  failedPasswordAttempts: number
  params: AuthorizationParams
  expiresAt: number
}

interface AuthorizationCodeData extends PendingAuthorization {
  code: string
}

interface TokenData {
  token: string
  clientId: string
  scopes: string[]
  resource: URL
  expiresAt: number
}

export class InMemoryClientsStore implements OAuthRegisteredClientsStore {
  private readonly clients = new Map<string, OAuthClientInformationFull>()

  async getClient(clientId: string): Promise<OAuthClientInformationFull | undefined> {
    return this.clients.get(clientId)
  }

  async registerClient(
    client: Omit<OAuthClientInformationFull, "client_id" | "client_id_issued_at">,
  ): Promise<OAuthClientInformationFull> {
    const registeredClient = client as OAuthClientInformationFull

    if (!registeredClient.client_id) {
      throw new InvalidRequestError("The authorization server did not generate a client_id")
    }
    if (this.clients.size >= MAX_REGISTERED_CLIENTS) {
      throw new InvalidRequestError("Too many OAuth clients are registered")
    }

    for (const redirectUri of registeredClient.redirect_uris) {
      const url = new URL(redirectUri)
      const isLoopback = ["localhost", "127.0.0.1", "[::1]"].includes(url.hostname)
      if (url.protocol !== "https:" && !(isLoopback && url.protocol === "http:")) {
        throw new InvalidRequestError("OAuth redirect URIs must use HTTPS or a loopback HTTP address")
      }
    }

    const requestedScopes = registeredClient.scope?.split(" ").filter(Boolean) ?? []
    if (requestedScopes.some((scope) => scope !== MCP_SCOPE)) {
      throw new InvalidScopeError(`Only the ${MCP_SCOPE} scope is supported`)
    }

    this.clients.set(registeredClient.client_id, registeredClient)
    return registeredClient
  }
}

export interface AuthorizationDecision {
  approved: boolean
  password?: string
  requestId: string
}

export class SingleUserOAuthProvider implements OAuthServerProvider {
  readonly clientsStore = new InMemoryClientsStore()

  private readonly pendingAuthorizations = new Map<string, PendingAuthorization>()
  private readonly authorizationCodes = new Map<string, AuthorizationCodeData>()
  private readonly accessTokens = new Map<string, TokenData>()
  private readonly refreshTokens = new Map<string, TokenData>()

  constructor(
    private readonly password: string,
    private readonly resourceUrl: URL,
  ) {}

  async authorize(client: OAuthClientInformationFull, params: AuthorizationParams, res: Response): Promise<void> {
    this.cleanupExpiredState()
    const resource = this.requireResource(params.resource)
    const scopes = this.validateScopes(params.scopes)
    const requestId = randomUUID()

    this.pendingAuthorizations.set(requestId, {
      client,
      failedPasswordAttempts: 0,
      params: { ...params, resource, scopes },
      expiresAt: Date.now() + AUTHORIZATION_REQUEST_TTL_MS,
    })

    const clientName = escapeHtml(client.client_name || "ChatGPT MCP client")
    const resourceLabel = escapeHtml(this.resourceUrl.href)
    res
      .status(200)
      .set("Cache-Control", "no-store")
      .type("html")
      .send(renderAuthorizationPage(requestId, clientName, resourceLabel))
  }

  completeAuthorization(decision: AuthorizationDecision): URL {
    this.cleanupExpiredState()
    const pending = this.pendingAuthorizations.get(decision.requestId)
    if (!pending) {
      throw new InvalidGrantError("The authorization request is invalid or has expired")
    }

    if (!decision.approved) {
      this.pendingAuthorizations.delete(decision.requestId)
      return this.createAuthorizationRedirect(pending, {
        error: "access_denied",
        error_description: "The resource owner denied the request",
      })
    }

    if (!decision.password || !safeSecretEqual(decision.password, this.password)) {
      pending.failedPasswordAttempts += 1
      if (pending.failedPasswordAttempts >= 5) {
        this.pendingAuthorizations.delete(decision.requestId)
      }
      throw new AccessDeniedError("The administrator password is incorrect")
    }

    this.pendingAuthorizations.delete(decision.requestId)
    const code = createOpaqueToken()
    this.authorizationCodes.set(code, {
      ...pending,
      code,
      expiresAt: Date.now() + AUTHORIZATION_CODE_TTL_MS,
    })

    return this.createAuthorizationRedirect(pending, { code })
  }

  async challengeForAuthorizationCode(client: OAuthClientInformationFull, authorizationCode: string): Promise<string> {
    this.cleanupExpiredState()
    const codeData = this.authorizationCodes.get(authorizationCode)
    if (!codeData || codeData.client.client_id !== client.client_id) {
      throw new InvalidGrantError("Invalid or expired authorization code")
    }
    return codeData.params.codeChallenge
  }

  async exchangeAuthorizationCode(
    client: OAuthClientInformationFull,
    authorizationCode: string,
    _codeVerifier?: string,
    redirectUri?: string,
    resource?: URL,
  ): Promise<OAuthTokens> {
    this.cleanupExpiredState()
    const codeData = this.authorizationCodes.get(authorizationCode)
    if (!codeData || codeData.client.client_id !== client.client_id) {
      throw new InvalidGrantError("Invalid or expired authorization code")
    }
    if (redirectUri && redirectUri !== codeData.params.redirectUri) {
      throw new InvalidGrantError("redirect_uri does not match the authorization request")
    }

    const requestedResource = this.requireResource(resource)
    if (requestedResource.href !== codeData.params.resource?.href) {
      throw new InvalidTargetError("resource does not match the authorization request")
    }

    this.authorizationCodes.delete(authorizationCode)
    return this.issueTokenPair(client.client_id, codeData.params.scopes ?? [MCP_SCOPE], requestedResource)
  }

  async exchangeRefreshToken(
    client: OAuthClientInformationFull,
    refreshToken: string,
    scopes?: string[],
    resource?: URL,
  ): Promise<OAuthTokens> {
    this.cleanupExpiredState()
    const tokenData = this.refreshTokens.get(refreshToken)
    if (!tokenData || tokenData.clientId !== client.client_id) {
      throw new InvalidGrantError("Invalid or expired refresh token")
    }

    const requestedResource = this.requireResource(resource)
    if (requestedResource.href !== tokenData.resource.href) {
      throw new InvalidTargetError("resource does not match the refresh token")
    }

    const requestedScopes = scopes?.length ? this.validateScopes(scopes) : tokenData.scopes
    if (requestedScopes.some((scope) => !tokenData.scopes.includes(scope))) {
      throw new InvalidScopeError("A refresh request cannot expand its original scopes")
    }

    this.refreshTokens.delete(refreshToken)
    return this.issueTokenPair(client.client_id, requestedScopes, requestedResource)
  }

  async verifyAccessToken(token: string): Promise<AuthInfo> {
    this.cleanupExpiredState()
    const tokenData = this.accessTokens.get(token)
    if (!tokenData) {
      throw new InvalidTokenError("Invalid or expired access token")
    }
    if (tokenData.resource.href !== this.resourceUrl.href) {
      throw new InvalidTokenError("The access token was issued for another resource")
    }

    return {
      token,
      clientId: tokenData.clientId,
      scopes: tokenData.scopes,
      expiresAt: Math.floor(tokenData.expiresAt / 1000),
      resource: tokenData.resource,
      extra: { subject: "single-user" },
    }
  }

  async revokeToken(client: OAuthClientInformationFull, request: OAuthTokenRevocationRequest): Promise<void> {
    const accessToken = this.accessTokens.get(request.token)
    if (accessToken?.clientId === client.client_id) {
      this.accessTokens.delete(request.token)
    }

    const refreshToken = this.refreshTokens.get(request.token)
    if (refreshToken?.clientId === client.client_id) {
      this.refreshTokens.delete(request.token)
    }
  }

  private issueTokenPair(clientId: string, scopes: string[], resource: URL): OAuthTokens {
    const accessToken = createOpaqueToken()
    const refreshToken = createOpaqueToken()
    const now = Date.now()

    this.accessTokens.set(accessToken, {
      token: accessToken,
      clientId,
      scopes,
      resource,
      expiresAt: now + ACCESS_TOKEN_TTL_MS,
    })
    this.refreshTokens.set(refreshToken, {
      token: refreshToken,
      clientId,
      scopes,
      resource,
      expiresAt: now + REFRESH_TOKEN_TTL_MS,
    })

    return {
      access_token: accessToken,
      refresh_token: refreshToken,
      token_type: "Bearer",
      expires_in: ACCESS_TOKEN_TTL_MS / 1000,
      scope: scopes.join(" "),
    }
  }

  private requireResource(resource?: URL): URL {
    if (!resource || resource.href !== this.resourceUrl.href) {
      throw new InvalidTargetError(`resource must be ${this.resourceUrl.href}`)
    }
    return resource
  }

  private validateScopes(scopes?: string[]): string[] {
    const requestedScopes = scopes?.length ? [...new Set(scopes)] : [MCP_SCOPE]
    if (requestedScopes.some((scope) => scope !== MCP_SCOPE)) {
      throw new InvalidScopeError(`Only the ${MCP_SCOPE} scope is supported`)
    }
    return requestedScopes
  }

  private createAuthorizationRedirect(pending: PendingAuthorization, result: Record<string, string>): URL {
    const redirect = new URL(pending.params.redirectUri)
    for (const [key, value] of Object.entries(result)) {
      redirect.searchParams.set(key, value)
    }
    if (pending.params.state) {
      redirect.searchParams.set("state", pending.params.state)
    }
    return redirect
  }

  private cleanupExpiredState(): void {
    const now = Date.now()
    for (const [key, value] of this.pendingAuthorizations) {
      if (value.expiresAt <= now) this.pendingAuthorizations.delete(key)
    }
    for (const [key, value] of this.authorizationCodes) {
      if (value.expiresAt <= now) this.authorizationCodes.delete(key)
    }
    for (const [key, value] of this.accessTokens) {
      if (value.expiresAt <= now) this.accessTokens.delete(key)
    }
    for (const [key, value] of this.refreshTokens) {
      if (value.expiresAt <= now) this.refreshTokens.delete(key)
    }
  }
}

function createOpaqueToken(): string {
  return randomBytes(32).toString("base64url")
}

function safeSecretEqual(actual: string, expected: string): boolean {
  const actualHash = createHash("sha256").update(actual).digest()
  const expectedHash = createHash("sha256").update(expected).digest()
  return timingSafeEqual(actualHash, expectedHash)
}

function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, (character) => {
    const entities: Record<string, string> = {
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#039;",
    }
    return entities[character]
  })
}

function renderAuthorizationPage(requestId: string, clientName: string, resource: string): string {
  return `<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>授权 Microsoft To Do MCP</title>
  <style>
    :root { color-scheme: light; font-family: Inter, ui-sans-serif, system-ui, sans-serif; }
    body { margin: 0; min-height: 100vh; display: grid; place-items: center; background: #f4f5f7; color: #172033; }
    main { width: min(420px, calc(100% - 40px)); padding: 32px; border: 1px solid #dfe3ea; border-radius: 16px; background: #fff; box-shadow: 0 18px 50px rgba(23, 32, 51, .1); }
    h1 { margin: 0 0 12px; font-size: 24px; }
    p { margin: 0 0 22px; color: #536076; line-height: 1.6; }
    code { overflow-wrap: anywhere; color: #243a73; }
    label { display: grid; gap: 8px; font-weight: 600; }
    input { box-sizing: border-box; width: 100%; padding: 12px; border: 1px solid #b9c1cf; border-radius: 9px; font: inherit; }
    .actions { display: flex; gap: 10px; margin-top: 20px; }
    button { flex: 1; padding: 12px; border: 0; border-radius: 9px; font: inherit; font-weight: 700; cursor: pointer; }
    .approve { background: #2563eb; color: #fff; }
    .deny { background: #e8ebf0; color: #303a4c; }
  </style>
</head>
<body>
  <main>
    <h1>授权 Microsoft To Do</h1>
    <p><strong>${clientName}</strong> 正在请求访问 <code>${resource}</code>。授权后它可以读取和修改这个单用户的 To Do 数据。</p>
    <form method="post" action="/oauth/approve">
      <input type="hidden" name="request_id" value="${escapeHtml(requestId)}">
      <label>管理员口令
        <input type="password" name="password" required autocomplete="current-password" autofocus>
      </label>
      <div class="actions">
        <button class="deny" type="submit" name="decision" value="deny" formnovalidate>拒绝</button>
        <button class="approve" type="submit" name="decision" value="approve">授权</button>
      </div>
    </form>
  </main>
</body>
</html>`
}
