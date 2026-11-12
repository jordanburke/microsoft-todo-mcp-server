// src/token-manager.ts
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "fs"
import { Http, None, Option, Some, Try } from "functype"
import { homedir } from "os"
import { dirname, join, resolve } from "path"

interface TokenData {
  accessToken: string
  refreshToken: string
  expiresAt: number
}

interface StoredTokenData extends TokenData {
  clientId?: string
  clientSecret?: string
  tenantId?: string
}

type TokenResponse = { access_token: string; refresh_token?: string; expires_in: number }

const FALLBACK_LIFETIME_MS = 3600 * 1000
const REFRESH_BUFFER_MS = 5 * 60 * 1000
const SCOPE = "offline_access Tasks.Read Tasks.ReadWrite Tasks.Read.Shared Tasks.ReadWrite.Shared User.Read"

// Work and school accounts get a JWT access token, so the real expiry is in its `exp`
// claim and does not have to be guessed. Personal accounts get an opaque token, where
// the one-hour assumption is all there is — but the caller stamps it once instead of
// on every call, which is what used to keep the refresh branch out of reach.
const expiryFromAccessToken = (accessToken: string): number =>
  Option(accessToken.split(".")[1])
    .flatMap((payload) =>
      Try(() => JSON.parse(Buffer.from(payload, "base64url").toString("utf8")) as { exp?: unknown }).toOption(),
    )
    .flatMap(({ exp }) => (typeof exp === "number" ? Some(exp * 1000 - REFRESH_BUFFER_MS) : None<number>()))
    .orElse(Date.now() + FALLBACK_LIFETIME_MS)

// A token file's contents, or None when it is missing or unreadable.
const readTokenFile = (path: string, label: string): Option<StoredTokenData> =>
  existsSync(path)
    ? Try(() => JSON.parse(readFileSync(path, "utf8")) as StoredTokenData | null).fold(
        (error) => {
          console.error(`Error reading ${label}:`, error)
          return None<StoredTokenData>()
        },
        (tokens) => Option(tokens),
      )
    : None<StoredTokenData>()

export class TokenManager {
  private tokenFilePath: string
  private currentTokens: Option<StoredTokenData> = None()
  private envTokens: Option<TokenData> = None()
  // A renewal under way, keyed by the refresh token it spends. Entra rotates the refresh
  // token on use, so a second concurrent renewal with the same one would race the first.
  private inFlight: Option<{ refreshToken: string; renewal: Promise<TokenData | null> }> = None()

  constructor() {
    this.tokenFilePath = TokenManager.resolveTokenFilePath()
    mkdirSync(dirname(this.tokenFilePath), { recursive: true })
    console.error(`Token file path: ${this.tokenFilePath}`)
  }

  private static resolveTokenFilePath(): string {
    // MSTODO_TOKEN_FILE is documented in the README as the way to relocate the token
    // store. It used to be read in cli.ts and handed to startServer, which discarded
    // it, so setting it moved nothing and every tool call failed to authenticate.
    const override = process.env.MSTODO_TOKEN_FILE
    if (override) return resolve(override)

    // Otherwise a consistent location across platforms
    const configDir =
      process.platform === "win32"
        ? join(process.env.APPDATA || join(homedir(), "AppData", "Roaming"), "microsoft-todo-mcp")
        : join(homedir(), ".config", "microsoft-todo-mcp")

    return join(configDir, "tokens.json")
  }

  // The environment pair is a bootstrap credential, not a live one: nothing rewrites
  // those variables when a token is renewed. Stamp its deadline once, when the value
  // is first seen, so the deadline stops moving with the clock.
  private readEnvTokens(): Option<TokenData> {
    const accessToken = process.env.MS_TODO_ACCESS_TOKEN
    const refreshToken = process.env.MS_TODO_REFRESH_TOKEN
    if (!accessToken || !refreshToken) return None()

    const seen = this.envTokens.filter((tokens) => tokens.accessToken === accessToken)
    this.envTokens = seen.isSome()
      ? seen
      : Some({ accessToken, refreshToken, expiresAt: expiryFromAccessToken(accessToken) })

    return this.envTokens
  }

  // Try to get tokens from multiple sources
  async getTokens(): Promise<TokenData | null> {
    // 1. Environment variables first (for backward compatibility)
    const envTokens = this.readEnvTokens()
    if (envTokens.isSome()) {
      // Anything renewed in this process supersedes the environment pair. Nothing
      // rewrites those variables, so the original values stay frozen at whatever they
      // were at launch; preferring them would serve a stale access token forever and
      // would re-send a refresh token the identity provider has already rotated away.
      const live: TokenData = this.currentTokens.orElse(envTokens.value)
      if (Date.now() < live.expiresAt) return live

      // Hand back the expired token rather than nothing: the caller's 401 says more
      // about what went wrong than a bare "not authenticated".
      return (await this.refreshToken(live.refreshToken)) ?? live
    }

    // 2. The token file
    const stored = readTokenFile(this.tokenFilePath, "token file")
    if (stored.isSome()) {
      this.currentTokens = stored
      const tokens = stored.value
      if (Date.now() > tokens.expiresAt) {
        return (await this.refreshToken(tokens.refreshToken)) ?? tokens
      }
      return tokens
    }

    // 3. The legacy ./tokens.json, migrated to the token file as found
    const legacy = readTokenFile(join(process.cwd(), "tokens.json"), "legacy token file")
    if (legacy.isSome()) {
      this.saveTokens(legacy.value)
      return legacy.value
    }

    return null
  }

  // A 401 proves the access token is dead, whatever its recorded expiry says. Personal
  // accounts get opaque tokens whose expiry can only be assumed, so an environment seed
  // minted on an earlier day still looks valid for its first hour in this process, and
  // getTokens() would hand the same dead token straight back. Refresh it now instead.
  async refreshRejectedToken(rejectedAccessToken: string): Promise<TokenData | null> {
    const live = (this.currentTokens.isSome() ? this.currentTokens : this.readEnvTokens()).orNull()
    if (live === null) return null

    // Another request already replaced the token this one was sent with.
    if (live.accessToken !== rejectedAccessToken) return live

    return this.refreshToken(live.refreshToken)
  }

  /** Renew with `refreshToken`, joining a renewal already under way for the same token. */
  async refreshToken(refreshToken: string): Promise<TokenData | null> {
    const pending = this.inFlight.filter((f) => f.refreshToken === refreshToken)
    if (pending.isSome()) return pending.value.renewal

    const renewal = this.renew(refreshToken).finally(() => {
      this.inFlight = None()
    })
    this.inFlight = Some({ refreshToken, renewal })
    return renewal
  }

  private async renew(refreshToken: string): Promise<TokenData | null> {
    // Client credentials come from the stored tokens, then the environment
    const stored = this.currentTokens.orUndefined()
    const clientId = stored?.clientId || process.env.CLIENT_ID
    const clientSecret = stored?.clientSecret || process.env.CLIENT_SECRET
    const tenantId = stored?.tenantId || process.env.TENANT_ID || "organizations"

    if (!clientId || !clientSecret) {
      console.error("Missing client credentials for token refresh")
      return null
    }

    const result = await Http.post<TokenResponse>(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        refresh_token: refreshToken,
        grant_type: "refresh_token",
        scope: SCOPE,
      }).toString(),
      parseAs: "json",
    }).run()

    if (result.isLeft()) {
      const error = result.value
      console.error(
        error && typeof error === "object" && "_tag" in error && error._tag === "HttpStatusError"
          ? `Token refresh failed: ${error.body}`
          : `Error refreshing token: ${String((error as { cause?: unknown })?.cause ?? error)}`,
      )
      this.promptForReauth()
      return null
    }

    const data = result.value.data
    const newTokens: StoredTokenData = {
      accessToken: data.access_token,
      refreshToken: data.refresh_token || refreshToken,
      expiresAt: Date.now() + data.expires_in * 1000 - REFRESH_BUFFER_MS,
      clientId,
      clientSecret,
      tenantId,
    }

    this.saveTokens(newTokens)
    await this.updateClaudeConfig(newTokens)
    return newTokens
  }

  saveTokens(tokens: StoredTokenData): void {
    this.currentTokens = Some(tokens)
    writeFileSync(this.tokenFilePath, JSON.stringify(tokens, null, 2), "utf8")
  }

  // Update Claude config automatically
  async updateClaudeConfig(tokens: TokenData): Promise<void> {
    const claudeConfigPath =
      process.platform === "win32"
        ? join(process.env.APPDATA || "", "Claude", "claude_desktop_config.json")
        : process.platform === "darwin"
          ? join(homedir(), "Library", "Application Support", "Claude", "claude_desktop_config.json")
          : join(homedir(), ".config", "Claude", "claude_desktop_config.json")

    if (!existsSync(claudeConfigPath)) return

    Try(() => {
      const config = JSON.parse(readFileSync(claudeConfigPath, "utf8"))
      const entry = config.mcpServers?.["microsoft-todo"]
      if (!entry) return

      entry.env = {
        ...entry.env,
        MS_TODO_ACCESS_TOKEN: tokens.accessToken,
        MS_TODO_REFRESH_TOKEN: tokens.refreshToken,
      }
      writeFileSync(claudeConfigPath, JSON.stringify(config, null, 2), "utf8")
      console.error("Updated Claude config with new tokens")
    }).fold(
      (error) => console.error("Could not update Claude config:", error),
      () => undefined,
    )
  }

  promptForReauth(): void {
    console.error(`
=================================================================
TOKEN REFRESH FAILED - REAUTHENTICATION REQUIRED

Your Microsoft To Do tokens have expired and could not be refreshed.

To fix this:
1. Open a new terminal
2. Navigate to the microsoft-todo-mcp-server directory
3. Run: pnpm run auth
4. Complete the authentication in your browser
5. Restart Claude Desktop to use the new tokens

Your tokens are stored in: ${this.tokenFilePath}
=================================================================
    `)
  }

  // Store client credentials with tokens for future refreshes
  async storeCredentials(clientId: string, clientSecret: string, tenantId: string): Promise<void> {
    this.currentTokens.fold(
      () => undefined,
      (tokens) => this.saveTokens({ ...tokens, clientId, clientSecret, tenantId }),
    )
  }
}

export const tokenManager = new TokenManager()
