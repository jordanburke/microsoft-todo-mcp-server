// HTTP mode: the To Do tools over stateless Streamable HTTP, behind somamcp's Azure provider.
// Each user signs in with their own Microsoft account, and each tool call uses that user's
// token. The server's own token file is never read.
import { mkdirSync, rmSync, writeFileSync } from "fs"
import { join } from "path"
import {
  type AzureProviderConfig,
  type AzureSession,
  createAzureProvider,
  createDiskStore,
  type OAuthAuthProvider,
} from "somamcp/node"

import { allowListRule, HTTP_SCOPES, type HttpConfig } from "./http-config.js"
import { sessionCredentials } from "./session-credentials.js"
import { createTodoServer } from "./todo-server.js"

export type HttpServer = {
  readonly server: ReturnType<typeof createTodoServer<AzureSession>>
  readonly provider: OAuthAuthProvider<AzureSession>
}

/** The Azure provider's settings for `config`, apart from the store. */
export const azureProviderConfig = (config: HttpConfig): Omit<AzureProviderConfig, "tokenStorage"> => ({
  baseUrl: config.baseUrl,
  clientId: config.clientId,
  clientSecret: config.clientSecret,
  tenantId: config.tenantId,
  scopes: HTTP_SCOPES,
  jwtSigningKey: config.jwtSigningKey,
  encryptionKey: config.encryptionKey,
  allowedRedirectUriPatterns: config.redirectUriPatterns,
  // OAuth 2.1 drops PKCE's plain method. somamcp 2.7+ refuses it by default; this pins that.
  allowPlainPkce: false,
  acceptSignIn: config.allowedUsers.length > 0 ? allowListRule(config.allowedUsers) : undefined,
})

/** The OAuth provider and the server for `config`, not yet listening. */
export const createHttpServer = (config: HttpConfig): HttpServer => {
  const provider = createAzureProvider({
    ...azureProviderConfig(config),
    tokenStorage: createDiskStore({ directory: config.storeDirectory }),
  })
  const server = createTodoServer<AzureSession>({ credentials: sessionCredentials, auth: provider })
  return { server, provider }
}

// The store must be writable before the server takes requests: otherwise it starts, says it is
// listening, and fails every client registration with a 500.
const checkStore = (directory: string): void => {
  mkdirSync(directory, { recursive: true, mode: 0o700 })
  const probe = join(directory, `.write-check-${process.pid}`)
  writeFileSync(probe, "")
  rmSync(probe)
}

/** Start listening, and say where. */
export const startHttpServer = async (config: HttpConfig): Promise<HttpServer> => {
  checkStore(config.storeDirectory)
  const http = createHttpServer(config)
  if (config.allowedUsers.length === 0) {
    console.error(
      "Warning: MSTODO_HTTP_ALLOWED_USERS is not set, so any Microsoft account the app registration accepts can sign in (each sees only its own To Do).",
    )
  }
  await http.server.start({
    transportType: "httpStream",
    httpStream: { host: config.host, port: config.port, endpoint: "/mcp" },
  })
  console.error(`Microsoft To Do MCP server listening on ${config.host}:${config.port}`)
  console.error(`Connect clients to ${config.baseUrl}/mcp`)
  console.error(`Register ${config.baseUrl}/oauth/callback as a redirect URI on the app registration`)
  return http
}
