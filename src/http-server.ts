// HTTP mode: the To Do tools over stateless Streamable HTTP, behind somamcp's Azure provider.
// Each user signs in with their own Microsoft account, and each tool call uses that user's
// token. The server's own token file is never read.
import { type AzureSession, createAzureProvider, createDiskStore, type OAuthAuthProvider } from "somamcp/node"

import { allowListRule, HTTP_SCOPES, type HttpConfig } from "./http-config.js"
import { sessionCredentials } from "./session-credentials.js"
import { createTodoServer } from "./todo-server.js"

export type HttpServer = {
  readonly server: ReturnType<typeof createTodoServer<AzureSession>>
  readonly provider: OAuthAuthProvider<AzureSession>
}

/** The OAuth provider and the server for `config`, not yet listening. */
export const createHttpServer = (config: HttpConfig): HttpServer => {
  const provider = createAzureProvider({
    baseUrl: config.baseUrl,
    clientId: config.clientId,
    clientSecret: config.clientSecret,
    tenantId: config.tenantId,
    scopes: HTTP_SCOPES,
    jwtSigningKey: config.jwtSigningKey,
    encryptionKey: config.encryptionKey,
    tokenStorage: createDiskStore({ directory: config.storeDirectory }),
    allowedRedirectUriPatterns: config.redirectUriPatterns,
    acceptSignIn: config.allowedUsers.length > 0 ? allowListRule(config.allowedUsers) : undefined,
  })
  const server = createTodoServer<AzureSession>({ credentials: sessionCredentials, auth: provider })
  return { server, provider }
}

/** Start listening, and say where. */
export const startHttpServer = async (config: HttpConfig): Promise<HttpServer> => {
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
