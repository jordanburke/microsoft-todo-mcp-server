// Settings for HTTP mode, read from the environment. Every problem is reported at once, and
// the server refuses to start rather than fall back to a weaker setting: without the signing
// and encryption keys, somamcp would generate random ones and sign everyone out on restart.
import { Either, Left, Right } from "functype"
import { homedir } from "os"
import { join, resolve } from "path"

export type HttpConfig = {
  /** The server's public origin, such as `https://todo.example.com`: the OAuth issuer, and the base of /mcp. */
  readonly baseUrl: string
  readonly host: string
  readonly port: number
  /** The Azure app registration that users sign in through. */
  readonly clientId: string
  readonly clientSecret: string
  readonly tenantId: string
  readonly jwtSigningKey: string
  readonly encryptionKey: string
  /** Where sign-ins and Microsoft's tokens are kept, encrypted. */
  readonly storeDirectory: string
  /** Redirect URIs an MCP client may register, as globs. */
  readonly redirectUriPatterns: readonly string[]
  /** Accounts allowed to sign in (object id, sign-in name or email), lower-cased; empty for any. */
  readonly allowedUsers: readonly string[]
}

/** The Graph scopes a sign-in asks for. `openid` makes Microsoft send an id token naming the user. */
export const HTTP_SCOPES = [
  "openid",
  "profile",
  "email",
  "offline_access",
  "User.Read",
  "Tasks.Read",
  "Tasks.ReadWrite",
  "Tasks.Read.Shared",
  "Tasks.ReadWrite.Shared",
] as const

const DEFAULT_PORT = 3100
const DEFAULT_REDIRECT_URI_PATTERNS = ["https://claude.ai/*", "http://localhost:*/*", "http://127.0.0.1:*/*"]
const MIN_KEY_LENGTH = 32
const LOOPBACK = ["localhost", "127.0.0.1", "[::1]"]
// A tenant's id (a GUID), one of Entra ID's shared names, or a verified domain.
const TENANT =
  /^([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}|common|organizations|consumers|[a-z0-9-]+(\.[a-z0-9-]+)+)$/i

const list = (value: string): string[] =>
  value
    .split(",")
    .map((item) => item.trim())
    .filter((item) => item !== "")

/**
 * Why a redirect URI pattern is unsafe, or undefined if it is safe. somamcp's `*` and `?` match
 * any characters, including `.` and `/`, so a wildcard anywhere before the path lets a client
 * register a host it runs (`https://claude.ai*` matches `https://claude.ai.evil.example/`). The
 * scheme and host must be literal; only a loopback host may leave the port open.
 */
export const redirectPatternProblem = (pattern: string): string | undefined => {
  // The host is a bracketed IPv6 address or a name with no / : ? # [ ].
  const parts = /^(https?):\/\/(\[[0-9a-fA-F:.]+\]|[^/:?#[\]]*)(?::([^/?#]*))?(\/.*)?$/.exec(pattern)
  const [, scheme, host, port] = parts ?? []
  const unsafe = `MSTODO_HTTP_REDIRECT_URIS pattern ${pattern}`
  if (!parts) return `${unsafe} must start with http:// or https:// and a host`
  if (!host || /[*?@]/.test(host)) return `${unsafe} must name its host exactly, with no wildcard`
  // somamcp matches case-sensitively, and a browser sends the host in lower case.
  if (host !== host.toLowerCase()) return `${unsafe} must write its host in lower case`
  const loopback = LOOPBACK.includes(host)
  if (scheme === "http" && !loopback) return `${unsafe} must use https unless its host is localhost`
  if (port !== undefined && !/^\d+$/.test(port) && !(loopback && port === "*"))
    return `${unsafe} may leave the port open only for localhost`
  return undefined
}

/** HTTP mode's settings, or every reason they are not usable. */
export const readHttpConfig = (env: NodeJS.ProcessEnv): Either<string[], HttpConfig> => {
  const problems: string[] = []
  const required = (name: string): string => {
    const value = env[name]?.trim()
    if (!value) problems.push(`${name} is not set`)
    return value ?? ""
  }
  const key = (name: string): string => {
    const value = required(name)
    if (value && value.length < MIN_KEY_LENGTH)
      problems.push(`${name} must be at least ${MIN_KEY_LENGTH} characters (try: openssl rand -base64 48)`)
    return value
  }

  // The OAuth routes live at the server's root, so the base URL is an origin: no path, query,
  // fragment or credentials. A reverse proxy that mounts it under a path must strip the path.
  const baseUrl = required("MSTODO_HTTP_BASE_URL").replace(/\/+$/, "")
  if (baseUrl) {
    const parsed = URL.canParse(baseUrl) ? new URL(baseUrl) : undefined
    if (!parsed) problems.push(`MSTODO_HTTP_BASE_URL is not a URL: ${baseUrl}`)
    else {
      if (parsed.protocol !== "https:" && !(parsed.protocol === "http:" && LOOPBACK.includes(parsed.hostname)))
        problems.push("MSTODO_HTTP_BASE_URL must use https (plain http is allowed only for localhost)")
      if (parsed.username || parsed.password)
        problems.push("MSTODO_HTTP_BASE_URL must not hold a user name or password")
      if (parsed.pathname !== "/" || parsed.search || parsed.hash)
        problems.push(
          "MSTODO_HTTP_BASE_URL must be an origin only, such as https://todo.example.com (no path, query or fragment)",
        )
    }
  }

  const portText = env.MSTODO_HTTP_PORT?.trim()
  const port = portText ? Number(portText) : DEFAULT_PORT
  if ((portText && !/^\d+$/.test(portText)) || port < 1 || port > 65535)
    problems.push(`MSTODO_HTTP_PORT is not a port: ${portText}`)

  const clientId = required("CLIENT_ID")
  const clientSecret = required("CLIENT_SECRET")
  const tenantId = required("TENANT_ID")
  if (tenantId && !TENANT.test(tenantId))
    problems.push(
      `TENANT_ID must be a tenant id (a GUID), a domain, or common, organizations or consumers: ${tenantId}`,
    )

  const jwtSigningKey = key("JWT_SIGNING_KEY")
  const encryptionKey = key("TOKEN_ENCRYPTION_KEY")
  if (jwtSigningKey && jwtSigningKey === encryptionKey)
    problems.push("JWT_SIGNING_KEY and TOKEN_ENCRYPTION_KEY must differ")

  const redirectUriPatterns =
    env.MSTODO_HTTP_REDIRECT_URIS === undefined ? DEFAULT_REDIRECT_URI_PATTERNS : list(env.MSTODO_HTTP_REDIRECT_URIS)
  if (redirectUriPatterns.length === 0) problems.push("MSTODO_HTTP_REDIRECT_URIS is set but lists no patterns")
  for (const pattern of redirectUriPatterns) {
    const problem = redirectPatternProblem(pattern)
    if (problem) problems.push(problem)
  }

  const allowedUsers =
    env.MSTODO_HTTP_ALLOWED_USERS === undefined
      ? []
      : list(env.MSTODO_HTTP_ALLOWED_USERS).map((user) => user.toLowerCase())
  if (env.MSTODO_HTTP_ALLOWED_USERS !== undefined && allowedUsers.length === 0)
    problems.push("MSTODO_HTTP_ALLOWED_USERS is set but lists no accounts (unset it to allow any account)")

  const config: HttpConfig = {
    baseUrl,
    host: env.MSTODO_HTTP_HOST?.trim() || "127.0.0.1",
    port,
    clientId,
    clientSecret,
    tenantId,
    jwtSigningKey,
    encryptionKey,
    storeDirectory: resolve(
      env.MSTODO_HTTP_STORE?.trim() || join(homedir(), ".config", "microsoft-todo-mcp", "http-store"),
    ),
    redirectUriPatterns,
    allowedUsers,
  }

  return problems.length > 0 ? Left(problems) : Right(config)
}

/**
 * The `acceptSignIn` rule for the allow-list: the reason to refuse a sign-in, or undefined to
 * accept it. An account matches by object id, sign-in name or email, ignoring case.
 */
export const allowListRule =
  (allowedUsers: readonly string[]) =>
  (claims: Readonly<Record<string, unknown>>): string | undefined => {
    if (allowedUsers.length === 0) return undefined
    const names = ["oid", "preferred_username", "upn", "email"]
      .map((claim) => claims[claim])
      .filter((value): value is string => typeof value === "string" && value !== "")
      .map((value) => value.toLowerCase())
    return names.some((name) => allowedUsers.includes(name))
      ? undefined
      : "This account is not allowed to use this server."
  }
