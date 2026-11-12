// Settings for HTTP mode, read from the environment. Every problem is reported at once, and
// the server refuses to start rather than fall back to a weaker setting: without the signing
// and encryption keys, somamcp would generate random ones and sign everyone out on restart.
import { Either, Left, Right } from "functype"
import { homedir } from "os"
import { join, resolve } from "path"

export type HttpConfig = {
  /** The server's public URL, without a trailing slash: the OAuth issuer, and the base of /mcp. */
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
  /** Accounts allowed to sign in (sign-in name, email or object id), or empty for any. */
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

const list = (value: string | undefined): string[] =>
  (value ?? "")
    .split(",")
    .map((item) => item.trim())
    .filter((item) => item !== "")

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

  const baseUrl = required("MSTODO_HTTP_BASE_URL").replace(/\/+$/, "")
  if (baseUrl) {
    const parsed = URL.canParse(baseUrl) ? new URL(baseUrl) : undefined
    const local = parsed && ["localhost", "127.0.0.1", "[::1]"].includes(parsed.hostname)
    if (!parsed) problems.push(`MSTODO_HTTP_BASE_URL is not a URL: ${baseUrl}`)
    else if (parsed.protocol !== "https:" && !(parsed.protocol === "http:" && local))
      problems.push("MSTODO_HTTP_BASE_URL must use https (plain http is allowed only for localhost)")
  }

  const portText = env.MSTODO_HTTP_PORT?.trim()
  const port = portText ? Number(portText) : DEFAULT_PORT
  if (!Number.isInteger(port) || port < 1 || port > 65535) problems.push(`MSTODO_HTTP_PORT is not a port: ${portText}`)

  const redirectUriPatterns = env.MSTODO_HTTP_REDIRECT_URIS
    ? list(env.MSTODO_HTTP_REDIRECT_URIS)
    : DEFAULT_REDIRECT_URI_PATTERNS
  for (const pattern of redirectUriPatterns) {
    if (/^https?:\/\/\*/.test(pattern))
      problems.push(`MSTODO_HTTP_REDIRECT_URIS pattern ${pattern} lets anyone register a host they run`)
  }

  const config: HttpConfig = {
    baseUrl,
    host: env.MSTODO_HTTP_HOST?.trim() || "127.0.0.1",
    port,
    clientId: required("CLIENT_ID"),
    clientSecret: required("CLIENT_SECRET"),
    tenantId: required("TENANT_ID"),
    jwtSigningKey: key("JWT_SIGNING_KEY"),
    encryptionKey: key("TOKEN_ENCRYPTION_KEY"),
    storeDirectory: resolve(
      env.MSTODO_HTTP_STORE?.trim() || join(homedir(), ".config", "microsoft-todo-mcp", "http-store"),
    ),
    redirectUriPatterns,
    allowedUsers: list(env.MSTODO_HTTP_ALLOWED_USERS).map((user) => user.toLowerCase()),
  }

  return problems.length > 0 ? Left(problems) : Right(config)
}

/**
 * The `acceptSignIn` rule for the allow-list: the reason to refuse a sign-in, or undefined to
 * accept it. An account matches by sign-in name, email or object id, ignoring case.
 */
export const allowListRule =
  (allowedUsers: readonly string[]) =>
  (claims: Readonly<Record<string, unknown>>): string | undefined => {
    if (allowedUsers.length === 0) return undefined
    const names = ["preferred_username", "upn", "email", "oid"]
      .map((claim) => claims[claim])
      .filter((value): value is string => typeof value === "string" && value !== "")
      .map((value) => value.toLowerCase())
    return names.some((name) => allowedUsers.includes(name))
      ? undefined
      : "This account is not allowed to use this server."
  }
