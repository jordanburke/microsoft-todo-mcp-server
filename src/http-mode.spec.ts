// HTTP mode: its settings, its sign-in allow-list, and above all that each caller's tool
// calls use that caller's own Microsoft token, never the host's.
import { mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from "fs"
import { tmpdir } from "os"
import { dirname, join, resolve } from "path"
import type { AuthProvider } from "somamcp"
import { type AzureSession, createAzureProvider } from "somamcp/node"
import { fileURLToPath } from "url"
import { afterEach, beforeEach, describe, expect, it } from "vitest"

import { allowListRule, type HttpConfig, readHttpConfig, redirectPatternProblem } from "./http-config.js"
import { azureProviderConfig, createHttpServer } from "./http-server.js"
import { sessionCredentials } from "./session-credentials.js"
import { FakeNetwork, graphError, json, resetNetwork, restoreNetwork } from "./test-support/graph-harness.js"
import { createTodoServer } from "./todo-server.js"

const validEnv = {
  MSTODO_HTTP_BASE_URL: "https://todo.example.com/",
  CLIENT_ID: "client",
  CLIENT_SECRET: "secret",
  TENANT_ID: "consumers",
  JWT_SIGNING_KEY: "j".repeat(48),
  TOKEN_ENCRYPTION_KEY: "e".repeat(48),
}

describe("readHttpConfig", () => {
  it("lists every missing setting at once", () => {
    expect(
      readHttpConfig({}).fold(
        (problems) => problems,
        () => [],
      ),
    ).toEqual([
      "MSTODO_HTTP_BASE_URL is not set",
      "CLIENT_ID is not set",
      "CLIENT_SECRET is not set",
      "TENANT_ID is not set",
      "JWT_SIGNING_KEY is not set",
      "TOKEN_ENCRYPTION_KEY is not set",
    ])
  })

  it("fills in the defaults", () => {
    const config = readHttpConfig(validEnv).orThrow(new Error("expected a config"))

    expect(config).toMatchObject({
      baseUrl: "https://todo.example.com",
      host: "127.0.0.1",
      port: 3100,
      redirectUriPatterns: ["https://claude.ai/*", "http://localhost:*/*", "http://127.0.0.1:*/*"],
      allowedUsers: [],
    })
    expect(config.storeDirectory).toBe(join(process.env.HOME ?? "", ".config", "microsoft-todo-mcp", "http-store"))
  })

  it.each([
    [{ JWT_SIGNING_KEY: "short" }, "JWT_SIGNING_KEY must be at least 32 characters"],
    [{ TOKEN_ENCRYPTION_KEY: "short" }, "TOKEN_ENCRYPTION_KEY must be at least 32 characters"],
    [{ MSTODO_HTTP_BASE_URL: "http://todo.example.com" }, "MSTODO_HTTP_BASE_URL must use https"],
    [{ MSTODO_HTTP_BASE_URL: "not a url" }, "MSTODO_HTTP_BASE_URL is not a URL"],
    [{ MSTODO_HTTP_PORT: "70000" }, "MSTODO_HTTP_PORT is not a port"],
    [{ MSTODO_HTTP_PORT: "0x1F90" }, "MSTODO_HTTP_PORT is not a port"],
    [{ MSTODO_HTTP_PORT: "3430.0" }, "MSTODO_HTTP_PORT is not a port"],
    [{ MSTODO_HTTP_BASE_URL: "https://todo.example.com/todo" }, "must be an origin only"],
    [{ MSTODO_HTTP_BASE_URL: "https://todo.example.com/?a=1" }, "must be an origin only"],
    [{ MSTODO_HTTP_BASE_URL: "https://todo.example.com/#frag" }, "must be an origin only"],
    [{ MSTODO_HTTP_BASE_URL: "https://todo.example.com?" }, "must be an origin only"],
    [{ MSTODO_HTTP_BASE_URL: "https://todo.example.com#" }, "must be an origin only"],
    [{ MSTODO_HTTP_BASE_URL: "https://user:pw@todo.example.com" }, "must not hold a user name or password"],
    [{ TENANT_ID: "../../evil" }, "TENANT_ID must be a tenant id"],
    [{ TOKEN_ENCRYPTION_KEY: "j".repeat(48) }, "must differ"],
    [{ MSTODO_HTTP_REDIRECT_URIS: " , " }, "lists no patterns"],
    [{ MSTODO_HTTP_ALLOWED_USERS: " , " }, "lists no accounts"],
    [{ MSTODO_HTTP_REDIRECT_URIS: "https://claude.ai*" }, "must name its host exactly"],
  ])("refuses %j", (override, problem) => {
    const problems = readHttpConfig({ ...validEnv, ...override }).fold(
      (found) => found,
      () => [],
    )

    expect(problems.join("\n")).toContain(problem)
  })

  it("allows plain http for localhost, and reads the lists", () => {
    const config = readHttpConfig({
      ...validEnv,
      MSTODO_HTTP_BASE_URL: "http://localhost:3100",
      MSTODO_HTTP_REDIRECT_URIS: "https://claude.ai/*, https://app.example.com/callback",
      MSTODO_HTTP_ALLOWED_USERS: "Someone@Outlook.com, 00000000-aaaa",
      MSTODO_HTTP_STORE: "store",
    }).orThrow(new Error("expected a config"))

    expect(config.redirectUriPatterns).toEqual(["https://claude.ai/*", "https://app.example.com/callback"])
    expect(config.allowedUsers).toEqual(["someone@outlook.com", "00000000-aaaa"])
    expect(config.storeDirectory).toBe(resolve("store"))
  })
})

describe("redirectPatternProblem", () => {
  // Each of these let a client register a host it runs, because somamcp's * matches . and /.
  it.each([
    "https://*",
    "http://localhost*",
    "https://claude.ai*",
    "*",
    "*://claude.ai/*",
    "https://app.*.com/cb",
    "*claude.ai/*",
    "http*://claude.ai/*",
    "https://?*",
    "https://claude.ai@evil.example/*",
    "https://claude.ai:*/*",
    "http://claude.ai/*",
    "claude.ai/*",
    "http://LOCALHOST:*/*",
    "http://[::1]evil.example/*",
    "https://\\/*",
    "https://\t/*",
    "https://claude ai/*",
    "https://claude_ai.example/*",
    "https://.claude.ai/*",
    "https://claude.ai:0/*",
    "https://claude.ai:65536/*",
    "https://app.example.com:99999/cb",
    "https://claude.ai:000080/*",
    "https://1.2.3.999/*",
    "https://[:::]/*",
    "https://[1::2::3]/*",
  ])("refuses %s", (pattern) => {
    expect(redirectPatternProblem(pattern)).toBeDefined()
  })

  it.each([
    "https://claude.ai/*",
    "http://localhost:*/*",
    "http://127.0.0.1:*/*",
    "https://app.example.com/callback",
    "https://app.example.com:8443/cb/*",
    "http://localhost:3000/cb",
    "http://[::1]:*/*",
  ])("accepts %s, and so does somamcp's provider", (pattern) => {
    expect(redirectPatternProblem(pattern)).toBeUndefined()
    // somamcp refuses an unsafe pattern when the provider is built; the two checks must agree,
    // or a config this server accepts would still fail to start.
    const config = readHttpConfig({ ...validEnv, MSTODO_HTTP_REDIRECT_URIS: pattern }).orThrow(new Error("x"))
    expect(() => createAzureProvider(azureProviderConfig(config)).close()).not.toThrow()
  })

  it("is refused by somamcp's provider too, when this check is bypassed", () => {
    const config = readHttpConfig(validEnv).orThrow(new Error("x"))
    expect(() =>
      createAzureProvider({ ...azureProviderConfig(config), allowedRedirectUriPatterns: ["https://claude.ai*"] }),
    ).toThrow(/allowedRedirectUriPatterns/)
  })
})

describe("azureProviderConfig", () => {
  it("passes every security setting to the provider", () => {
    const config = readHttpConfig({
      ...validEnv,
      TENANT_ID: "11111111-2222-3333-4444-555555555555",
      MSTODO_HTTP_ALLOWED_USERS: "oid-1",
    }).orThrow(new Error("expected a config"))

    const provider = azureProviderConfig(config)

    expect(provider).toMatchObject({
      baseUrl: "https://todo.example.com",
      clientId: "client",
      clientSecret: "secret",
      tenantId: "11111111-2222-3333-4444-555555555555",
      jwtSigningKey: "j".repeat(48),
      encryptionKey: "e".repeat(48),
      allowedRedirectUriPatterns: ["https://claude.ai/*", "http://localhost:*/*", "http://127.0.0.1:*/*"],
      allowPlainPkce: false,
    })
    expect(provider.scopes).toEqual(expect.arrayContaining(["openid", "offline_access", "Tasks.ReadWrite"]))
    expect(provider.acceptSignIn?.({ oid: "oid-1" })).toBeUndefined()
    expect(provider.acceptSignIn?.({ oid: "someone-else" })).toBe("This account is not allowed to use this server.")
  })

  it("sets no sign-in rule without an allow-list", () => {
    expect(azureProviderConfig(readHttpConfig(validEnv).orThrow(new Error("x"))).acceptSignIn).toBeUndefined()
  })
})

describe("allowListRule", () => {
  const rule = allowListRule(["someone@outlook.com", "oid-1"])

  it("accepts a listed sign-in name or object id, ignoring case", () => {
    expect(rule({ preferred_username: "SomeOne@outlook.com" })).toBeUndefined()
    expect(rule({ oid: "oid-1", preferred_username: "other@outlook.com" })).toBeUndefined()
  })

  it("refuses anyone else, and a sign-in with no claims", () => {
    expect(rule({ preferred_username: "other@outlook.com" })).toBe("This account is not allowed to use this server.")
    expect(rule({})).toBeDefined()
  })

  it("accepts everyone when the list is empty", () => {
    expect(allowListRule([])({ preferred_username: "anyone@example.com" })).toBeUndefined()
  })
})

// A stand-in for the Azure provider: each bearer maps to one user's session.
const SESSIONS: Record<string, AzureSession> = {
  "bearer-a": { accessToken: "graph-a", claims: {}, scopes: [], upn: "a@example.com", expiresAt: 4102444800 },
  "bearer-b": { accessToken: "graph-b", claims: {}, scopes: [], upn: "b@example.com" },
}
const fakeProvider: AuthProvider<AzureSession> = {
  authenticate: async (request) => {
    const bearer = request?.headers.get("authorization")?.replace(/^Bearer\s+/i, "") ?? ""
    return SESSIONS[bearer] as AzureSession
  },
}

const callAs = async (
  server: ReturnType<typeof createTodoServer<AzureSession>>,
  bearer: string | undefined,
  name: string,
  args: Record<string, unknown> = {},
): Promise<{ status: number; text: string }> => {
  const response = await server.fetch(
    new Request("http://localhost/mcp", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        accept: "application/json, text/event-stream",
        ...(bearer ? { authorization: `Bearer ${bearer}` } : {}),
      },
      body: JSON.stringify({ jsonrpc: "2.0", id: 1, method: "tools/call", params: { name, arguments: args } }),
    }),
  )
  const raw = await response.text()
  const data = raw.split("\n").find((line) => line.startsWith("data: "))
  const message = data ? JSON.parse(data.slice("data: ".length)) : undefined
  return { status: response.status, text: message?.result?.content?.[0]?.text ?? "" }
}

describe("tool calls over HTTP", () => {
  const net = new FakeNetwork()
  let server: ReturnType<typeof createTodoServer<AzureSession>>

  beforeEach(() => {
    // The host has tokens of its own, in the environment and in its token file. HTTP mode
    // must never send them.
    process.env.MS_TODO_ACCESS_TOKEN = "host-access-token"
    process.env.MS_TODO_REFRESH_TOKEN = "host-refresh-token"
    writeFileSync(
      process.env.MSTODO_TOKEN_FILE ?? "",
      JSON.stringify({ accessToken: "host-file-token", refreshToken: "host-file-refresh", expiresAt: 0 }),
    )
    resetNetwork(net)
    server = createTodoServer<AzureSession>({ credentials: sessionCredentials, auth: fakeProvider })
  })
  afterEach(() => {
    delete process.env.MS_TODO_ACCESS_TOKEN
    delete process.env.MS_TODO_REFRESH_TOKEN
    restoreNetwork(net)
  })

  it("sends each caller's own Microsoft token", async () => {
    net.on("GET", "/me/todo/lists", json({ value: [{ id: "L1", displayName: "Tasks" }] }))

    await callAs(server, "bearer-a", "get-task-lists")
    await callAs(server, "bearer-b", "get-task-lists")

    expect(net.calls.map((call) => call.headers.authorization)).toEqual(["Bearer graph-a", "Bearer graph-b"])
  })

  it("reports Graph's 401 without renewing, and never sends the host's tokens", async () => {
    net.on("GET", "/me/todo/lists", graphError(401, "InvalidAuthenticationToken", "expired"))
    net.onTokenRefresh(json({ access_token: "should-not-be-asked", expires_in: 3600 }))

    const { text } = await callAs(server, "bearer-a", "get-task-lists")

    expect(text).toBe("Failed to retrieve task lists")
    expect(net.calls.map((call) => `${call.method} ${call.url}`)).toEqual([
      "GET https://graph.microsoft.com/v1.0/me/todo/lists",
    ])
    expect(JSON.stringify(net.calls)).not.toMatch(/host-/)
  })

  it("refuses a call with no sign-in before any request to Graph", async () => {
    const { status } = await callAs(server, undefined, "get-task-lists")

    expect(status).toBe(401)
    expect(net.calls).toEqual([])
  })

  it("gives no token to a call without a session", async () => {
    const auth = await sessionCredentials.graphAuth({ session: undefined } as Parameters<
      typeof sessionCredentials.graphAuth
    >[0])

    expect(auth.isNone()).toBe(true)
  })

  it("auth-status reports the caller's sign-in, not the host's tokens", async () => {
    expect((await callAs(server, "bearer-a", "auth-status")).text).toBe(
      "Signed in as a@example.com. The Microsoft token expires at 1/1/2100, 12:00:00 AM UTC.",
    )
    expect((await callAs(server, "bearer-b", "auth-status")).text).toBe("Signed in as b@example.com.")
  })
})

describe("createHttpServer with the Azure provider", () => {
  let store: string
  let http: ReturnType<typeof createHttpServer>
  const config = (): HttpConfig =>
    readHttpConfig({ ...validEnv, MSTODO_HTTP_STORE: store }).orThrow(new Error("expected a config"))

  beforeEach(() => {
    store = mkdtempSync(join(tmpdir(), "mstodo-http-store-"))
    http = createHttpServer(config())
  })
  afterEach(() => {
    http.provider.close()
    rmSync(store, { recursive: true, force: true })
  })

  it("challenges an unauthenticated MCP request with the resource metadata URL", async () => {
    const response = await http.server.fetch(
      new Request("https://todo.example.com/mcp", {
        method: "POST",
        headers: { "content-type": "application/json", accept: "application/json, text/event-stream" },
        body: JSON.stringify({ jsonrpc: "2.0", id: 1, method: "tools/list", params: {} }),
      }),
    )

    expect(response.status).toBe(401)
    expect(response.headers.get("www-authenticate")).toContain(
      'resource_metadata="https://todo.example.com/.well-known/oauth-protected-resource',
    )
  })

  it("keeps registered clients in the configured store, encrypted", async () => {
    const response = await http.server.fetch(
      new Request("https://todo.example.com/oauth/register", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ redirect_uris: ["https://claude.ai/cb"], client_name: "spec-client" }),
      }),
    )

    expect(response.status).toBe(201)
    const stored = readdirSync(store, { recursive: true, withFileTypes: true }).filter((entry) => entry.isFile())
    expect(stored.length).toBeGreaterThan(0)
    const contents = stored.map((entry) => readFileSync(join(entry.parentPath, entry.name), "utf8")).join("\n")
    expect(contents).not.toContain("spec-client")
  })

  it("refuses to register a redirect URI off the list", async () => {
    const response = await http.server.fetch(
      new Request("https://todo.example.com/oauth/register", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ redirect_uris: ["https://evil.example/cb"], client_name: "evil" }),
      }),
    )

    expect(response.status).toBe(400)
  })

  it("publishes OAuth metadata with the configured base URL and Graph scopes", async () => {
    const response = await http.server.fetch(
      new Request("https://todo.example.com/.well-known/oauth-authorization-server"),
    )
    const metadata = await response.json()

    expect(metadata.issuer).toBe("https://todo.example.com")
    expect(metadata.scopes_supported).toEqual(expect.arrayContaining(["openid", "offline_access", "Tasks.ReadWrite"]))
  })
})

const SRC = dirname(fileURLToPath(import.meta.url))

describe("what HTTP mode imports", () => {
  // Every module the HTTP entry point can load. None of them may reach the token manager,
  // which reads the host's tokens.
  const reachable = (entry: string): Set<string> => {
    const seen = new Set<string>()
    const visit = (file: string) => {
      if (seen.has(file)) return
      seen.add(file)
      const source = readFileSync(file, "utf8")
      // from "./x", import "./x" and import("./x")
      for (const [, spec] of source.matchAll(/(?:from|import)\s*\(?\s*"(\.{1,2}\/[^"]+)"/g)) {
        visit(resolve(dirname(file), spec.replace(/\.js$/, ".ts")))
      }
    }
    visit(resolve(SRC, entry))
    return seen
  }

  it("never reaches the token manager or the local credentials", () => {
    const files = [...reachable("http.ts")].map((file) => file.slice(SRC.length + 1))

    expect(files).toContain("todo-server.ts")
    expect(files).toContain("session-credentials.ts")
    expect(files).not.toContain("token-manager.ts")
    expect(files).not.toContain("local-credentials.ts")
  })
})
