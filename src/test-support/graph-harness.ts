// Shared harness for specs that drive the MCP tools in-process.
//
// Every spec talks to the server only through its JSON-RPC surface (`server.fetch`) and
// to Microsoft only through a stubbed `globalThis.fetch`. Nothing in between is mocked,
// so the same specs hold whatever the Graph client and token manager are built from.
//
// Safety: the repo's ./tokens.json is a real credential, and a successful token refresh
// rewrites Claude Desktop's config under $HOME. `startServer` therefore moves the working
// directory, the token file and HOME/APPDATA into a fresh temp directory before the server
// module (and its token-manager singleton) is imported.
import { mkdtempSync, rmSync } from "fs"
import { tmpdir } from "os"
import { join } from "path"
import { vi } from "vitest"

export const GRAPH = "https://graph.microsoft.com/v1.0"
export const LOGIN = "https://login.microsoftonline.com"

export type Call = {
  readonly method: string
  readonly url: string
  readonly headers: Readonly<Record<string, string>>
  readonly body?: unknown
  readonly rawBody?: string
}

export type Reply = Response | (() => Response)
export type Route = (call: Call) => Reply | undefined

export type ToolResult = {
  readonly text: string
  readonly isError: boolean
  readonly content: ReadonlyArray<{ type: string; text?: string }>
}

// Response helpers. Graph sends JSON with a content-type for data and errors, and an
// empty body with no content-type for 204 No Content (every DELETE).
export const json = (body: unknown, status = 200): Response => Response.json(body, { status })
export const noContent = (): Response => new Response(null, { status: 204 })
export const graphError = (status: number, code: string, message: string): Response =>
  json({ error: { code, message } }, status)

export const jwtExpiringIn = (seconds: number): string =>
  `header.${Buffer.from(JSON.stringify({ exp: Math.floor(Date.now() / 1000) + seconds })).toString("base64url")}.sig`

const headersOf = (init?: RequestInit): Record<string, string> => {
  const out: Record<string, string> = {}
  new Headers(init?.headers).forEach((value, key) => (out[key] = value))
  return out
}

const bodyOf = (init?: RequestInit): { body?: unknown; rawBody?: string } => {
  if (init?.body === undefined || init.body === null) return {}
  const rawBody = String(init.body)
  try {
    return { body: JSON.parse(rawBody), rawBody }
  } catch {
    return { body: rawBody, rawBody }
  }
}

/**
 * A fake network. Routes are tried newest first; the first one that returns a reply wins.
 * A request no route answers is recorded and answered with a 599, so a missing route
 * shows up in the assertions instead of hanging or reaching the real Graph.
 */
export class FakeNetwork {
  calls: Call[] = []
  unrouted: Call[] = []
  private routes: Route[] = []

  reset(): void {
    this.calls = []
    this.unrouted = []
    this.routes = []
  }

  /** Answer `method` requests whose URL (without the Graph base) equals or matches `path`. */
  on(method: string, path: string | RegExp, reply: Reply | ((call: Call) => Reply)): this {
    this.routes.unshift((call) => {
      if (call.method !== method) return undefined
      const rel = call.url.startsWith(GRAPH) ? call.url.slice(GRAPH.length) : call.url
      const hit = typeof path === "string" ? rel === path : path.test(rel)
      if (!hit) return undefined
      return typeof reply === "function" && reply.length > 0 ? (reply as (c: Call) => Reply)(call) : (reply as Reply)
    })
    return this
  }

  /** Answer the token endpoint (any tenant). */
  onTokenRefresh(reply: Reply | ((call: Call) => Reply)): this {
    return this.on("POST", new RegExp(`^${LOGIN.replace(/\./g, "\\.")}/[^/]+/oauth2/v2\\.0/token$`), reply)
  }

  route(route: Route): this {
    this.routes.unshift(route)
    return this
  }

  /** Requests sent to Graph, as "METHOD url-without-base". */
  graphRequests(): string[] {
    return this.calls.filter((c) => c.url.startsWith(GRAPH)).map((c) => `${c.method} ${c.url.slice(GRAPH.length)}`)
  }

  install(): void {
    vi.stubGlobal(
      "fetch",
      vi.fn(async (input: string | URL | Request, init?: RequestInit) => {
        const url = typeof input === "string" ? input : input instanceof URL ? input.href : input.url
        const call: Call = {
          method: (init?.method ?? "GET").toUpperCase(),
          url,
          headers: headersOf(init),
          ...bodyOf(init),
        }
        this.calls.push(call)
        for (const route of this.routes) {
          const reply = route(call)
          if (reply !== undefined) return typeof reply === "function" ? reply() : reply.clone()
        }
        this.unrouted.push(call)
        return json({ error: { code: "unrouted", message: `no fake route for ${call.method} ${url}` } }, 599)
      }),
    )
  }
}

export type TokenSetup = { kind: "env"; accessToken?: string; refreshToken?: string } | { kind: "none" }

export type Harness = {
  readonly dir: string
  readonly net: FakeNetwork
  callTool(name: string, args?: Record<string, unknown>): Promise<ToolResult>
  rpc(
    method: string,
    params?: Record<string, unknown>,
  ): Promise<{ result?: unknown; error?: { code: number; message: string } }>
  cleanup(): void
}

const ENV_KEYS = [
  "HOME",
  "APPDATA",
  "MSTODO_TOKEN_FILE",
  "MS_TODO_ACCESS_TOKEN",
  "MS_TODO_REFRESH_TOKEN",
  "CLIENT_ID",
  "CLIENT_SECRET",
  "TENANT_ID",
] as const

/**
 * Isolate the process, then import a fresh copy of the server. Call it in beforeAll when the
 * tests share token state, or in beforeEach (with cleanup in afterEach) when each test needs
 * its own token manager.
 */
export async function startServer(tokens: TokenSetup = { kind: "env" }): Promise<Harness> {
  const saved = Object.fromEntries(ENV_KEYS.map((k) => [k, process.env[k]]))
  const originalCwd = process.cwd()
  const dir = mkdtempSync(join(tmpdir(), "mstodo-spec-"))

  process.chdir(dir)
  process.env.HOME = dir
  process.env.APPDATA = join(dir, "AppData")
  process.env.MSTODO_TOKEN_FILE = join(dir, "tokens.json")
  process.env.CLIENT_ID = "spec-client-id"
  process.env.CLIENT_SECRET = "spec-client-secret"
  delete process.env.TENANT_ID
  delete process.env.MS_TODO_ACCESS_TOKEN
  delete process.env.MS_TODO_REFRESH_TOKEN
  if (tokens.kind === "env") {
    process.env.MS_TODO_ACCESS_TOKEN = tokens.accessToken ?? jwtExpiringIn(3600)
    process.env.MS_TODO_REFRESH_TOKEN = tokens.refreshToken ?? "spec-refresh"
  }

  const net = new FakeNetwork()
  // A fresh module graph each time, so a spec can start from clean token-manager state
  // by calling startServer in beforeEach rather than beforeAll.
  vi.resetModules()
  const { server } = await import("../todo-index.js")
  let rpcId = 0

  const rpc: Harness["rpc"] = async (method, params) => {
    const response = await server.fetch(
      new Request("http://localhost/mcp", {
        method: "POST",
        headers: { "content-type": "application/json", accept: "application/json, text/event-stream" },
        body: JSON.stringify({ jsonrpc: "2.0", id: ++rpcId, method, params: params ?? {} }),
      }),
    )
    const raw = await response.text()
    const data = raw.split("\n").find((line) => line.startsWith("data: "))
    return JSON.parse(data ? data.slice("data: ".length) : raw)
  }

  return {
    dir,
    net,
    rpc,
    async callTool(name, args = {}) {
      const message = await rpc("tools/call", { name, arguments: args })
      if (message.error) throw new Error(`JSON-RPC error ${message.error.code}: ${message.error.message}`)
      const result = message.result as { content: ToolResult["content"]; isError?: boolean }
      return {
        text: result.content.map((c) => c.text ?? "").join("\n"),
        isError: result.isError === true,
        content: result.content,
      }
    },
    cleanup() {
      process.chdir(originalCwd)
      rmSync(dir, { recursive: true, force: true })
      for (const k of ENV_KEYS) {
        if (saved[k] === undefined) delete process.env[k]
        else process.env[k] = saved[k]
      }
    },
  }
}

/** Per-test setup: fresh fake network, silenced server logs. Call in beforeEach. */
export function resetNetwork(net: FakeNetwork): void {
  net.reset()
  net.install()
  vi.spyOn(console, "error").mockImplementation(() => {})
}

/** Per-test teardown. Call in afterEach. */
export function restoreNetwork(): void {
  vi.unstubAllGlobals()
  vi.restoreAllMocks()
}
