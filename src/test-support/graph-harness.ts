// Shared harness for specs that drive the MCP tools in-process.
//
// Every spec talks to the server only through its JSON-RPC surface (`server.fetch`) and
// to Microsoft only through a stubbed `globalThis.fetch`. Nothing in between is mocked,
// so the same specs hold whatever the Graph client and token manager are built from.
//
// Safety: setup.ts (a vitest setupFile) already blocks the real network and moves HOME.
// `startServer` additionally gives each server its own temp directory for the working
// directory, the token file and HOME/APPDATA, before the server module (and its
// token-manager singleton) is imported.
import { mkdtempSync, rmSync, writeFileSync } from "fs"
import { tmpdir } from "os"
import { join } from "path"
import { expect, vi } from "vitest"

export const GRAPH = "https://graph.microsoft.com/v1.0"
export const LOGIN = "https://login.microsoftonline.com"
const TOKEN_ENDPOINT = /^\/[^/]+\/oauth2\/v2\.0\/token$/

export type Call = {
  readonly method: string
  /** The URL exactly as passed to fetch. */
  readonly url: string
  /** Host-relative path plus decoded query (`/me/todo/lists?$filter=status eq 'completed'`), for Graph and login URLs. */
  readonly path: string
  readonly host: string
  readonly headers: Readonly<Record<string, string>>
  readonly body?: unknown
  readonly rawBody?: string
  /** True when the body was a URLSearchParams, which fetch sends as a form. */
  readonly formBody: boolean
}

/** A reply: a Response (cloned per use), or a function of the call that builds one (and may throw). */
export type Reply = Response | ((call: Call) => Response)
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

/**
 * A URL's path and query with the query decoded, so `%24filter=status+eq+%27x%27`,
 * `$filter=status%20eq%20'x'` and `$filter=status eq 'x'` all read the same. The order of
 * parameters is kept. Paths are relative to the Graph base, or to the host otherwise.
 */
export const normalizePath = (url: string): string => {
  const parsed = new URL(url)
  const base = new URL(GRAPH)
  const path =
    parsed.host === base.host && parsed.pathname.startsWith(base.pathname)
      ? parsed.pathname.slice(base.pathname.length)
      : parsed.pathname
  const query = [...parsed.searchParams].map(([k, v]) => `${k}=${v}`).join("&")
  return decodeURIComponent(path) + (query ? `?${query}` : "")
}

const capture = async (input: string | URL | Request, init?: RequestInit): Promise<Call> => {
  const request = input instanceof Request ? input : undefined
  const url = typeof input === "string" ? input : input instanceof URL ? input.href : input.url
  const headers: Record<string, string> = {}
  new Headers(init?.headers ?? request?.headers).forEach((value, key) => (headers[key] = value))
  const rawInit = init?.body ?? undefined
  const formBody = rawInit instanceof URLSearchParams
  const rawBody =
    rawInit !== undefined && rawInit !== null
      ? String(rawInit)
      : request
        ? (await request.clone().text()) || undefined
        : undefined
  const body =
    rawBody === undefined
      ? undefined
      : (() => {
          try {
            return JSON.parse(rawBody)
          } catch {
            return rawBody
          }
        })()
  return {
    method: (init?.method ?? request?.method ?? "GET").toUpperCase(),
    url,
    path: normalizePath(url),
    host: new URL(url).host,
    headers,
    body,
    rawBody,
    formBody,
  }
}

/**
 * A fake network. Routes are tried newest first; the first one that returns a reply wins.
 * A request no route answers is recorded in `unrouted` and answered with a 599, and
 * `restoreNetwork` fails the test if any were recorded, unless `allowUnrouted` is set.
 */
export class FakeNetwork {
  calls: Call[] = []
  unrouted: Call[] = []
  allowUnrouted = false
  private routes: Route[] = []

  reset(): void {
    this.calls = []
    this.unrouted = []
    this.allowUnrouted = false
    this.routes = []
  }

  /** Answer Graph `method` requests whose normalized path equals or matches `path`. */
  on(method: string, path: string | RegExp, reply: Reply): this {
    const graphHost = new URL(GRAPH).host
    this.routes.unshift((call) => {
      if (call.method !== method || call.host !== graphHost) return undefined
      return (typeof path === "string" ? call.path === normalizePath(GRAPH + path) : path.test(call.path))
        ? reply
        : undefined
    })
    return this
  }

  /** Answer the token endpoint (any tenant). */
  onTokenRefresh(reply: Reply): this {
    const loginHost = new URL(LOGIN).host
    this.routes.unshift((call) =>
      call.method === "POST" && call.host === loginHost && TOKEN_ENDPOINT.test(call.path) ? reply : undefined,
    )
    return this
  }

  route(route: Route): this {
    this.routes.unshift(route)
    return this
  }

  /** Requests sent to Graph, as "METHOD normalized-path". */
  graphRequests(): string[] {
    const graphHost = new URL(GRAPH).host
    return this.calls.filter((c) => c.host === graphHost).map((c) => `${c.method} ${c.path}`)
  }

  install(): void {
    vi.stubGlobal(
      "fetch",
      vi.fn(async (input: string | URL | Request, init?: RequestInit) => {
        const call = await capture(input, init)
        this.calls.push(call)
        for (const route of this.routes) {
          const reply = route(call)
          if (reply !== undefined) return (typeof reply === "function" ? reply(call) : reply).clone()
        }
        this.unrouted.push(call)
        return json({ error: { code: "unrouted", message: `no fake route for ${call.method} ${call.url}` } }, 599)
      }),
    )
  }
}

export type TokenSetup =
  | { kind: "env"; accessToken?: string; refreshToken?: string }
  | { kind: "file"; tokens: Record<string, unknown> }
  | { kind: "none" }

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
 * its own token manager. Start it after faking the clock, if a spec fakes it: an env token's
 * expiry is read when the server first sees it.
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
  if (tokens.kind === "file") {
    writeFileSync(join(dir, "tokens.json"), JSON.stringify(tokens.tokens))
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

/**
 * Per-test teardown. Call in afterEach. Fails the test if a request reached no route,
 * because a mistyped route otherwise looks like a Graph failure. Restores the blocking
 * fetch that setup.ts installed.
 */
export function restoreNetwork(net?: FakeNetwork): void {
  const unrouted = net && !net.allowUnrouted ? net.unrouted.map((c) => `${c.method} ${c.url}`) : []
  vi.unstubAllGlobals()
  vi.restoreAllMocks()
  expect(unrouted, "requests no fake route answered").toEqual([])
}
