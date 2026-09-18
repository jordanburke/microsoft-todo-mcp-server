import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

// Keep the token manager off the real filesystem: a successful refresh calls
// saveTokens() (writes tokens.json) and updateClaudeConfig() (may rewrite the
// Claude desktop config). Mocking fs makes those no-ops — existsSync -> false
// also short-circuits the Claude-config write — so the test can never clobber a
// developer's real tokens.json or config.
vi.mock("fs", () => ({
  existsSync: vi.fn(() => false),
  mkdirSync: vi.fn(),
  readFileSync: vi.fn(() => "{}"),
  writeFileSync: vi.fn(),
}))

import { TokenManager } from "./token-manager.js"

const ENV_KEYS = ["MS_TODO_ACCESS_TOKEN", "MS_TODO_REFRESH_TOKEN", "CLIENT_ID", "CLIENT_SECRET", "TENANT_ID"] as const
const saved: Record<string, string | undefined> = {}

function okTokenResponse() {
  return {
    ok: true,
    json: async () => ({ access_token: "fresh-access", refresh_token: "fresh-refresh", expires_in: 3600 }),
    text: async () => "",
  } as unknown as Response
}

beforeEach(() => {
  for (const k of ENV_KEYS) saved[k] = process.env[k]
  process.env.MS_TODO_ACCESS_TOKEN = "seed-access"
  process.env.MS_TODO_REFRESH_TOKEN = "seed-refresh"
  process.env.CLIENT_ID = "client-id"
  process.env.CLIENT_SECRET = "client-secret"
  process.env.TENANT_ID = "consumers"
})

afterEach(() => {
  for (const k of ENV_KEYS) {
    if (saved[k] === undefined) delete process.env[k]
    else process.env[k] = saved[k]
  }
  vi.restoreAllMocks()
  vi.unstubAllGlobals()
})

describe("TokenManager env-seed cold start", () => {
  it("refreshes the seed once on cold start, then serves the in-process cache", async () => {
    const fetchMock = vi.fn(async () => okTokenResponse())
    vi.stubGlobal("fetch", fetchMock)

    const tm = new TokenManager()
    const first = await tm.getTokens()
    expect(first?.accessToken).toBe("fresh-access")
    expect(fetchMock).toHaveBeenCalledTimes(1)

    // A second call within the cached expiry must not refresh again.
    const second = await tm.getTokens()
    expect(second?.accessToken).toBe("fresh-access")
    expect(fetchMock).toHaveBeenCalledTimes(1)
  })

  it("requests fully-qualified Microsoft Graph scopes (fixes MSA IDX14100)", async () => {
    let body = new URLSearchParams()
    const fetchMock = vi.fn(async (_url: string, init: RequestInit) => {
      body = new URLSearchParams(String(init?.body))
      return okTokenResponse()
    })
    vi.stubGlobal("fetch", fetchMock)

    await new TokenManager().getTokens()

    const scope = body.get("scope") ?? ""
    expect(scope).toContain("https://graph.microsoft.com/Tasks.ReadWrite")
    expect(scope.split(" ")).not.toContain("Tasks.ReadWrite") // no bare scope names
  })

  it("returns the seed token without any network call when client creds are absent", async () => {
    delete process.env.CLIENT_ID
    delete process.env.CLIENT_SECRET
    const fetchMock = vi.fn(async () => okTokenResponse())
    vi.stubGlobal("fetch", fetchMock)

    const tokens = await new TokenManager().getTokens()
    expect(tokens?.accessToken).toBe("seed-access")
    expect(fetchMock).not.toHaveBeenCalled()
  })
})
