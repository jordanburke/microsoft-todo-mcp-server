import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "fs"
import { tmpdir } from "os"
import { join } from "path"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import { TokenManager } from "./token-manager.js"

// A Graph access token is a JWT, so its real expiry is readable from the `exp` claim.
// Tests build one rather than relying on a fixed string, so they do not rot.
const jwtExpiringIn = (seconds: number): string => {
  const payload = Buffer.from(JSON.stringify({ exp: Math.floor(Date.now() / 1000) + seconds })).toString("base64url")
  return `header.${payload}.signature`
}

let dir: string
let tokenFile: string

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), "mstodo-tokens-"))
  tokenFile = join(dir, "tokens.json")
  process.env.MSTODO_TOKEN_FILE = tokenFile
  delete process.env.MS_TODO_ACCESS_TOKEN
  delete process.env.MS_TODO_REFRESH_TOKEN
  process.env.CLIENT_ID = "client-id"
  process.env.CLIENT_SECRET = "client-secret"
})

afterEach(() => {
  rmSync(dir, { recursive: true, force: true })
  vi.unstubAllGlobals()
  delete process.env.MSTODO_TOKEN_FILE
  delete process.env.MS_TODO_ACCESS_TOKEN
  delete process.env.MS_TODO_REFRESH_TOKEN
  delete process.env.CLIENT_ID
  delete process.env.CLIENT_SECRET
})

const stubRefresh = (accessToken: string) => {
  const fetchMock = vi.fn(async () =>
    Response.json({ access_token: accessToken, refresh_token: "refreshed-refresh", expires_in: 3600 }),
  )
  vi.stubGlobal("fetch", fetchMock)
  return fetchMock
}

describe("MSTODO_TOKEN_FILE", () => {
  it("reads tokens from the path the variable names", async () => {
    writeFileSync(
      tokenFile,
      JSON.stringify({ accessToken: "from-custom-path", refreshToken: "r", expiresAt: Date.now() + 3600_000 }),
    )

    const tokens = await new TokenManager().getTokens()

    expect(tokens?.accessToken).toBe("from-custom-path")
  })

  it("writes refreshed tokens back to that same path", async () => {
    writeFileSync(tokenFile, JSON.stringify({ accessToken: "stale", refreshToken: "r", expiresAt: 0 }))
    stubRefresh("renewed")

    await new TokenManager().getTokens()

    expect(JSON.parse(readFileSync(tokenFile, "utf8")).accessToken).toBe("renewed")
  })
})

describe("environment token pair", () => {
  it("does not treat an expired access token as valid", async () => {
    process.env.MS_TODO_ACCESS_TOKEN = jwtExpiringIn(-60)
    process.env.MS_TODO_REFRESH_TOKEN = "env-refresh"
    const fetchMock = stubRefresh("renewed-from-env")

    const tokens = await new TokenManager().getTokens()

    expect(fetchMock).toHaveBeenCalledOnce()
    expect(tokens?.accessToken).toBe("renewed-from-env")
  })

  it("keeps reporting an expired token as expired on later calls", async () => {
    // The deadline used to be re-stamped to now+1h on every call, so it always sat an
    // hour ahead and the refresh branch could never be reached.
    process.env.MS_TODO_ACCESS_TOKEN = jwtExpiringIn(-60)
    process.env.MS_TODO_REFRESH_TOKEN = "env-refresh"
    const fetchMock = vi.fn(async () => Response.json({ error: "invalid_grant" }, { status: 400 }))
    vi.stubGlobal("fetch", fetchMock)

    const manager = new TokenManager()
    await manager.getTokens()
    await manager.getTokens()

    expect(fetchMock).toHaveBeenCalledTimes(2)
  })

  it("uses an unexpired token without refreshing", async () => {
    process.env.MS_TODO_ACCESS_TOKEN = jwtExpiringIn(3600)
    process.env.MS_TODO_REFRESH_TOKEN = "env-refresh"
    const fetchMock = stubRefresh("should-not-be-used")

    const tokens = await new TokenManager().getTokens()

    expect(fetchMock).not.toHaveBeenCalled()
    expect(tokens?.accessToken).toBe(process.env.MS_TODO_ACCESS_TOKEN)
  })

  it("prefers a token refreshed in this process over the stale environment pair", async () => {
    process.env.MS_TODO_ACCESS_TOKEN = jwtExpiringIn(-60)
    process.env.MS_TODO_REFRESH_TOKEN = "env-refresh"
    const fetchMock = stubRefresh("renewed-once")

    const manager = new TokenManager()
    await manager.getTokens()
    const second = await manager.getTokens()

    expect(fetchMock).toHaveBeenCalledOnce()
    expect(second?.accessToken).toBe("renewed-once")
  })
})
