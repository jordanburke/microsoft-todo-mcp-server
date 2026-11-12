import { chmodSync, existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "fs"
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
let originalCwd: string
let originalHome: string | undefined
let originalAppData: string | undefined

beforeEach(() => {
  vi.useFakeTimers({ shouldAdvanceTime: true })
  dir = mkdtempSync(join(tmpdir(), "mstodo-tokens-"))
  tokenFile = join(dir, "custom.json")
  // getTokens() falls back to ./tokens.json, which in this repo is a real credential.
  originalCwd = process.cwd()
  process.chdir(dir)
  // A successful refresh rewrites Claude Desktop's config under $HOME; keep it in the temp dir.
  originalHome = process.env.HOME
  originalAppData = process.env.APPDATA
  process.env.HOME = dir
  process.env.APPDATA = join(dir, "AppData")
  delete process.env.TENANT_ID
  process.env.MSTODO_TOKEN_FILE = tokenFile
  delete process.env.MS_TODO_ACCESS_TOKEN
  delete process.env.MS_TODO_REFRESH_TOKEN
  process.env.CLIENT_ID = "client-id"
  process.env.CLIENT_SECRET = "client-secret"
})

afterEach(() => {
  vi.useRealTimers()
  process.chdir(originalCwd)
  if (originalHome === undefined) delete process.env.HOME
  else process.env.HOME = originalHome
  if (originalAppData === undefined) delete process.env.APPDATA
  else process.env.APPDATA = originalAppData
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

// Which refresh_token a given call sent, read back out of the posted form body.
const refreshTokenSentOn = (fetchMock: ReturnType<typeof vi.fn>, call: number): string =>
  new URLSearchParams(String(fetchMock.mock.calls[call][1].body)).get("refresh_token") ?? ""

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

  it("stamps an opaque token's expiry once instead of on every call", async () => {
    // Personal Microsoft accounts get an opaque access token, so there is no `exp` to
    // read and the one-hour assumption is all there is. It must not move with the clock.
    process.env.MS_TODO_ACCESS_TOKEN = "opaque-not-a-jwt"
    process.env.MS_TODO_REFRESH_TOKEN = "env-refresh"
    const fetchMock = stubRefresh("renewed-from-opaque")
    const manager = new TokenManager()

    const first = await manager.getTokens()
    vi.setSystemTime(Date.now() + 61 * 60 * 1000)
    const second = await manager.getTokens()

    expect(first?.accessToken).toBe("opaque-not-a-jwt")
    expect(fetchMock).toHaveBeenCalledOnce()
    expect(second?.accessToken).toBe("renewed-from-opaque")
  })

  it("refreshes with the rotated token, not the original environment one", async () => {
    // Entra rotates the refresh token on every use. Re-sending the environment value
    // would pin the chain to that token's own lifetime and drop each replacement.
    process.env.MS_TODO_ACCESS_TOKEN = jwtExpiringIn(-60)
    process.env.MS_TODO_REFRESH_TOKEN = "env-refresh"
    const fetchMock = stubRefresh("renewed")
    const manager = new TokenManager()

    await manager.getTokens()
    vi.setSystemTime(Date.now() + 61 * 60 * 1000)
    await manager.getTokens()

    expect(refreshTokenSentOn(fetchMock, 0)).toBe("env-refresh")
    expect(refreshTokenSentOn(fetchMock, 1)).toBe("refreshed-refresh")
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

describe("refreshRejectedToken", () => {
  // An opaque seed minted on an earlier day sits inside its assumed hour, so getTokens()
  // hands it back unrefreshed. Only Graph's 401 shows it is dead.
  beforeEach(() => {
    process.env.MS_TODO_ACCESS_TOKEN = "opaque-stale-seed"
    process.env.MS_TODO_REFRESH_TOKEN = "env-refresh"
  })

  it("refreshes a token Graph rejected even though it still looks valid", async () => {
    const fetchMock = stubRefresh("renewed-after-401")
    const manager = new TokenManager()

    expect((await manager.getTokens())?.accessToken).toBe("opaque-stale-seed")
    expect(fetchMock).not.toHaveBeenCalled()

    const refreshed = await manager.refreshRejectedToken("opaque-stale-seed")

    expect(refreshed?.accessToken).toBe("renewed-after-401")
    expect(refreshTokenSentOn(fetchMock, 0)).toBe("env-refresh")
    expect((await manager.getTokens())?.accessToken).toBe("renewed-after-401")
  })

  it("does not refresh again when another request already replaced the rejected token", async () => {
    const fetchMock = stubRefresh("renewed-after-401")
    const manager = new TokenManager()
    await manager.refreshRejectedToken("opaque-stale-seed")

    const second = await manager.refreshRejectedToken("opaque-stale-seed")

    expect(fetchMock).toHaveBeenCalledOnce()
    expect(second?.accessToken).toBe("renewed-after-401")
  })

  it("returns null when the refresh fails", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn(async () => Response.json({ error: "invalid_grant" }, { status: 400 })),
    )

    expect(await new TokenManager().refreshRejectedToken("opaque-stale-seed")).toBeNull()
  })
})

const writeTokenFile = (tokens: Record<string, unknown>, path = tokenFile) =>
  writeFileSync(path, JSON.stringify(tokens))
const readTokenFile = () => JSON.parse(readFileSync(tokenFile, "utf8"))
const tokenUrl = (fetchMock: ReturnType<typeof vi.fn>, call = 0): string => String(fetchMock.mock.calls[call][0])

describe("token file", () => {
  it("returns a live token without refreshing", async () => {
    writeTokenFile({ accessToken: "live", refreshToken: "r", expiresAt: Date.now() + 60_000 })
    const fetchMock = stubRefresh("unused")

    expect((await new TokenManager().getTokens())?.accessToken).toBe("live")
    expect(fetchMock).not.toHaveBeenCalled()
  })

  it("hands back the expired token when renewal fails", async () => {
    writeTokenFile({ accessToken: "stale", refreshToken: "r", expiresAt: 0 })
    vi.stubGlobal(
      "fetch",
      vi.fn(async () => Response.json({ error: "invalid_grant" }, { status: 400 })),
    )

    expect((await new TokenManager().getTokens())?.accessToken).toBe("stale")
    expect(readTokenFile().accessToken).toBe("stale")
  })

  it("renews with the client credentials and tenant stored in the file, ahead of the environment", async () => {
    writeTokenFile({
      accessToken: "stale",
      refreshToken: "file-refresh",
      expiresAt: 0,
      clientId: "file-client",
      clientSecret: "file-secret",
      tenantId: "contoso-tenant",
    })
    const fetchMock = stubRefresh("renewed")

    await new TokenManager().getTokens()

    expect(tokenUrl(fetchMock)).toBe("https://login.microsoftonline.com/contoso-tenant/oauth2/v2.0/token")
    const form = new URLSearchParams(String(fetchMock.mock.calls[0][1].body))
    expect([form.get("client_id"), form.get("client_secret"), form.get("refresh_token")]).toEqual([
      "file-client",
      "file-secret",
      "file-refresh",
    ])
    expect(readTokenFile()).toMatchObject({
      accessToken: "renewed",
      refreshToken: "refreshed-refresh",
      clientId: "file-client",
      clientSecret: "file-secret",
      tenantId: "contoso-tenant",
    })
  })

  it("falls back to TENANT_ID, then to the organizations tenant", async () => {
    writeTokenFile({ accessToken: "stale", refreshToken: "r", expiresAt: 0 })
    process.env.TENANT_ID = "env-tenant"
    const fetchMock = stubRefresh("renewed")
    await new TokenManager().getTokens()
    delete process.env.TENANT_ID

    writeTokenFile({ accessToken: "stale", refreshToken: "r", expiresAt: 0 })
    await new TokenManager().getTokens()

    expect(tokenUrl(fetchMock, 0)).toContain("/env-tenant/")
    expect(tokenUrl(fetchMock, 1)).toContain("/organizations/")
  })

  it("keeps the old refresh token when the identity provider does not rotate it", async () => {
    writeTokenFile({ accessToken: "stale", refreshToken: "keep-me", expiresAt: 0 })
    vi.stubGlobal(
      "fetch",
      vi.fn(async () => Response.json({ access_token: "renewed", expires_in: 3600 })),
    )

    await new TokenManager().getTokens()

    expect(readTokenFile().refreshToken).toBe("keep-me")
  })

  it("stamps a renewed token to expire five minutes before expires_in", async () => {
    vi.setSystemTime(new Date("2026-11-11T12:00:00Z"))
    writeTokenFile({ accessToken: "stale", refreshToken: "r", expiresAt: 0 })
    stubRefresh("renewed")

    await new TokenManager().getTokens()

    expect(readTokenFile().expiresAt).toBe(new Date("2026-11-11T12:55:00Z").getTime())
  })

  it("does not call the token endpoint without client credentials", async () => {
    delete process.env.CLIENT_ID
    writeTokenFile({ accessToken: "stale", refreshToken: "r", expiresAt: 0 })
    const fetchMock = stubRefresh("unused")

    expect((await new TokenManager().getTokens())?.accessToken).toBe("stale")
    expect(fetchMock).not.toHaveBeenCalled()
  })

  it("treats an unreadable file as no tokens", async () => {
    writeFileSync(tokenFile, "{not json")

    expect(await new TokenManager().getTokens()).toBeNull()
  })

  it("renews once when two callers find the token expired at the same moment", async () => {
    writeTokenFile({ accessToken: "stale", refreshToken: "r", expiresAt: 0 })
    const fetchMock = stubRefresh("renewed")
    const manager = new TokenManager()

    const [first, second] = await Promise.all([manager.getTokens(), manager.getTokens()])

    expect(fetchMock).toHaveBeenCalledOnce()
    expect([first?.accessToken, second?.accessToken]).toEqual(["renewed", "renewed"])
  })

  it("starts a fresh renewal once the previous one has finished", async () => {
    writeTokenFile({ accessToken: "stale", refreshToken: "r", expiresAt: 0 })
    vi.stubGlobal(
      "fetch",
      vi.fn(async () => Response.json({ error: "invalid_grant" }, { status: 400 })),
    )
    const manager = new TokenManager()

    await manager.getTokens()
    await manager.getTokens()

    expect(vi.mocked(fetch)).toHaveBeenCalledTimes(2)
  })
})

describe("legacy ./tokens.json", () => {
  it("is read and copied to the token file when that file does not exist", async () => {
    writeTokenFile(
      { accessToken: "legacy", refreshToken: "r", expiresAt: Date.now() + 60_000 },
      join(dir, "tokens.json"),
    )

    const tokens = await new TokenManager().getTokens()

    expect(tokens?.accessToken).toBe("legacy")
    expect(readTokenFile().accessToken).toBe("legacy")
  })

  it("is ignored once the token file exists", async () => {
    writeTokenFile(
      { accessToken: "legacy", refreshToken: "r", expiresAt: Date.now() + 60_000 },
      join(dir, "tokens.json"),
    )
    writeTokenFile({ accessToken: "current", refreshToken: "r", expiresAt: Date.now() + 60_000 })

    expect((await new TokenManager().getTokens())?.accessToken).toBe("current")
  })

  it("is returned as found, even when expired", async () => {
    writeTokenFile({ accessToken: "legacy-stale", refreshToken: "r", expiresAt: 0 }, join(dir, "tokens.json"))
    const fetchMock = stubRefresh("unused")

    expect((await new TokenManager().getTokens())?.accessToken).toBe("legacy-stale")
    expect(fetchMock).not.toHaveBeenCalled()
  })
})

describe("no tokens anywhere", () => {
  it("returns null", async () => {
    expect(await new TokenManager().getTokens()).toBeNull()
  })
})

describe("environment pair and token file across restarts", () => {
  it("records which environment pair a renewal descends from, as a fingerprint", async () => {
    process.env.MS_TODO_ACCESS_TOKEN = jwtExpiringIn(-60)
    process.env.MS_TODO_REFRESH_TOKEN = "env-refresh"
    stubRefresh("renewed")

    await new TokenManager().getTokens()

    const saved = readTokenFile()
    expect(saved.seededFrom).toMatch(/^[0-9a-f]{16}$/)
    expect(JSON.stringify(saved)).not.toContain("env-refresh")
  })

  it("picks up the token file on the next launch when it was renewed from the same pair", async () => {
    process.env.MS_TODO_ACCESS_TOKEN = jwtExpiringIn(-60)
    process.env.MS_TODO_REFRESH_TOKEN = "env-refresh"
    const fetchMock = stubRefresh("renewed")
    await new TokenManager().getTokens()

    const relaunched = await new TokenManager().getTokens()

    expect(relaunched?.accessToken).toBe("renewed")
    expect(fetchMock).toHaveBeenCalledOnce()
  })

  it("renews from the file's rotated refresh token, not the original pair, once that expires", async () => {
    process.env.MS_TODO_ACCESS_TOKEN = jwtExpiringIn(-60)
    process.env.MS_TODO_REFRESH_TOKEN = "env-refresh"
    const fetchMock = stubRefresh("renewed")
    await new TokenManager().getTokens()
    vi.setSystemTime(Date.now() + 2 * 3600 * 1000)

    await new TokenManager().getTokens()

    expect(refreshTokenSentOn(fetchMock, 1)).toBe("refreshed-refresh")
  })

  it("adopts a file holding the same refresh token, as older versions left them, and renews with its credentials", async () => {
    // An older version renewed, wrote the token file, then copied the same pair into
    // Claude Desktop's config as env tokens, without the client credentials.
    delete process.env.CLIENT_ID
    delete process.env.CLIENT_SECRET
    process.env.MS_TODO_ACCESS_TOKEN = "copied-access"
    process.env.MS_TODO_REFRESH_TOKEN = "copied-refresh"
    writeTokenFile({
      accessToken: "copied-access",
      refreshToken: "copied-refresh",
      expiresAt: 0,
      clientId: "file-client",
      clientSecret: "file-secret",
    })
    const fetchMock = stubRefresh("healed")

    const tokens = await new TokenManager().getTokens()

    expect(tokens?.accessToken).toBe("healed")
    expect(refreshTokenSentOn(fetchMock, 0)).toBe("copied-refresh")
    expect(new URLSearchParams(String(fetchMock.mock.calls[0][1].body)).get("client_id")).toBe("file-client")
  })

  it("starts over from a new environment pair, such as one from a fresh sign-in", async () => {
    process.env.MS_TODO_ACCESS_TOKEN = jwtExpiringIn(-60)
    process.env.MS_TODO_REFRESH_TOKEN = "env-refresh"
    stubRefresh("renewed")
    await new TokenManager().getTokens()

    process.env.MS_TODO_ACCESS_TOKEN = jwtExpiringIn(3600)
    process.env.MS_TODO_REFRESH_TOKEN = "new-sign-in"

    expect((await new TokenManager().getTokens())?.accessToken).toBe(process.env.MS_TODO_ACCESS_TOKEN)
  })
})

describe("environment pair over the token file", () => {
  it("wins even when the file holds a newer token", async () => {
    process.env.MS_TODO_ACCESS_TOKEN = jwtExpiringIn(3600)
    process.env.MS_TODO_REFRESH_TOKEN = "env-refresh"
    writeTokenFile({ accessToken: "file-token", refreshToken: "r", expiresAt: Date.now() + 7200_000 })

    expect((await new TokenManager().getTokens())?.accessToken).toBe(process.env.MS_TODO_ACCESS_TOKEN)
  })
})

describe("storeCredentials", () => {
  it("adds client credentials to the loaded tokens and saves them", async () => {
    writeTokenFile({ accessToken: "live", refreshToken: "r", expiresAt: Date.now() + 60_000 })
    const manager = new TokenManager()
    await manager.getTokens()

    await manager.storeCredentials("cid", "secret", "tenant")

    expect(readTokenFile()).toMatchObject({
      accessToken: "live",
      clientId: "cid",
      clientSecret: "secret",
      tenantId: "tenant",
    })
  })

  it("does nothing before any tokens are loaded", async () => {
    await new TokenManager().storeCredentials("cid", "secret", "tenant")

    expect(existsSync(tokenFile)).toBe(false)
  })
})

describe("Claude Desktop config after a renewal", () => {
  // Where the token manager looks on each platform, under the temp HOME/APPDATA.
  const claudeDir = () =>
    process.platform === "win32"
      ? join(process.env.APPDATA ?? "", "Claude")
      : process.platform === "darwin"
        ? join(dir, "Library", "Application Support", "Claude")
        : join(dir, ".config", "Claude")
  const claudeConfig = () => join(claudeDir(), "claude_desktop_config.json")
  const writeClaudeConfig = (config: unknown) => {
    mkdirSync(claudeDir(), { recursive: true })
    writeFileSync(claudeConfig(), JSON.stringify(config))
  }

  it("is never written, even with a microsoft-todo entry", async () => {
    const config = {
      mcpServers: {
        "microsoft-todo": { command: "npx", env: { KEEP: "1" } },
        other: { command: "x", env: { A: "b" } },
      },
    }
    writeClaudeConfig(config)
    writeTokenFile({ accessToken: "stale", refreshToken: "r", expiresAt: 0 })
    stubRefresh("renewed")

    await new TokenManager().getTokens()

    expect(JSON.parse(readFileSync(claudeConfig(), "utf8"))).toEqual(config)
  })

  it("leaves a config without that entry untouched, and creates none", async () => {
    writeClaudeConfig({ mcpServers: { mstodo: { command: "npx" } } })
    writeTokenFile({ accessToken: "stale", refreshToken: "r", expiresAt: 0 })
    stubRefresh("renewed")

    await new TokenManager().getTokens()

    expect(JSON.parse(readFileSync(claudeConfig(), "utf8"))).toEqual({ mcpServers: { mstodo: { command: "npx" } } })
  })
})

// Root ignores file modes, so these run only where chmod can make a file read-only.
const canMakeReadOnly = process.getuid?.() !== 0

describe.runIf(canMakeReadOnly)("a token file that cannot be written", () => {
  it("still hands back the renewed token, and keeps using it", async () => {
    writeTokenFile({ accessToken: "stale", refreshToken: "r", expiresAt: 0 })
    chmodSync(tokenFile, 0o444)
    const fetchMock = stubRefresh("renewed")
    const manager = new TokenManager()

    const first = await manager.getTokens()
    const second = await manager.refreshRejectedToken("stale")

    expect(first?.accessToken).toBe("renewed")
    expect(second?.accessToken).toBe("renewed")
    expect(fetchMock).toHaveBeenCalledOnce()
    expect(readTokenFile().accessToken).toBe("stale")
  })

  it("keeps using the renewed token instead of renewing again on every call", async () => {
    writeTokenFile({ accessToken: "stale", refreshToken: "r", expiresAt: 0 })
    chmodSync(tokenFile, 0o444)
    const fetchMock = stubRefresh("renewed")
    const manager = new TokenManager()

    await manager.getTokens()
    const again = await manager.getTokens()

    expect(again?.accessToken).toBe("renewed")
    expect(fetchMock).toHaveBeenCalledOnce()
  })

  it("still returns legacy tokens it could not copy", async () => {
    mkdirSync(join(dir, "ro"))
    process.env.MSTODO_TOKEN_FILE = join(dir, "ro", "tokens.json")
    chmodSync(join(dir, "ro"), 0o555)
    writeTokenFile(
      { accessToken: "legacy", refreshToken: "r", expiresAt: Date.now() + 60_000 },
      join(dir, "tokens.json"),
    )

    try {
      expect((await new TokenManager().getTokens())?.accessToken).toBe("legacy")
    } finally {
      chmodSync(join(dir, "ro"), 0o755)
    }
  })
})

describe("overlapping renewals for different refresh tokens", () => {
  it("does not let one finishing clear the other", async () => {
    const releases: Array<() => void> = []
    const fetchMock = vi.fn(
      (_url: string, init: RequestInit) =>
        new Promise<Response>((resolve) => {
          const token = new URLSearchParams(String(init.body)).get("refresh_token")
          releases.push(() =>
            resolve(
              Response.json({ access_token: `renewed-${token}`, refresh_token: `next-${token}`, expires_in: 3600 }),
            ),
          )
        }),
    )
    vi.stubGlobal("fetch", fetchMock)
    const manager = new TokenManager()

    const a = manager.refreshToken("a")
    const b = manager.refreshToken("b")
    await vi.waitFor(() => expect(releases).toHaveLength(2))
    releases[0]()
    await a
    const bAgain = manager.refreshToken("b")
    releases[1]()

    expect((await b)?.accessToken).toBe("renewed-b")
    expect((await bAgain)?.accessToken).toBe("renewed-b")
    expect(fetchMock).toHaveBeenCalledTimes(2)
  })
})
