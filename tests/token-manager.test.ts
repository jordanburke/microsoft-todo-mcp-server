import { beforeEach, describe, expect, it } from "vitest"
import { mkdtempSync, writeFileSync } from "fs"
import { tmpdir } from "os"
import { join } from "path"
import { TokenManager } from "../src/token-manager.js"

beforeEach(() => {
  delete process.env.MS_TODO_ACCESS_TOKEN
  delete process.env.MS_TODO_REFRESH_TOKEN
})

describe("TokenManager", () => {
  it("reads tokens from a custom token file path", async () => {
    const dir = mkdtempSync(join(tmpdir(), "mstodo-tokens-"))
    const file = join(dir, "tokens.json")
    writeFileSync(
      file,
      JSON.stringify({
        accessToken: "access-123",
        refreshToken: "refresh-123",
        expiresAt: Date.now() + 60 * 60 * 1000,
      }),
    )

    const manager = new TokenManager()
    manager.setTokenFilePath(file)

    const tokens = await manager.getTokens()
    expect(tokens?.accessToken).toBe("access-123")
    expect(tokens?.refreshToken).toBe("refresh-123")
  })

  it("prefers environment variables over the token file", async () => {
    const dir = mkdtempSync(join(tmpdir(), "mstodo-tokens-"))
    const file = join(dir, "tokens.json")
    writeFileSync(
      file,
      JSON.stringify({
        accessToken: "from-file",
        refreshToken: "from-file",
        expiresAt: Date.now() + 60 * 60 * 1000,
      }),
    )

    process.env.MS_TODO_ACCESS_TOKEN = "from-env"
    process.env.MS_TODO_REFRESH_TOKEN = "from-env"

    const manager = new TokenManager()
    manager.setTokenFilePath(file)

    const tokens = await manager.getTokens()
    expect(tokens?.accessToken).toBe("from-env")
  })

  it("returns null when no token sources are available", async () => {
    const dir = mkdtempSync(join(tmpdir(), "mstodo-tokens-"))
    const manager = new TokenManager()
    manager.setTokenFilePath(join(dir, "missing.json"))

    const tokens = await manager.getTokens()
    expect(tokens).toBeNull()
  })
})
