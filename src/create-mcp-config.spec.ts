import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "fs"
import { tmpdir } from "os"
import { join, resolve } from "path"
import { afterEach, beforeEach, describe, expect, it } from "vitest"

import { createMcpConfig, mcpConfigFor, tokenFileWarnings } from "./create-mcp-config.js"

describe("mcpConfigFor", () => {
  it("points the server at the token file by absolute path, with no tokens in the config", () => {
    const config = mcpConfigFor("some/dir/tokens.json")

    expect(config).toEqual({
      mcpServers: {
        microsoftTodo: {
          command: "npx",
          args: ["--yes", "microsoft-todo-mcp-server"],
          env: { MSTODO_TOKEN_FILE: resolve("some/dir/tokens.json") },
        },
      },
    })
  })
})

describe("tokenFileWarnings", () => {
  it("is empty for a file the server can renew on its own", () => {
    expect(tokenFileWarnings({ refreshToken: "r", clientId: "c", clientSecret: "s" })).toEqual([])
  })

  it("names what is missing", () => {
    expect(tokenFileWarnings({ accessToken: "a" })).toHaveLength(2)
    expect(tokenFileWarnings({ refreshToken: "r" })).toEqual([
      "has no clientId/clientSecret, so renewal also needs CLIENT_ID and CLIENT_SECRET in the server's env",
    ])
  })
})

describe("createMcpConfig", () => {
  let dir: string

  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), "mstodo-config-"))
  })
  afterEach(() => rmSync(dir, { recursive: true, force: true }))

  it("writes the config without copying any token into it", () => {
    const tokens = join(dir, "tokens.json")
    const out = join(dir, "mcp.json")
    writeFileSync(
      tokens,
      JSON.stringify({
        accessToken: "secret-access",
        refreshToken: "secret-refresh",
        clientId: "c",
        clientSecret: "s",
      }),
    )

    expect(createMcpConfig(tokens, out)).toEqual([])

    const written = readFileSync(out, "utf8")
    expect(JSON.parse(written)).toEqual(mcpConfigFor(tokens))
    expect(written).not.toContain("secret")
  })
})
