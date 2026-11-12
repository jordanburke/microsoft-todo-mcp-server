import { mkdirSync, mkdtempSync, rmSync, symlinkSync, writeFileSync } from "fs"
import { tmpdir } from "os"
import { join } from "path"
import { pathToFileURL } from "url"
import { afterEach, beforeEach, describe, expect, it } from "vitest"

import { isMain } from "./is-main.js"

describe("isMain", () => {
  let dir: string
  let script: string

  beforeEach(() => {
    // A space in the path, as in many home folders, and a symlinked bin, as npm installs one.
    dir = mkdtempSync(join(tmpdir(), "mstodo is-main "))
    mkdirSync(join(dir, "bin"))
    script = join(dir, "create-mcp-config.js")
    writeFileSync(script, "")
  })
  afterEach(() => rmSync(dir, { recursive: true, force: true }))

  it("is true for the script itself, even with a space in its path", () => {
    expect(isMain(pathToFileURL(script).href, script)).toBe(true)
  })

  it("is true when node was started through a symlink to it", () => {
    const link = join(dir, "bin", "mstodo-config")
    symlinkSync(script, link)

    expect(isMain(pathToFileURL(script).href, link)).toBe(true)
  })

  it("is false for another script, and when there is no entry", () => {
    const other = join(dir, "other.js")
    writeFileSync(other, "")

    expect(isMain(pathToFileURL(script).href, other)).toBe(false)
    expect(isMain(pathToFileURL(script).href, undefined)).toBe(false)
  })
})
