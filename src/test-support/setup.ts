// Runs in every test worker before any spec file is imported (vitest `setupFiles`).
//
// Two guards that hold even for a spec that forgets the harness, or for code that captures
// `fetch` or reads HOME at import time:
// - The real network is unreachable. Specs install a fake `fetch` per test; between and
//   outside those, `fetch` throws instead of reaching Graph or the token endpoint.
// - HOME, APPDATA and the working directory point at a throwaway directory, so nothing can
//   read the repo's ./tokens.json or write the real ~/.config.
import { mkdtempSync, rmSync } from "fs"
import { tmpdir } from "os"
import { join } from "path"
import { afterAll } from "vitest"

const home = mkdtempSync(join(tmpdir(), "mstodo-spec-home-"))
process.env.HOME = home
process.env.USERPROFILE = home
process.env.APPDATA = join(home, "AppData")
process.env.MSTODO_TOKEN_FILE = join(home, "tokens.json")
for (const key of ["MS_TODO_ACCESS_TOKEN", "MS_TODO_REFRESH_TOKEN", "CLIENT_ID", "CLIENT_SECRET", "TENANT_ID"]) {
  delete process.env[key]
}
process.chdir(home)
afterAll(() => rmSync(home, { recursive: true, force: true }))

globalThis.fetch = (async (input: string | URL | Request) => {
  const url = typeof input === "string" ? input : input instanceof URL ? input.href : input.url
  throw new Error(`real network blocked in specs: ${url}`)
}) as typeof fetch
