#!/usr/bin/env node
// Keeps server.json in step with package.json.
//
// The MCP Registry rejects a submission whose server.json version does not name
// a version that exists on npm, and it verifies ownership by matching
// package.json `mcpName` against server.json `name`. Both are easy to forget on
// a release, and the failure only shows up in the publish job.
//
// `npm version` bumps package.json alone, so server.json is generated from it
// rather than hand-edited.
//
//   node scripts/sync-server-json.mjs           write server.json
//   node scripts/sync-server-json.mjs --check   fail if it is stale
//
// Runs automatically via the npm `version` lifecycle, and --check runs in
// prepublishOnly so a hand-edit cannot reach npm or the registry.

import { readFile, writeFile } from "node:fs/promises"

const TARGET = "server.json"

const pkg = JSON.parse(await readFile("package.json", "utf8"))
const server = JSON.parse(await readFile(TARGET, "utf8"))

if (server.name !== pkg.mcpName) {
  console.error(`✗ package.json mcpName "${pkg.mcpName}" does not match ${TARGET} name "${server.name}".`)
  console.error("  The registry uses that match to verify npm package ownership.")
  process.exit(1)
}

const synced = {
  ...server,
  version: pkg.version,
  packages: server.packages.map((p) => (p.registryType === "npm" ? { ...p, version: pkg.version } : p)),
}

const contents = `${JSON.stringify(synced, null, 2)}\n`
const check = process.argv.includes("--check")

if (check) {
  const current = await readFile(TARGET, "utf8")
  if (current !== contents) {
    console.error(`✗ ${TARGET} is out of sync with package.json (${pkg.version}).`)
    console.error("  Run: pnpm sync:server-json")
    process.exit(1)
  }
  console.log(`✓ ${TARGET} matches package.json (${pkg.version})`)
} else {
  await writeFile(TARGET, contents)
  console.log(`✓ wrote ${TARGET} (${pkg.version})`)
}
