# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Common Development Commands

### Build and Development

```bash
pnpm install         # Install dependencies (pnpm 12)
pnpm run build       # Build with ts-builds (tsdown) to dist/ directory
pnpm run dev         # Build and run CLI in one command
pnpm run validate    # Full chain: format, lint, typecheck, test, build
```

Build tooling is provided by **ts-builds** (4.x); package scripts delegate to the
`ts-builds` CLI. Output extension is forced to `.js` via `tsdown.config.ts` so the
`bin`/`main` paths and `./*.js` source imports keep resolving.

### Authentication and Setup

```bash
pnpm run auth        # Start OAuth authentication server (port 3000)
pnpm run create-config # Generate mcp.json from tokens.json
```

`pnpm run auth` reads `.env` once at startup and writes `tokens.json` (with `clientId`,
`clientSecret` and `tenantId`, which the token manager needs to renew) to the working
directory. It never exits on its own. That is also why `mstodo-setup` does not finish: its
close handler, which moves the tokens and edits Claude Desktop's config, only runs when the
auth server is stopped, and then reports failure.

### Running the Server

```bash
pnpm run cli         # Run MCP server via CLI wrapper
pnpm start           # Run MCP server directly
```

### Live Smoke (before/after comparison)

```bash
pnpm run build && node scripts/smoke-live.mjs --label <name>   # writes .smoke/<name>.json
node scripts/smoke-live.mjs --compare <a> <b>                  # diff two snapshots
```

Drives every tool over stdio against the real account in `tokens.json`, and records the
`initialize` result, full tool schemas, each call's normalized result (text, `isError`, content
keys, `structuredContent`, JSON-RPC error codes) and stdout hygiene. Writes stay inside two
temporary "MCP smoke … (safe to delete)" lists. Cleanup runs even if the server dies (it starts a
fresh one), also deletes smoke lists left by an earlier run, and trusts only a fresh
`get-task-lists` as proof. Existing lists are read for shape only. `--cmd "<command>"` runs a
different server entry point. `.smoke/` is gitignored because snapshots hold account data.

Two runs of the same build compare clean. `--compare` exits 1 only on a behaviour difference;
steps marked real data also change when the account does, so they are reported but do not fail.
A run exits 1 if the server died early or cleanup was not verified.

Snapshots for the somamcp migration (#35): `.smoke/v1.2.0-before-somamcp.json` and
`v1.2.0-repeat.json` were recorded against a v1.2.0 build in its own worktree (`--cmd "node
<worktree>/dist/cli.js"`), and `after-somamcp.json` against somamcp 2.4.0. To compare an older
build, build it in a worktree and pass `--cmd`, rather than checking out old code here. The
harness's input dates are in 2031 on purpose: the normalizer masks today's date, so an input
date that falls on the day of a run reads as `<today>` and differs from every other day's
baseline. Graph rejects `$orderby=title` on tasks, so the harness orders by `createdDateTime`.

## ts-builds & pnpm 12 Notes

Built with **ts-builds 4.0.2** (tsdown, vitest 5) on **pnpm 12**. Non-obvious, load-bearing constraints:

- **Node 22.18+ or 24.11+ required to build.** tsdown 0.22 declares `engines.node` `^22.18.0 || >=24.11.0`, the tightest floor in the toolchain; ts-builds 4 and its vitest 5 need only `^22.12.0 || ^24.0.0 || >=26.0.0`. The published server itself runs on Node 20+. CI runs Node 22.x/24.x; the publish workflow reads `.nvmrc`, pinned to the exact patch `24.21.0` so a non-retryable release job does not vary with the runner's tool cache. **pnpm, not npm, performs the OIDC token exchange** for trusted publishing — it never shells out to the npm CLI — so `packageManager` (pnpm 12.16.1) is the version that gates publishing, not the bundled npm. Proven under pnpm 12.16.1: v1.1.4 published on 2026-10-28, the repo's first CI release, signing a provenance statement over OIDC with no npm token anywhere.
- **Keep `globals` in `publicHoistPattern`** (`pnpm-workspace.yaml`). `eslint.config.js` imports `globals`, which `*eslint*` does not match. Without the explicit hoist line, resolution falls back to pnpm's incidental `.pnpm` virtual-store hoist — passes locally but **fails CI** with `Cannot find package 'globals'`. It is load-bearing, not redundant; ts-builds `init` only adds it for new projects, and `doctor` does not retro-fit this file.
- **First-party release-age excludes are globs, not pins.** `minimumReleaseAgeExclude` is `*functype*`, `somamcp` and `ts-builds`. `pnpm add` auto-adds per-version pins for first-party packages within the 24h cooldown — collapse them back to the glob (pins re-trip on every release).
- **`allowBuilds: esbuild: false`** is required under pnpm's `strictDepBuilds` (esbuild's binary ships via `@esbuild/<platform>` optional deps; build script not needed).
- **pnpm 12 locks itself.** `pnpm-lock.yaml` opens with a `packageManagerDependencies` document pinning `pnpm` and `@pnpm/exe` to the `packageManager` version. Bumping `packageManager` changes that header; it is not a dependency change. pnpm 12 also fails installs on an unknown `pnpm-workspace.yaml` setting.
- **tsdown emits `.js`, not `.mjs`.** `tsdown.config.ts` forces `outExtensions: () => ({ js: ".js" })` so the `bin`/`main` paths and `./*.js` imports resolve. Don't drop it.
- **`pnpm install` in a non-TTY shell** may abort with `ERR_PNPM_ABORTED_REMOVE_MODULES_DIR_NO_TTY` — prefix with `CI=true`. Note `CI=true` makes `--frozen-lockfile` the default, so add `--no-frozen-lockfile` to `pnpm install` when intentionally updating the lockfile. `pnpm add` rejects that flag (`Unknown option: 'frozen-lockfile'`) and updates the lockfile anyway, so run plain `CI=true pnpm add <pkg>`.

## Architecture Overview

This is a Model Context Protocol (MCP) server that enables AI assistants to interact with Microsoft To Do via the Microsoft Graph API. The codebase follows a modular architecture with four main components:

1. **MCP Server** (`src/todo-index.ts`): Built on somamcp (MCP SDK 2.x underneath), serving 17 tools over stdio: 16 for Microsoft To Do plus somamcp's `info`
2. **CLI Wrapper** (`src/cli.ts`): Executable entry point; it only calls `startServer()`, and `src/token-manager.ts` finds and refreshes the tokens
3. **Auth Server** (`src/auth-server.ts`): Express server implementing OAuth 2.0 flow with MSAL
4. **Config Generator** (`src/create-mcp-config.ts`): Utility to create MCP configuration files

### somamcp Notes

- **Tools** register with `server.addTool({ name, description, parameters: z.object({...}), execute })`. Handlers return `{ content: [...] }`. Parameters must be Zod 4, which somamcp needs to describe them as JSON Schema.
- **`serverInfo.version`** is read from `package.json` at startup (one level up from `dist/` and `src/`), so it tracks releases.
- **`info`** is somamcp's introspection tool, on by default (`enableIntrospection`). It is the 17th tool.
- **Specs drive tools through `server.fetch()`** with JSON-RPC (`src/recurrence.spec.ts`). The backend is stateless, so no `initialize` is needed, and each reply is a single server-sent event (`data: {...}`).
- **SDK 2.x error shapes:** an unknown tool is a JSON-RPC `-32602` error; invalid arguments are `isError` results. Messages name the bad field (`listId: Invalid input: ...`) only from somamcp 2.4.0 (SomaMCP #70). somamcp 2.0–2.3 gave only the issue text; the pre-somamcp server (SDK 1.x) did name the field.
- **Fix somamcp at the source.** It is Jordan's library (`~/IdeaProjects/SomaMCP`, session `somamcp-*`); report or fix bugs there rather than shimming here.

### Key Architectural Patterns

- **Token Management**: Tokens are stored in `tokens.json` and renewed 5 minutes before expiry, and again whenever Graph returns 401 (`refreshRejectedToken`)
- **Multi-tenant Support**: Configurable for different Microsoft account types via TENANT_ID
- **Error Handling**: Special handling for personal Microsoft accounts (MailboxNotEnabledForRESTAPI)
- **Type Safety**: Strict TypeScript with Zod 4 schemas for parameter validation
- **Tests** live beside the code as `src/*.spec.ts`. The repo's `./tokens.json` is a real credential, and a successful refresh rewrites Claude Desktop's config under `$HOME`, so specs `chdir` into a temp directory and point `MSTODO_TOKEN_FILE`, `HOME` and `APPDATA` there before any `TokenManager` reads tokens. `src/test-support/graph-harness.ts` does this, then imports a fresh `todo-index.ts` (whose import builds the token-manager singleton) and drives tools through `server.fetch()`
- **Specs touch only the boundaries**: JSON-RPC in, a stubbed `globalThis.fetch` out (`FakeNetwork`: routes per method and path, a 599 for anything unrouted, separate token-endpoint routes). Nothing in between is mocked, so the specs hold across a rewrite of the Graph client or token manager. Mocks match Graph's wire shapes: `noContent()` for a DELETE's 204 with no body, `graphError()` for its error JSON, exact request URLs (`get-tasks` URL-encodes its OData parameters; `archive-completed-tasks` sends its `$filter` raw). Vitest pins `TZ=UTC`, and date-dependent specs fake only `Date`
- **`KNOWN ISSUE:` tests pin today's wrong behaviour** from the list below. A fix flips its test in the same commit

### Microsoft Graph API Integration

The server communicates with Microsoft Graph API v1.0:

- Base URL: `https://graph.microsoft.com/v1.0`
- Three-level hierarchy: Lists → Tasks → Checklist Items
- Supports OData query parameters for filtering and sorting

### Environment Configuration

- `MSTODO_TOKEN_FILE`: Custom path for tokens.json. Without it, the token manager uses the per-user config directory (`~/.config/microsoft-todo-mcp/tokens.json`, or `%APPDATA%\microsoft-todo-mcp\tokens.json` on Windows), then falls back to `./tokens.json` and migrates it
- `.env` file required for authentication with CLIENT_ID, CLIENT_SECRET, TENANT_ID, REDIRECT_URI

## Releasing

Publishing runs in `.github/workflows/publish.yml` when a GitHub release is published (OIDC trusted publishing, with provenance):

```bash
npm version <patch|minor|major> && git push --follow-tags origin main
gh release create vX.Y.Z --title vX.Y.Z --generate-notes --verify-tag
```

`npm version` also syncs `server.json` through the `version` script. The same release event runs `publish-registry.yml`, which waits for the npm tarball and then submits `server.json` to the MCP Registry; if it fails, rerun it with `workflow_dispatch` rather than cutting a new release. After publishing, wait for the tarball to answer 200 before installing: npm's CDN caches a 404 from an early request for several minutes (seen on v1.1.4 and v1.2.0).

## Known Issues (not yet fixed)

- `delete-task` and `delete-task-list` report success without checking Graph's response, so an unknown ID "succeeds".
- `get-tasks` offers `orderby=title`, which Graph rejects, and `count`, which Graph ignores (the server prints the number returned).
- `get-task-lists-organized` with `groupBy=shared` prints no list IDs, even with `includeIds`.
- Env-token setups (`MS_TODO_ACCESS_TOKEN`/`MS_TODO_REFRESH_TOKEN`, as `create-mcp-config.ts` writes) cannot renew without `CLIENT_ID`/`CLIENT_SECRET` (and `TENANT_ID` unless it is `organizations`) in the env too: on the env path no stored tokens exist yet to supply them.
- After a renewal, `TokenManager.updateClaudeConfig` writes the new tokens into a Claude Desktop entry named `microsoft-todo`. On the next launch those env tokens take precedence over the token file, and they carry no client credentials (see the previous item).
- The auth server's `/refresh` route rewrites `tokens.json` without the refresh token or client credentials.
- `postinstall` prints `npx microsoft-todo-mcp-server setup`, which starts the server, not setup; the setup command is `mstodo-setup`.
- The personal-account `MailboxNotEnabledForRESTAPI` warning in `auth-status` and the sign-in page is shown to every personal account, though a live.com account works fully.

## Important Notes

- Always run `pnpm run build` (or `pnpm run validate`) after modifying TypeScript files (ts-builds/tsdown bundling)
- The auth server runs on port 3000 by default
- Tokens are automatically refreshed using the refresh token when needed
- Personal Microsoft accounts work against the To Do API (verified with a live.com account on 2026-10-31); `MailboxNotEnabledForRESTAPI` means Microsoft has not enabled the API for that mailbox
