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
pnpm start           # Run MCP server (dist/cli.js)
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

Built with **ts-builds 4.5.0** (tsdown, vitest 5) on **pnpm 12**. Non-obvious, load-bearing constraints:

- **Node 22.18+ or 24.11+ required to build.** tsdown 0.23 declares `engines.node` `^22.18.0 || ^24.11.0 || >=26.0.0`, the tightest floor in the toolchain; ts-builds 4 and its vitest 5 need only `^22.12.0 || ^24.0.0 || >=26.0.0`. The published server itself runs on Node 20+. CI runs Node 22.x/24.x; the publish workflow reads `.nvmrc`, pinned to the exact patch `24.21.0` so a non-retryable release job does not vary with the runner's tool cache. **pnpm, not npm, performs the OIDC token exchange** for trusted publishing — it never shells out to the npm CLI — so `packageManager` (pnpm 12.16.1) is the version that gates publishing, not the bundled npm. Proven under pnpm 12.16.1: v1.1.4 published on 2026-10-28, the repo's first CI release, signing a provenance statement over OIDC with no npm token anywhere.
- **Keep `globals` in `publicHoistPattern`** (`pnpm-workspace.yaml`). `eslint.config.js` imports `globals`, which `*eslint*` does not match. Without the explicit hoist line, resolution falls back to pnpm's incidental `.pnpm` virtual-store hoist — passes locally but **fails CI** with `Cannot find package 'globals'`. It is load-bearing, not redundant; ts-builds `init` only adds it for new projects, and `doctor` does not retro-fit this file.
- **First-party release-age excludes are globs, not pins.** `minimumReleaseAgeExclude` is `*functype*`, `somamcp` and `ts-builds`. `pnpm add` auto-adds per-version pins for first-party packages within the 24h cooldown — collapse them back to the glob (pins re-trip on every release).
- **`allowBuilds: esbuild: false`** is required under pnpm's `strictDepBuilds` (esbuild's binary ships via `@esbuild/<platform>` optional deps; build script not needed).
- **pnpm 12 locks itself.** `pnpm-lock.yaml` opens with a `packageManagerDependencies` document pinning `pnpm` and `@pnpm/exe` to the `packageManager` version. Bumping `packageManager` changes that header; it is not a dependency change. pnpm 12 also fails installs on an unknown `pnpm-workspace.yaml` setting.
- **tsdown emits `.js`, not `.mjs`.** `tsdown.config.ts` forces `outExtensions: () => ({ js: ".js" })` so the `bin`/`main` paths and `./*.js` imports resolve. Don't drop it.
- **`pnpm install` in a non-TTY shell** may abort with `ERR_PNPM_ABORTED_REMOVE_MODULES_DIR_NO_TTY` — prefix with `CI=true`. Note `CI=true` makes `--frozen-lockfile` the default, so add `--no-frozen-lockfile` to `pnpm install` when intentionally updating the lockfile. `pnpm add` rejects that flag (`Unknown option: 'frozen-lockfile'`) and updates the lockfile anyway, so run plain `CI=true pnpm add <pkg>`.

## Architecture Overview

This is a Model Context Protocol (MCP) server that enables AI assistants to interact with Microsoft To Do via the Microsoft Graph API. The codebase follows a modular architecture with these main components:

1. **MCP Server** (`src/todo-server.ts`): `createTodoServer({ credentials, auth? })` builds a somamcp 2.10 server (MCP SDK 2.x underneath) with 17 tools: 16 for Microsoft To Do plus somamcp's `info`. Each tool call gets its Graph token from `credentials` (`src/credentials.ts`: a `CredentialSource`), never from the token manager directly
2. **stdio entry** (`src/todo-index.ts`): `server` is `createTodoServer` with `localCredentials` (`src/local-credentials.ts`), the token manager's tokens; `startServer()` serves it over stdio. Only `local-credentials.ts` imports the token manager
3. **Graph Client** (`src/graph-client.ts`): Graph requests as functype `IO` over `Http.client`, with a typed error channel (`GraphError` = `HttpError | MailboxNotEnabled | PageLimit`). A request takes the call's `GraphAuth`: its token, and `renew` for a 401
4. **HTTP mode** (`src/http.ts`, bin `mstodo-http`): `src/http-config.ts` reads and validates the env, refusing to start on any problem (so somamcp never falls back to random keys); `redirectPatternProblem` requires a literal scheme and host in each redirect pattern (somamcp 2.7+ also refuses unsafe patterns when the provider is built; this check runs first, with the setting's name, and is a little stricter. `http-mode.spec.ts` pins that every sample it accepts also builds a provider); `src/http-server.ts` builds somamcp's Azure provider from `azureProviderConfig` (a pure function the spec pins field by field; PKCE `plain` off) (`createAzureProvider`, `createDiskStore`, an optional `MSTODO_HTTP_ALLOWED_USERS` allow-list as `acceptSignIn`) and `createTodoServer` with `sessionCredentials` (`src/session-credentials.ts`): each call's own `session.accessToken`, no renewal (the client's refresh renews Microsoft's token), no fallback. `http-mode.spec.ts` walks the import graph from `http.ts` and fails if it reaches `token-manager.ts` or `local-credentials.ts`. `http-sign-in.spec.ts` runs a whole sign-in through the real provider with only Microsoft faked: register, consent, callback, redemption, tool calls, refresh, rotation, the allow-list, consent at a single-slash `/oauth/consent`, `plain` PKCE refused, and a code that survives a wrong verifier
5. **CLI Wrapper** (`src/cli.ts`): Executable entry point; it only calls `startServer()`, and `src/token-manager.ts` finds and refreshes the tokens
6. **Auth Server** (`src/auth-server.ts`): Express server implementing OAuth 2.0 flow with MSAL (msal-node 7). It reads the refresh token out of MSAL's serialized cache (`RefreshToken.<key>.secret`), so check that shape after an MSAL major bump
7. **Config Generator** (`src/create-mcp-config.ts`): writes an `mcp.json` whose entry sets `MSTODO_TOKEN_FILE` to the token file; it copies no tokens

### somamcp Notes

- **Tools** register with `server.addTool({ name, description, parameters: z.object({...}), execute })`. Handlers return `{ content: [...] }`. Parameters must be Zod 4, which somamcp needs to describe them as JSON Schema.
- **`serverInfo.version`** is read from `package.json` at startup (one level up from `dist/` and `src/`), so it tracks releases.
- **`info`** is somamcp's introspection tool, on by default (`enableIntrospection`). It is the 17th tool.
- **Specs drive tools through `server.fetch()`** with JSON-RPC (`src/recurrence.spec.ts`). The backend is stateless, so no `initialize` is needed, and each reply is a single server-sent event (`data: {...}`).
- **SDK 2.x error shapes:** an unknown tool is a JSON-RPC `-32602` error; invalid arguments are `isError` results. Messages name the bad field (`listId: Invalid input: ...`) only from somamcp 2.4.0 (SomaMCP #70). somamcp 2.0–2.3 gave only the issue text; the pre-somamcp server (SDK 1.x) did name the field.
- **Fix somamcp at the source.** It is Jordan's library (`~/IdeaProjects/SomaMCP`, session `somamcp-*`); report or fix bugs there rather than shimming here.

### Key Architectural Patterns

- **functype at the edges**: a tool builds its request as an `IO`, runs it once at the MCP boundary with `.run()`, and folds the `Either` into text (`graphErrorText` renders a failure). No module-level error state. A 401 renews the token and retries once with the same body; paging follows `@odata.nextLink` to the end or the limit, capped at 100 pages. DELETE is read raw (Graph's 204 has no body); every other reply must be JSON
- **Token Management**: Tokens are stored in `tokens.json` and renewed 5 minutes before expiry, and again whenever Graph returns 401 (`refreshRejectedToken`). `TokenManager` keeps a Promise API but holds tokens as `Option`, reads files with `Try`, renews through `Http`, and joins a renewal already under way for the same refresh token. A renewal from the env pair records the pair's fingerprint (`seededFrom`, a truncated SHA-256 of the refresh token); a later launch with that same pair, or with a file holding the same refresh token (as older versions left them), uses the token file instead. The token manager never writes to Claude Desktop's config
- **Multi-tenant Support**: Configurable for different Microsoft account types via TENANT_ID
- **Error Handling**: `MailboxNotEnabledForRESTAPI` becomes its own error (`MailboxNotEnabled`) with an explanation; nothing warns personal accounts up front, because they work
- **Type Safety**: Strict TypeScript with Zod 4 schemas for parameter validation
- **Tests** live beside the code as `src/*.spec.ts`. The repo's `./tokens.json` is a real credential, and the token manager writes under `$HOME`, so specs `chdir` into a temp directory and point `MSTODO_TOKEN_FILE`, `HOME` and `APPDATA` there before any `TokenManager` reads tokens. `src/test-support/graph-harness.ts` does this, then imports a fresh `todo-index.ts` (whose import builds the token-manager singleton) and drives tools through `server.fetch()`
- **`src/test-support/setup.ts`** (a vitest `setupFiles`) runs before any spec file loads: it makes `fetch` throw, and points `HOME`, `APPDATA`, the token file and the cwd at a throwaway directory. That holds even for code that captures `fetch` or reads `HOME` at import time
- **Specs touch only the boundaries**: JSON-RPC in, a stubbed `globalThis.fetch` out (`FakeNetwork`: routes per method and path, separate token-endpoint routes). A request no route answers fails the test in `restoreNetwork(h.net)`, unless the spec sets `net.allowUnrouted`. Nothing in between is mocked, so the specs hold across a rewrite of the Graph client or token manager. Paths are compared with the query decoded, so `%24top=3` and `$top=3` match. Mocks match Graph's wire shapes: `noContent()` for a DELETE's 204 with no body, `graphError()` for its error JSON. Vitest pins `TZ=UTC` and an `en_US` locale and uses the forks pool (specs `chdir`). Date-dependent specs fake only `Date`, and start the server under the fake clock
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

- Env-token setups (`MS_TODO_ACCESS_TOKEN`/`MS_TODO_REFRESH_TOKEN` with no matching token file, such as an `mcp.json` from an older `mstodo-config`) cannot renew without `CLIENT_ID`/`CLIENT_SECRET` (and `TENANT_ID` unless it is `organizations`) in the env too: on the env path no stored tokens exist yet to supply them.

## Important Notes

- Always run `pnpm run build` (or `pnpm run validate`) after modifying TypeScript files (ts-builds/tsdown bundling)
- The auth server runs on port 3000 by default
- Tokens are automatically refreshed using the refresh token when needed
- Personal Microsoft accounts work against the To Do API (verified with a live.com account on 2026-10-31); `MailboxNotEnabledForRESTAPI` means Microsoft has not enabled the API for that mailbox
- **HTTP mode's real sign-in was verified on 2026-11-12** with a live.com account, `TENANT_ID=common`, on `http://localhost:3100` (the app registration has `http://localhost:3100/oauth/callback` for this). Microsoft granted `openid profile email User.Read Tasks.Read Tasks.ReadWrite Tasks.Read.Shared Tasks.ReadWrite.Shared`: the short spellings `HTTP_SCOPES` asks for, with `offline_access` left out of the report though a refresh token was issued. The id token carried `preferred_username`, so `upn` and an email allow-list work on personal accounts. Tool calls read and wrote as that user; a refresh renewed Microsoft's token, and the spent refresh token was refused (400). To repeat it: run `dist/http.js` with `MSTODO_HTTP_BASE_URL=http://localhost:3100` and fresh keys, then drive an OAuth client (register, PKCE authorize, loopback callback, redeem) and open the authorize link in a browser on the same machine
