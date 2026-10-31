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

### Running the Server

```bash
pnpm run cli         # Run MCP server via CLI wrapper
pnpm start           # Run MCP server directly
```

### Live Smoke (before/after comparison)

```bash
pnpm run build && node scripts/smoke-live.mjs --label <name>   # writes .smoke/<name>.json
node scripts/smoke-live.mjs --compare <a> <b>                  # diff two snapshots; exit 1 if any differ
```

Drives every tool over stdio against the real account in `tokens.json`. Writes stay inside two
temporary "MCP smoke … (safe to delete)" lists, deleted at the end; existing lists are read for
shape only (counts, which fields print). `--cmd "<command>"` runs a different server entry point.
`.smoke/` is gitignored because snapshots hold account data. Two runs of the same build compare
clean, so any diff is a behaviour change (or, for steps marked real data, the account changed).
The baseline for the somamcp migration is `.smoke/v1.2.0-before-somamcp.json`, recorded on
v1.2.0 (98f79e9). Graph rejects `$orderby=title` on tasks, so the harness orders by `createdDateTime`.

## ts-builds & pnpm 12 Notes

Built with **ts-builds 4.0.2** (tsdown, vitest 5) on **pnpm 12**. Non-obvious, load-bearing constraints:

- **Node >= 22.12 required.** ts-builds 4 bundles vitest 5, whose `engines.node` is `^22.12.0 || ^24.0.0 || >=26.0.0`. CI runs Node 22.x/24.x; the publish workflow reads `.nvmrc`, pinned to the exact patch `24.21.0` so a non-retryable release job does not vary with the runner's tool cache. **pnpm, not npm, performs the OIDC token exchange** for trusted publishing — it never shells out to the npm CLI — so `packageManager` (pnpm 12.16.1) is the version that gates publishing, not the bundled npm. Proven under pnpm 12.16.1: v1.1.4 published on 2026-10-28, the repo's first CI release, signing a provenance statement over OIDC with no npm token anywhere.
- **Keep `globals` in `publicHoistPattern`** (`pnpm-workspace.yaml`). `eslint.config.js` imports `globals`, which `*eslint*` does not match. Without the explicit hoist line, resolution falls back to pnpm's incidental `.pnpm` virtual-store hoist — passes locally but **fails CI** with `Cannot find package 'globals'`. It is load-bearing, not redundant; ts-builds `init` only adds it for new projects, and `doctor` does not retro-fit this file.
- **First-party release-age excludes are globs, not pins.** `minimumReleaseAgeExclude` is `*functype*` + `ts-builds`. `pnpm add` auto-adds per-version pins for first-party packages within the 24h cooldown — collapse them back to the glob (pins re-trip on every release).
- **`allowBuilds: esbuild: false`** is required under pnpm's `strictDepBuilds` (esbuild's binary ships via `@esbuild/<platform>` optional deps; build script not needed).
- **pnpm 12 locks itself.** `pnpm-lock.yaml` opens with a `packageManagerDependencies` document pinning `pnpm` and `@pnpm/exe` to the `packageManager` version. Bumping `packageManager` changes that header; it is not a dependency change. pnpm 12 also fails installs on an unknown `pnpm-workspace.yaml` setting.
- **tsdown emits `.js`, not `.mjs`.** `tsdown.config.ts` forces `outExtensions: () => ({ js: ".js" })` so the `bin`/`main` paths and `./*.js` imports resolve. Don't drop it.
- **`pnpm install` in a non-TTY shell** may abort with `ERR_PNPM_ABORTED_REMOVE_MODULES_DIR_NO_TTY` — prefix with `CI=true`. Note `CI=true` makes `--frozen-lockfile` the default, so add `--no-frozen-lockfile` when intentionally updating the lockfile.

## Architecture Overview

This is a Model Context Protocol (MCP) server that enables AI assistants to interact with Microsoft To Do via the Microsoft Graph API. The codebase follows a modular architecture with four main components:

1. **MCP Server** (`src/todo-index.ts`): Core server implementing the MCP protocol with 13 tools for Microsoft To Do operations
2. **CLI Wrapper** (`src/cli.ts`): Executable entry point that handles token loading from environment or file
3. **Auth Server** (`src/auth-server.js`): Express server implementing OAuth 2.0 flow with MSAL
4. **Config Generator** (`src/create-mcp-config.ts`): Utility to create MCP configuration files

### Key Architectural Patterns

- **Token Management**: Tokens are stored in `tokens.json` with automatic refresh 5 minutes before expiration
- **Multi-tenant Support**: Configurable for different Microsoft account types via TENANT_ID
- **Error Handling**: Special handling for personal Microsoft accounts (MailboxNotEnabledForRESTAPI)
- **Type Safety**: Strict TypeScript with Zod schemas for parameter validation

### Microsoft Graph API Integration

The server communicates with Microsoft Graph API v1.0:

- Base URL: `https://graph.microsoft.com/v1.0`
- Three-level hierarchy: Lists → Tasks → Checklist Items
- Supports OData query parameters for filtering and sorting

### Environment Configuration

- `MSTODO_TOKEN_FILE`: Custom path for tokens.json (defaults to ./tokens.json)
- `.env` file required for authentication with CLIENT_ID, CLIENT_SECRET, TENANT_ID, REDIRECT_URI

## Important Notes

- Always run `pnpm run build` (or `pnpm run validate`) after modifying TypeScript files (ts-builds/tsdown bundling)
- The auth server runs on port 3000 by default
- Tokens are automatically refreshed using the refresh token when needed
- Personal Microsoft accounts have limited API access compared to work/school accounts
