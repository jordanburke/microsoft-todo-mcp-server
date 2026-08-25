# AGENTS.md

Guidance for AI coding agents working in this repository.

## Project Overview

Microsoft To Do MCP server — a Model Context Protocol server that lets AI assistants interact with Microsoft To Do through the Microsoft Graph API v1.0. Fork of `@jhirono/todomcp`. Published to npm as `microsoft-todo-mcp-server`.

## Common Development Commands

Package manager: **pnpm 11** (pinned via `packageManager` field). Requires **Node >= 22.13**.

```bash
pnpm install            # Install dependencies (use CI=true in non-TTY shells)
pnpm run build          # Build with ts-builds (tsdown) to dist/
pnpm run dev            # Build and run CLI in one command
```

### Verification (run before finishing any change)

```bash
pnpm test               # Unit tests (vitest, tests/)
pnpm run validate       # Full chain: format, lint, typecheck, test, build
```

Code style is enforced by Prettier and ESLint (config in `.prettierrc` / `eslint.config.js`), both wrapped by the `ts-builds` CLI.

### ts-builds & pnpm 11 Notes

See `CLAUDE.md` ("ts-builds & pnpm 11 Notes") for load-bearing constraints: Node >= 22.13, `publicHoistPattern` globals hoist, tsdown `.js` output extension, `CI=true` for non-TTY installs.

### Authentication and Setup

```bash
pnpm run setup          # Interactive first-time setup wizard (npx entry point)
pnpm run auth           # Start OAuth 2.0 authentication server (port 3000)
pnpm run create-config  # Generate mcp.json from tokens.json
```

### Running the Server

```bash
pnpm run cli            # Run MCP server via CLI wrapper
pnpm start              # Run built server directly (node dist/todo-index.js)
```

## Architecture Overview

All source lives in `src/` (ESM — `"type": "module"`, so relative imports must use `.js` extensions).

| File                       | Role                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| -------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `src/todo-index.ts`        | Core MCP server. Registers **16 tools** via `server.tool()` (15 by default; `test-graph-api-exploration` requires `MSTODO_ENABLE_EXPLORATION=1`): `auth-status`, `get-task-lists`, `get-task-lists-organized`, `create-task-list`, `update-task-list`, `delete-task-list`, `get-tasks`, `create-task`, `update-task`, `delete-task`, `get-checklist-items`, `create-checklist-item`, `update-checklist-item`, `delete-checklist-item`, `archive-completed-tasks`, `test-graph-api-exploration` |
| `src/cli.ts`               | Executable entry point. Resolves tokens from env vars (`MS_TODO_ACCESS_TOKEN` / `MS_TODO_REFRESH_TOKEN`) or token file, then starts the server                                                                                                                                                                                                                                                                                                                                                 |
| `src/token-manager.ts`     | Token storage/refresh. Refreshes tokens automatically before expiry                                                                                                                                                                                                                                                                                                                                                                                                                            |
| `src/auth-server.ts`       | Express + MSAL OAuth 2.0 flow (TypeScript, runs on port 3000)                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| `src/setup.ts`             | Interactive setup wizard (guides Azure app registration, writes `.env` and tokens)                                                                                                                                                                                                                                                                                                                                                                                                             |
| `src/create-mcp-config.ts` | Generates MCP client config from stored tokens                                                                                                                                                                                                                                                                                                                                                                                                                                                 |

### Key Architectural Patterns

- **Token resolution order**: environment variables → token file → refresh via refresh token
- **Token file locations**: platform config dir (`~/.config/microsoft-todo-mcp/tokens.json` on macOS/Linux, `%APPDATA%\microsoft-todo-mcp\tokens.json` on Windows); CLI also accepts `MSTODO_TOKEN_FILE` or falls back to `./tokens.json`
- **Multi-tenant support**: account type configurable via `TENANT_ID`
- **Error handling**: special-cases personal Microsoft accounts (`MailboxNotEnabledForRESTAPI`)
- **Validation**: strict TypeScript + Zod schemas for all tool parameters
- **Pagination**: `makeGraphRequestAll()` follows `@odata.nextLink` (used by `get-tasks` when `all: true` and by `archive-completed-tasks`)
- **Diagnostic tool gating**: `test-graph-api-exploration` is only registered when `MSTODO_ENABLE_EXPLORATION=1`
- **Logging**: diagnostics go to `stderr` only (stdout is reserved for MCP protocol)

### Microsoft Graph API Integration

- Base URL: `https://graph.microsoft.com/v1.0`
- Three-level hierarchy: Lists → Tasks → Checklist Items
- OData query parameters used for filtering/sorting
- Required scopes: `Tasks.Read`, `Tasks.ReadWrite`, `User.Read`

## Environment Configuration

- `.env` needed for auth flows: `CLIENT_ID`, `CLIENT_SECRET`, `TENANT_ID`, `REDIRECT_URI` (default `http://localhost:3000/callback`)
- `MSTODO_TOKEN_FILE`: custom token file path (CLI)
- Never commit `.env`, `tokens.json`, or any real tokens/secrets

## Important Notes

- Always run `pnpm run build` (or `pnpm run validate`) after modifying TypeScript files if you need to test against `dist/`
- CI (`.github/workflows/ci.yml`) runs `pnpm run validate` on Node 22.x/24.x; `publish.yml` publishes on release
- Personal Microsoft accounts have limited Graph API access compared to work/school accounts
- Reference docs live in `docs/` and API specs in `spec/`
