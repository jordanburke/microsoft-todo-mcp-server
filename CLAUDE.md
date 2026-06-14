# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Common Development Commands

### Build and Development

```bash
pnpm install         # Install dependencies (pnpm 11)
pnpm run build       # Build with ts-builds (tsdown) to dist/ directory
pnpm run dev         # Build and run CLI in one command
pnpm run validate    # Full chain: format, lint, typecheck, test, build
```

Build tooling is provided by **ts-builds** (3.x); package scripts delegate to the
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
