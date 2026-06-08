# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this
repository.

## Common Development Commands

### Running the Server

```bash
deno task dev        # Run MCP server directly (todo-index.ts)
deno task start      # Run via CLI wrapper (cli.ts)
```

### Authentication

```bash
deno task auth       # Start OAuth authentication server (port 3000)
deno task setup      # Interactive setup wizard
```

### Configuration

```bash
deno task config     # Generate mcp.json from tokens.json
```

### Quality Checks

```bash
deno fmt             # Format code with Deno's built-in formatter
deno fmt --check     # Check formatting
deno lint            # Lint code
deno check src/**/*.ts  # Type check
```

## Architecture Overview

This is a Model Context Protocol (MCP) server that enables AI assistants to interact with Microsoft
To Do via the Microsoft Graph API. The codebase follows a modular architecture with four main
components:

1. **MCP Server** (`src/todo-index.ts`): Core server implementing the MCP protocol with 14 tools for
   Microsoft To Do operations
2. **CLI Wrapper** (`src/cli.ts`): Executable entry point that handles token loading from
   environment or file
3. **Auth Server** (`src/auth-server.ts`): Deno HTTP server implementing OAuth 2.0 flow with MSAL
4. **Config Generator** (`src/create-mcp-config.ts`): Utility to create MCP configuration files

### Key Architectural Patterns

- **Token Management** (`src/token-manager.ts`): Tokens are stored in `tokens.json` with automatic
  refresh 5 minutes before expiration
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

## Technology Stack

- **Runtime**: Deno 2.x (TypeScript native, no build step)
- **Dependencies via npm specifiers**: `@azure/msal-node`, `@modelcontextprotocol/sdk`, `zod`
- **Dependencies via JSR**: `@std/path`
- All imports are managed in `deno.json`

## Important Notes

- No build step needed — Deno runs TypeScript directly
- The auth server uses `Deno.serve()` instead of Express
- Tokens are automatically refreshed using the refresh token when needed
- Personal Microsoft accounts have limited API access compared to work/school accounts
- Required permissions: `--allow-env`, `--allow-read`, `--allow-write`, `--allow-net`
