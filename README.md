# Microsoft To Do MCP

[![CI](https://github.com/jordanburke/microsoft-todo-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/jordanburke/microsoft-todo-mcp-server/actions/workflows/ci.yml)
[![npm version](https://badge.fury.io/js/microsoft-todo-mcp-server.svg)](https://www.npmjs.com/package/microsoft-todo-mcp-server)

A Model Context Protocol (MCP) server that enables AI assistants such as ChatGPT, Claude, and Cursor to interact with Microsoft To Do through the Microsoft Graph API. It supports local stdio clients and a remote OAuth-protected Streamable HTTP endpoint.

## Features

- **16 MCP Tools**: Complete task management functionality including lists, tasks, checklist items, archiving, and organization features
- **Microsoft OAuth**: Delegated Graph authorization with automatic refresh after the initial interactive consent
- **Remote MCP**: Streamable HTTP endpoint for ChatGPT and other network clients
- **Built-in OAuth AS/RS**: Single-user authorization page, PKCE, dynamic client registration, refresh-token rotation, and RFC 8707 resource binding
- **Microsoft Graph API Integration**: Direct integration with Microsoft's official API
- **Multi-tenant Support**: Works with personal, work, and school Microsoft accounts
- **TypeScript**: Fully typed for reliability and developer experience
- **ESM Modules**: Modern JavaScript module system

## Prerequisites

- Node.js 22.13 or higher when building from source
- pnpm package manager
- A Microsoft account (personal, work, or school)
- Azure App Registration (see setup below)

## Installation

### Option 1: Global Installation (Recommended)

```bash
# Install globally using npm
npm install -g microsoft-todo-mcp-server

# Or using pnpm
pnpm install -g microsoft-todo-mcp-server

# Or run directly with npx (no installation)
npx microsoft-todo-mcp-server
```

The package provides four command aliases:

- `microsoft-todo-mcp-server` - Full package name
- `mstodo` - Short alias for the MCP server
- `mstodo-http` - OAuth-protected Streamable HTTP server
- `mstodo-config` - Configuration helper tool

### Option 2: Clone and Run Locally

```bash
git clone https://github.com/jordanburke/microsoft-todo-mcp-server.git
cd microsoft-todo-mcp-server
pnpm install
pnpm run build
```

## Azure App Registration

1. Go to the [Azure Portal](https://portal.azure.com)
2. Navigate to "App registrations" and create a new registration
3. Name your application (e.g., "To Do MCP")
4. For "Supported account types", select one of the following based on your needs:
   - **Accounts in this organizational directory only (Single tenant)** - For use within a single organization
   - **Accounts in any organizational directory (Any Azure AD directory - Multitenant)** - For use across multiple organizations
   - **Accounts in any organizational directory and personal Microsoft accounts** - For both work accounts and personal accounts
5. Set the Redirect URI to `http://localhost:3000/callback`
6. After creating the app, go to "Certificates & secrets" and create a new client secret
7. Go to "API permissions" and add the following permissions:
   - Microsoft Graph > Delegated permissions:
     - Tasks.Read
     - Tasks.ReadWrite
     - User.Read
8. Click "Grant admin consent" for these permissions

## Configuration

### Environment Setup

Create a `.env` file in the project root (required for authentication):

```env
CLIENT_ID=your_client_id
CLIENT_SECRET=your_client_secret
TENANT_ID=your_tenant_setting
REDIRECT_URI=http://localhost:3000/callback
```

### TENANT_ID Options

- `organizations` - For multi-tenant organizational accounts (default if not specified)
- `consumers` - For personal Microsoft accounts only
- `common` - For both organizational and personal accounts
- `your-specific-tenant-id` - For single-tenant configurations

**Examples:**

```env
# For multi-tenant organizational accounts (default)
TENANT_ID=organizations

# For personal Microsoft accounts
TENANT_ID=consumers

# For both organizational and personal accounts
TENANT_ID=common

# For a specific organization tenant
TENANT_ID=00000000-0000-0000-0000-000000000000
```

### Token Storage

The Microsoft sign-in flow stores an access token and refresh token in `tokens.json`. `pnpm setup` installs it at `~/.config/microsoft-todo-mcp/tokens.json` on Linux/macOS; `pnpm auth` first writes `tokens.json` in the current project. The server refreshes file-based tokens five minutes before expiration.

An Azure client ID and secret do not replace the initial Microsoft sign-in and consent. At least one interactive authorization is required before either the local or remote MCP server can access Microsoft To Do.

For local compatibility, tokens can also be supplied directly:

```bash
export MS_TODO_ACCESS_TOKEN=your_access_token
export MS_TODO_REFRESH_TOKEN=your_refresh_token
```

For a long-running remote deployment, prefer the persisted token file so refreshed tokens survive container restarts. Never commit `.env`, `tokens.json`, or generated MCP configuration containing tokens.

## Usage

### Complete Setup Workflow

#### Step 1: Authenticate with Microsoft

```bash
# If installed globally
git clone https://github.com/jordanburke/microsoft-todo-mcp-server.git
cd microsoft-todo-mcp-server
pnpm install
pnpm run auth

# Or if running locally
pnpm run auth
```

This opens a browser window for Microsoft authentication and creates a `tokens.json` file.

#### Step 2: Create MCP Configuration

```bash
# Generate MCP configuration file
pnpm run create-config

# Or use the global helper (if installed globally)
mstodo-config
```

This creates an `mcp.json` file with your authentication tokens.

#### Step 3: Configure Your AI Assistant

**For Claude Desktop:**

Add to your configuration file:

- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
- **Linux**: `~/.config/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "microsoftTodo": {
      "command": "npx",
      "args": ["--yes", "microsoft-todo-mcp-server"],
      "env": {
        "MS_TODO_ACCESS_TOKEN": "your_access_token",
        "MS_TODO_REFRESH_TOKEN": "your_refresh_token"
      }
    }
  }
}
```

**For Cursor:**

```bash
# Copy to Cursor's global configuration
cp mcp.json ~/.cursor/mcp-servers.json
```

## Remote MCP over HTTP

The remote mode exposes a Streamable HTTP endpoint at `/mcp`. It is intended for network clients that cannot execute the local stdio command, including a [ChatGPT custom MCP connector](https://chatgpt.com/plugins#settings/Connectors?create-connector=true&redirectAfter=%2Fplugins). OpenAI also documents remote MCP servers as public servers identified by a `server_url`, optionally protected by OAuth authorization in its [MCP and Connectors guide](https://developers.openai.com/api/docs/guides/tools-connectors-mcp).

### What the built-in OAuth protects

Remote mode has two separate authorization boundaries:

| Authorization             | Connection                                     | Purpose                                                           | Persistence                                         |
| ------------------------- | ---------------------------------------------- | ----------------------------------------------------------------- | --------------------------------------------------- |
| Microsoft delegated OAuth | This server → Microsoft Graph                  | Allows this server to read and update one Microsoft To Do account | `tokens.json`, including the refresh token          |
| Built-in MCP OAuth AS/RS  | ChatGPT or another remote client → this server | Controls which MCP clients can invoke the exposed tools           | In memory; clients reconnect after a server restart |

The built-in MCP OAuth server supports dynamic client registration, authorization code flow with S256 PKCE, RFC 8707 resource binding, access-token verification, refresh-token rotation, and revocation. It issues tokens scoped only to `mcp:tools` for this server's `/mcp` resource.

`MCP_OAUTH_PASSWORD` protects the single-user approval page. It is not the Microsoft account password and cannot obtain a Microsoft Graph token. Microsoft access and refresh tokens are never returned to ChatGPT or another MCP client.

### 1. Authorize Microsoft To Do once

Complete the existing interactive Microsoft authorization on a trusted machine:

```bash
pnpm install
pnpm build
pnpm setup
```

On Linux/macOS, this creates `~/.config/microsoft-todo-mcp/tokens.json`. A cloud deployment can securely copy that file to the server; filling only `CLIENT_ID`, `CLIENT_SECRET`, and `TENANT_ID` in `.env` does not perform the initial user consent automatically.

### 2. Configure the remote server

```bash
cp .env.example .env
openssl rand -base64 32
```

Put the generated value in `MCP_OAUTH_PASSWORD`, then configure both the Microsoft application and remote MCP settings:

```env
CLIENT_ID=your_azure_app_client_id
CLIENT_SECRET=your_azure_app_client_secret
TENANT_ID=organizations
REDIRECT_URI=http://localhost:3000/callback

# Public HTTPS origin only; do not add /mcp here.
MCP_PUBLIC_URL=https://todo.example.com
MCP_OAUTH_PASSWORD=your-long-random-password
MCP_HOST=127.0.0.1
MCP_PORT=3001
MCP_TRUST_PROXY=1
```

`MCP_PUBLIC_URL` must be the exact externally visible HTTPS origin. Its hostname is also added to the Host-header allowlist. `MCP_TRUST_PROXY` is the number of trusted reverse proxies between the client and this process.

### 3. Start the Streamable HTTP server

```bash
pnpm build
pnpm http
```

The public MCP URL is the origin plus `/mcp`:

```text
https://todo.example.com/mcp
```

Production deployments must terminate HTTPS in front of the Node process. A minimal Caddy configuration is:

```caddyfile
todo.example.com {
  reverse_proxy 127.0.0.1:3001
}
```

Do not expose the unencrypted Node port directly to the internet.

### 4. Add it to ChatGPT

1. Open [ChatGPT's custom connector creation page](https://chatgpt.com/plugins#settings/Connectors?create-connector=true&redirectAfter=%2Fplugins).
2. Choose **Server URL** and enter `https://todo.example.com/mcp`.
3. Choose **OAuth** authentication.
4. Save or connect the server. ChatGPT discovers the OAuth metadata and dynamically registers itself.
5. On the Microsoft To Do MCP approval page, enter `MCP_OAUTH_PASSWORD` and approve access.

No OAuth client ID or client secret needs to be created manually for ChatGPT. The built-in authorization server publishes the protected-resource and authorization-server metadata required for discovery.

### 5. Verify the public endpoints

```bash
curl https://todo.example.com/healthz
curl https://todo.example.com/.well-known/oauth-protected-resource/mcp
curl https://todo.example.com/.well-known/oauth-authorization-server
curl -i -X POST https://todo.example.com/mcp \
  -H 'Accept: application/json, text/event-stream' \
  -H 'Content-Type: application/json' \
  --data '{"jsonrpc":"2.0","method":"initialize","params":{"protocolVersion":"2025-06-18","capabilities":{},"clientInfo":{"name":"probe","version":"1"}},"id":1}'
```

The health and metadata requests should return `200`. The final request should return `401 Unauthorized` with a `WWW-Authenticate` header whose `resource_metadata` points to `/.well-known/oauth-protected-resource/mcp`. That challenge starts OAuth discovery; it does not indicate a server failure.

### Docker Compose

Copy the Microsoft token file into the bind-mounted data directory before starting the container:

```bash
mkdir -p data
cp ~/.config/microsoft-todo-mcp/tokens.json data/tokens.json
chmod 600 .env data/tokens.json
# The runtime image uses UID 1000 and must be able to update refreshed tokens.
chown 1000:1000 data/tokens.json
docker compose up -d --build
```

Compose reads the service environment from the project-root `.env` by default, binds the service only to host loopback, and mounts `./data` at `/home/node/.config/microsoft-todo-mcp` inside the container.

To keep the environment file under `data/` and publish the service on host port `3333`, use:

```bash
MCP_ENV_FILE=./data/.env MCP_PORT=3333 docker compose up -d --build
```

After changing the environment file, recreate the container so Docker reloads it:

```bash
MCP_ENV_FILE=./data/.env MCP_PORT=3333 docker compose up -d --force-recreate
```

### Single-user limitations

- Every approved MCP client acts on the same Microsoft To Do account represented by `data/tokens.json`.
- OAuth clients, authorization codes, and MCP access/refresh tokens are kept in process memory. Restarting the service requires reconnecting the custom MCP client.
- Run one application replica. This lightweight mode intentionally has no shared OAuth database or multi-user account system.
- Keep `.env` and `tokens.json` private. They contain the Azure client secret and Microsoft delegated credentials.

## Available Scripts

```bash
# Development & Building
pnpm run build        # Build TypeScript to JavaScript
pnpm run dev          # Build and run CLI in one command
pnpm run dev:http     # Build and run the remote HTTP server

# Running the Server
pnpm start            # Run MCP server directly
pnpm run cli          # Run MCP server via CLI wrapper
pnpm run http         # Run the OAuth-protected Streamable HTTP server
npx microsoft-todo-mcp-server  # Run globally installed version

# Authentication & Configuration
pnpm run auth         # Start OAuth authentication server
pnpm run setup        # Interactive Microsoft app and token setup
pnpm run create-config # Generate mcp.json from tokens.json

# Code Quality
pnpm run validate     # Format, lint, typecheck, test, and build
pnpm run format       # Format code with Prettier
pnpm run format:check # Check code formatting
pnpm run lint         # Run linting checks
pnpm run typecheck    # TypeScript type checking
pnpm run test         # Run Vitest tests
```

## MCP Tools

The server provides 16 tools for comprehensive Microsoft To Do management:

### Authentication

- **`auth-status`** - Check authentication status, token expiration, and account type

### Task Lists (Top-level Containers)

- **`get-task-lists`** - Retrieve all task lists with metadata (default, shared, etc.)
- **`get-task-lists-organized`** - Group task lists by naming patterns, emoji prefixes, and sharing status
- **`create-task-list`** - Create a new task list
- **`update-task-list`** - Rename an existing task list
- **`delete-task-list`** - Delete a task list and all its contents

### Tasks (Main Todo Items)

- **`get-tasks`** - Get tasks from a list with filtering, sorting, and pagination
  - Supports OData query parameters: `$filter`, `$select`, `$orderby`, `$top`, `$skip`, `$count`
  - `$select=title` is normalized out because Microsoft Graph returns `title` by default but rejects explicitly selecting it on this endpoint
- **`create-task`** - Create a new task with full property support
  - Title, description, due date, start date, importance, reminders, status, categories
- **`update-task`** - Update any task properties
- **`delete-task`** - Delete a task and all its checklist items

### Checklist Items (Subtasks)

- **`get-checklist-items`** - Get subtasks for a specific task
- **`create-checklist-item`** - Add a new subtask to a task
- **`update-checklist-item`** - Update subtask text or completion status
- **`delete-checklist-item`** - Remove a specific subtask

### Maintenance and Diagnostics

- **`archive-completed-tasks`** - Move older completed tasks into an archive list
- **`test-graph-api-exploration`** - Probe Graph list metadata and organization-related capabilities

## Architecture

### Project Structure

- **MCP Server** (`src/todo-index.ts`) - Core server implementing the MCP protocol
- **CLI Wrapper** (`src/cli.ts`) - Executable entry point with token management
- **Auth Server** (`src/auth-server.ts`) - Express server for OAuth 2.0 flow
- **Remote HTTP Server** (`src/http-server.ts`) - Streamable HTTP resource server and OAuth routes
- **Remote OAuth Provider** (`src/remote-auth.ts`) - Single-user in-memory authorization server
- **Config Generator** (`src/create-mcp-config.ts`) - Helper to create MCP configurations
- **HTTP/OAuth Tests** (`src/http-server.test.ts`) - End-to-end discovery, PKCE, token, and MCP request coverage
- **Query Tests** (`src/todo-index.test.ts`) - Microsoft Graph task-query normalization coverage

### Technical Details

- **Microsoft Graph API**: Uses v1.0 endpoints
- **Microsoft Authentication**: MSAL delegated flow for Microsoft Graph
- **Remote MCP Authentication**: OAuth 2.1-style authorization code flow with S256 PKCE and dynamic client registration
- **Token Management**: Automatic refresh 5 minutes before expiration
- **Build System**: ts-builds (tsdown) for fast TypeScript compilation
- **Module System**: ESM (ECMAScript modules)

## Limitations & Known Issues

### Personal Microsoft Accounts

- **MailboxNotEnabledForRESTAPI Error**: Personal Microsoft accounts (outlook.com, hotmail.com, live.com) have limited access to the To Do API through Microsoft Graph
- This is a Microsoft service limitation, not an issue with this application
- Work/school accounts have full API access

### API Limitations

- Rate limits apply according to Microsoft's policies
- Some features may be unavailable for personal accounts
- Shared lists have limited functionality

## Troubleshooting

### Authentication Issues

**Token acquisition failures**

- Verify `CLIENT_ID`, `CLIENT_SECRET`, and `TENANT_ID` in your `.env` file
- Ensure redirect URI matches exactly: `http://localhost:3000/callback`
- Check Azure App permissions are granted with admin consent

**Permission issues**

- Ensure all required Graph API permissions are added and consented
- For organizational accounts, admin consent may be required

### Account Type Configuration

**Work/School Accounts**

```env
TENANT_ID=organizations  # Multi-tenant
# Or use your specific tenant ID
```

**Personal Accounts**

```env
TENANT_ID=consumers  # Personal only
# Or TENANT_ID=common for both types
```

### Debugging

**Check authentication status:**

```bash
# Using the MCP tool
# In your AI assistant: "Check auth status"

# Or examine tokens directly
cat tokens.json | jq '.expiresAt'

# Convert timestamp to readable date
date -d @$(($(cat tokens.json | jq -r '.expiresAt') / 1000))
```

**Enable verbose logging:**

```bash
# The server logs to stderr for debugging
mstodo 2> debug.log
```

## Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Run `pnpm run lint` and `pnpm run typecheck` before submitting
4. Submit a pull request

## License

MIT License - See [LICENSE](LICENSE) file for details

## Acknowledgments

- Fork of [@jhirono/todomcp](https://github.com/jhirono/todomcp)
- Built on the [Model Context Protocol SDK](https://github.com/modelcontextprotocol/sdk)
- Uses [Microsoft Graph API](https://developer.microsoft.com/en-us/graph)

## Support

- [GitHub Issues](https://github.com/jordanburke/microsoft-todo-mcp-server/issues)
- [npm Package](https://www.npmjs.com/package/microsoft-todo-mcp-server)
