# Microsoft To Do MCP

[![CI](https://github.com/jordanburke/microsoft-todo-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/jordanburke/microsoft-todo-mcp-server/actions/workflows/ci.yml)
[![npm version](https://badge.fury.io/js/microsoft-todo-mcp-server.svg)](https://www.npmjs.com/package/microsoft-todo-mcp-server)

A Model Context Protocol (MCP) server that enables AI assistants like Claude and Cursor to interact with Microsoft To Do via the Microsoft Graph API. This service provides comprehensive task management capabilities through a secure OAuth 2.0 authentication flow.

## Just want to try it? Use Agent Todo

Running this server yourself means registering an app in Azure and managing its client secret. If you'd rather skip that, [Agent Todo](https://agent-todo.com) is a hosted MCP server that supports Microsoft To Do, from the same maintainer. You sign in to Agent Todo, connect your Microsoft To Do account on its Connections page, then add the connector to your AI assistant. There's nothing to register or install.

- **Free plan:** one connected task provider, such as Microsoft To Do, with read and write access.
- **Pro plan:** every provider (Google Tasks and Todoist too), with sync between them.

Connect it from any MCP client that supports remote servers. It signs you in with OAuth on first use.

```bash
# Claude Code
claude mcp add --transport http agent-todo https://api.agent-gate.com/agent-todo/mcp
```

In Claude Desktop or claude.ai, open Settings, then Connectors, and add `https://api.agent-gate.com/agent-todo/mcp` as a custom connector.

Want to run it on your own machine, or keep your tokens local? Keep reading to set up this server.

## Features

- **17 MCP Tools**: Complete task management functionality including lists, tasks, checklist items, and organization features
- **Recurring Tasks**: First-class support for Microsoft Graph `recurrence` (daily / weekly / absoluteMonthly / relativeMonthly / absoluteYearly / relativeYearly) on `create-task` and `update-task`
- **OAuth 2.0 Authentication**: Sign in once; tokens renew automatically
- **Microsoft Graph API Integration**: Direct integration with Microsoft's official API
- **Multi-tenant Support**: Works with personal, work, and school Microsoft accounts
- **Built on [somamcp](https://www.npmjs.com/package/somamcp)**: MCP SDK 2.x underneath, served over stdio
- **TypeScript**: Fully typed for reliability and developer experience
- **ESM Modules**: Modern JavaScript module system

## Prerequisites

- Node.js 20 or higher (CI tests 22.x and 24.x); building from source needs 22.18 or higher, or 24.11 or higher
- pnpm, only to build from source
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

The package provides five commands:

- `microsoft-todo-mcp-server` - The MCP server
- `mstodo` - Short alias for the MCP server
- `mstodo-config` - Writes an `mcp.json` that points the server at your `tokens.json`
- `mstodo-setup` - Guided setup. It does not yet finish on its own after you sign in, so use the steps under [Usage](#usage) instead
- `mstodo-http` - The MCP server over HTTP, for remote clients; each user signs in with their own Microsoft account (see [HTTP Mode](#http-mode-remote-clients))

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
5. Add a **Web** platform redirect URI: `http://localhost:3000/callback`. For [HTTP Mode](#http-mode-remote-clients), also add `<your base URL>/oauth/callback`
6. After creating the app, go to "Certificates & secrets" and create a new client secret. Note when it expires: sign-in and token renewal stop working on that date
7. Go to "API permissions" and add these Microsoft Graph delegated permissions. The sign-in requests these (plus the standard `openid` and `profile`), so adding them here is what lets an admin approve them ahead of time:
   - `Tasks.Read`, `Tasks.ReadWrite`, `Tasks.Read.Shared`, `Tasks.ReadWrite.Shared`
   - `User.Read`, `offline_access`
8. For a work or school tenant that requires it, an admin clicks "Grant admin consent". Personal accounts consent on the sign-in page

## Configuration

### Environment Setup

Create a `.env` file in the repository root, where you run `pnpm run auth` (required for sign-in):

```env
CLIENT_ID=your_client_id
CLIENT_SECRET=your_client_secret
TENANT_ID=your_tenant_setting
REDIRECT_URI=http://localhost:3000/callback
```

### TENANT_ID Options

- `organizations` - For multi-tenant organizational accounts (the default when `TENANT_ID` is unset)
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

The server stores authentication tokens in `tokens.json`. It looks in three places, in order:

1. `MSTODO_TOKEN_FILE`, if set
2. the per-user config directory — `~/.config/microsoft-todo-mcp/tokens.json`, or `%APPDATA%\microsoft-todo-mcp\tokens.json` on Windows
3. `tokens.json` in the working directory, which is migrated to whichever of (1) or (2) is in effect on first read

```bash
# Override the location
export MSTODO_TOKEN_FILE=/path/to/custom/tokens.json
```

You can also supply tokens directly, which takes precedence over the file:

```bash
export MS_TODO_ACCESS_TOKEN=your_access_token
export MS_TODO_REFRESH_TOKEN=your_refresh_token
```

These two are a starting point, not a live store — nothing rewrites your environment when a token is renewed. A renewal is kept for the rest of the session and written to the token file. On later launches the server uses that token file instead, as long as the environment still holds the same pair it was renewed from; set a new pair, for example after signing in again, and it starts over from that.

How the expiry is known depends on your account type. Work and school accounts get a JWT access token, so the server reads the real deadline from it and renews 5 minutes early. Personal accounts get an opaque token, where there is nothing to read and the server assumes an hour from when it first saw the value. Either way renewal needs `CLIENT_ID` and `CLIENT_SECRET`; without them the server cannot refresh, and calls start failing once the supplied token runs out.

## Usage

### Complete Setup Workflow

#### Step 1: Sign in with Microsoft

Sign-in runs from a clone of this repository, with the `.env` file from [Configuration](#configuration) in its root:

```bash
git clone https://github.com/jordanburke/microsoft-todo-mcp-server.git
cd microsoft-todo-mcp-server
pnpm install
pnpm run build
pnpm run auth
```

Open <http://localhost:3000> and sign in. Microsoft sends you back to `/callback`, and the server writes `tokens.json` in the repository root. Stop it with Ctrl+C when the page says authentication succeeded.

#### Step 2: Move the tokens where the server looks

`tokens.json` holds your refresh token and your client secret, so keep it private:

```bash
# macOS and Linux
mkdir -p ~/.config/microsoft-todo-mcp
mv tokens.json ~/.config/microsoft-todo-mcp/tokens.json
chmod 600 ~/.config/microsoft-todo-mcp/tokens.json
```

On Windows, move it to `%APPDATA%\microsoft-todo-mcp\tokens.json`. To keep it elsewhere, set `MSTODO_TOKEN_FILE` to its path in the next step.

#### Step 3: Configure Your AI Assistant

The server finds the token file itself and renews tokens with the client credentials stored in it, so the config needs no secrets.

**Claude Code:**

```bash
claude mcp add -s user microsoft-todo -- npx -y microsoft-todo-mcp-server

# Token file somewhere else:
claude mcp add -s user microsoft-todo -e MSTODO_TOKEN_FILE=/path/to/tokens.json -- npx -y microsoft-todo-mcp-server
```

`-s user` makes it available in every project; without it, Claude Code adds it only to the folder you ran the command in.

**Claude Desktop:** add this to its config file:

- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
- **Linux**: `~/.config/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "microsoftTodo": {
      "command": "npx",
      "args": ["--yes", "microsoft-todo-mcp-server"]
    }
  }
}
```

Add `"env": { "MSTODO_TOKEN_FILE": "/path/to/tokens.json" }` if the file is not in the default location.

**Cursor:** add the same `mcpServers` entry to `~/.cursor/mcp.json`.

**About `mstodo-config`:** `pnpm run create-config` (or `mstodo-config`) writes an `mcp.json` with the entry above, plus `MSTODO_TOKEN_FILE` set to the token file's absolute path. It copies no tokens into the config. It warns if the token file lacks `clientId` and `clientSecret`, which renewal needs.

**About token environment variables:** if you pass `MS_TODO_ACCESS_TOKEN` and `MS_TODO_REFRESH_TOKEN` instead, renewing them also needs `CLIENT_ID` and `CLIENT_SECRET` in the same `env` block, plus `TENANT_ID` unless it is `organizations`. Without them, calls start failing when the access token expires, within about an hour of signing in. The token-file setup above avoids this.

**Upgrading from an older version:** older versions copied renewed tokens into Claude Desktop's config as `MS_TODO_ACCESS_TOKEN` and `MS_TODO_REFRESH_TOKEN`, and older `mstodo-config` wrote them into `mcp.json`. The server now picks up its token file when that file holds the same refresh token, which heals the copied Claude Desktop entries. An old `mcp.json` sets no `MSTODO_TOKEN_FILE`, so it heals only if the server's token file (the default location, or `MSTODO_TOKEN_FILE`) is the one that pair came from. Either way, the fix that always works is to delete those two variables from the entry, so the server reads only the token file. Do the same if you sign in again: a fresh token file is ignored while an old pair is still set.

### Available Scripts

```bash
# Development & Building
pnpm run build        # Build TypeScript to JavaScript
pnpm run dev          # Build and run CLI in one command
pnpm run validate     # Format, lint, typecheck, test and build
pnpm test             # Run the tests

# Running the Server
pnpm start            # Run MCP server (dist/cli.js)
pnpm run cli          # Run MCP server via CLI wrapper
npx microsoft-todo-mcp-server  # Run globally installed version

# Authentication & Configuration
pnpm run auth         # Start OAuth authentication server
pnpm run create-config # Generate mcp.json from tokens.json

# Code Quality
pnpm run format       # Format code with Prettier
pnpm run format:check # Check code formatting
pnpm run lint         # Run linting checks
pnpm run typecheck    # TypeScript type checking
```

## HTTP Mode (remote clients)

`mstodo-http` serves the same tools over Streamable HTTP, for clients that connect by URL, such as a Claude custom connector. It is its own OAuth server in front of Microsoft, built on somamcp's Azure provider:

- **Each user signs in with their own Microsoft account**, and each tool call uses that user's token. The host's `tokens.json` and `MS_TODO_*` variables are never read in this mode.
- **The client never sees a Microsoft token.** It gets the server's own tokens; Microsoft's stay in an encrypted store on the server, which renews them when the client refreshes.
- **Stateless:** each request is served on its own, so it can run behind a load balancer. Sign-ins live in the store, not in memory.

### Setup

1. **On the app registration** (see [Azure App Registration](#azure-app-registration)), add a **Web** redirect URI: `<your base URL>/oauth/callback`, for example `https://todo.example.com/oauth/callback`.
2. **Set these**, in the environment or a `.env` in the working directory:

| Variable                                  | Required    | Meaning                                                                                                                                                                               |
| ----------------------------------------- | ----------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `MSTODO_HTTP_BASE_URL`                    | yes         | The server's public URL, such as `https://todo.example.com`. It must be `https`, except for `localhost`.                                                                              |
| `CLIENT_ID`, `CLIENT_SECRET`, `TENANT_ID` | yes         | The app registration. Use the tenant's id (a GUID) for a work or school tenant: sign-ins from other tenants are then refused. Use `consumers` or `common` for personal accounts.      |
| `JWT_SIGNING_KEY`, `TOKEN_ENCRYPTION_KEY` | yes         | At least 32 characters each, such as `openssl rand -base64 48`. Keep them: changing either signs everyone out.                                                                        |
| `MSTODO_HTTP_ALLOWED_USERS`               | recommended | Comma-separated accounts allowed to sign in, by email, sign-in name or object id. Without it, any account the app registration accepts can sign in, and each sees only its own To Do. |
| `MSTODO_HTTP_PORT`, `MSTODO_HTTP_HOST`    | no          | Where to listen. Defaults: `3100` on `127.0.0.1`. Use `0.0.0.0` in a container.                                                                                                       |
| `MSTODO_HTTP_STORE`                       | no          | Where sign-ins are stored, encrypted. Default `~/.config/microsoft-todo-mcp/http-store`.                                                                                              |
| `MSTODO_HTTP_REDIRECT_URIS`               | no          | Comma-separated redirect URI patterns a client may register. Default: `https://claude.ai/*`, `http://localhost:*/*`, `http://127.0.0.1:*/*`. A pattern like `https://*` is refused.   |

3. **Run it** behind HTTPS (a reverse proxy such as Caddy or nginx, or a tunnel):

```bash
mstodo-http          # or: npx -p microsoft-todo-mcp-server mstodo-http
```

It refuses to start, and lists what is wrong, until every required setting is valid.

4. **Connect a client** to `<your base URL>/mcp`. The client discovers the OAuth endpoints, registers itself, and sends you to Microsoft to sign in.

`auth-status` reports who is signed in and when their Microsoft token expires.

## MCP Tools

The server provides 17 tools for comprehensive Microsoft To Do management:

### Authentication and Server Info

- **`auth-status`** - Check authentication status, token expiration, and account type
- **`info`** - Report the server's name, version, build, runtime, and tool count (from [somamcp](https://www.npmjs.com/package/somamcp); build details come from the `SOMAMCP_BUILD_COMMIT`, `SOMAMCP_BUILD_DATE`, `SOMAMCP_BUILD_BRANCH` and `SOMAMCP_ENVIRONMENT` environment variables when set)

### Task Lists (Top-level Containers)

- **`get-task-lists`** - Retrieve all task lists with metadata (default, shared, etc.)
- **`create-task-list`** - Create a new task list
- **`update-task-list`** - Rename an existing task list
- **`delete-task-list`** - Delete a task list and all its contents
- **`get-task-lists-organized`** - Group task lists into folders by naming pattern, emoji prefix, or sharing status (`groupBy`: `category`, `shared`, or `type`; `includeIds` adds an ID reference)

### Tasks (Main Todo Items)

- **`get-tasks`** - Get tasks from a list with filtering, sorting, and pagination
  - Supports OData query parameters: `$filter`, `$orderby`, `$top`, `$skip`, `$count`
  - `orderby` works on date fields such as `createdDateTime`; Graph rejects `orderby=title`
  - `top` caps the total, and the server follows Graph's paging to collect every page
  - `count` prints the number of tasks returned; Graph itself returns no count for To Do tasks
  - `select` limits which fields print (`status`, `dueDateTime`, `importance`, `categories`, `recurrence`, `body`); it is applied by the server, not sent to Graph
- **`create-task`** - Create a new task with full property support
  - Title, description, due date, start date, importance, reminders, status, categories, **recurrence**
- **`update-task`** - Update any task properties including **recurrence** (pass `null` to stop a recurring task)
- **`delete-task`** - Delete a task and all its checklist items

### Recurring Tasks

Pass a `recurrence` object to `create-task` or `update-task` to make a task repeat.
Recurrence derives its first occurrence from the task's `dueDateTime`, which is
required. Examples:

```json
// Monthly bill, 15th of every month, forever
{
  "pattern": {"type": "absoluteMonthly", "interval": 1, "dayOfMonth": 15},
  "range":   {"type": "noEnd"}
}

// Weekly standup every Monday, ends 2027-01-01
{
  "pattern": {"type": "weekly", "interval": 1, "daysOfWeek": ["monday"]},
  "range":   {"type": "endDate", "endDate": "2027-01-01"}
}

// Annual birthday, on the 4th of September
{
  "pattern": {"type": "absoluteYearly", "interval": 1, "dayOfMonth": 4, "month": 9},
  "range":   {"type": "noEnd"}
}
```

Each pattern type needs its own fields, and the server refuses a rule that leaves one
out or adds one the type does not use: `weekly` and `relativeMonthly` take `daysOfWeek`,
`absoluteMonthly` takes `dayOfMonth`, `absoluteYearly` takes `dayOfMonth` and `month`,
and `relativeYearly` takes `daysOfWeek` and `month`. A `weekly` rule without
`firstDayOfWeek` gets `sunday`. An `endDate` range needs `endDate`, and a `numbered`
range needs `numberOfOccurrences`.

Stop a recurring task by passing `recurrence: null` to `update-task`. See
<https://learn.microsoft.com/en-us/graph/api/resources/recurrencepattern> for
the full schema.

### Checklist Items (Subtasks)

- **`get-checklist-items`** - Get subtasks for a specific task
- **`create-checklist-item`** - Add a new subtask to a task
- **`update-checklist-item`** - Update subtask text or completion status
- **`delete-checklist-item`** - Remove a specific subtask

### Maintenance and Diagnostics

- **`archive-completed-tasks`** - Move completed tasks older than `olderThanDays` (default 90) from one list to another; `dryRun` previews without moving
- **`test-graph-api-exploration`** - Run read-only Graph API probes (`odata-select`, `odata-expand`, `headers`, `extensions`, or `all`) to look for hidden list properties; for debugging, not everyday use

## Architecture

### Project Structure

- **MCP Server** (`src/todo-index.ts`) - The 16 To Do tools, registered on a somamcp server
- **Token Manager** (`src/token-manager.ts`) - Finds, renews and saves tokens
- **CLI Wrapper** (`src/cli.ts`) - Executable entry point; starts the server over stdio
- **Auth Server** (`src/auth-server.ts`) - Express server for the OAuth 2.0 sign-in
- **Setup** (`src/setup.ts`) - The guided `mstodo-setup` command
- **Config Generator** (`src/create-mcp-config.ts`) - Helper to create MCP configurations

### Technical Details

- **MCP framework**: [somamcp](https://www.npmjs.com/package/somamcp) 2.x, with Zod 4 parameter schemas
- **Microsoft Graph API**: Uses v1.0 endpoints
- **Authentication**: MSAL (Microsoft Authentication Library) authorization-code flow with a client secret
- **Token Management**: Renews 5 minutes before expiry, and again whenever Graph rejects a token
- **Build System**: ts-builds (tsdown) for fast TypeScript compilation
- **Module System**: ESM (ECMAScript modules)

## Limitations & Known Issues

### Personal Microsoft Accounts

- Personal accounts (outlook.com, hotmail.com, live.com) work. A live.com account tested on 2026-10-31 had full access: 30 lists, paging through a 399-task list, writes and recurring tasks
- If you see `MailboxNotEnabledForRESTAPI`, Microsoft has not enabled the To Do API for that mailbox. It is a Microsoft-side limit, not a bug in this server, and the server says so when it happens

### API Limitations

- Rate limits apply according to Microsoft's policies
- Shared lists have limited functionality

## Troubleshooting

### Authentication Issues

**Token acquisition failures**

- `unauthorized_client: The client does not exist or is not enabled for consumers`: Microsoft can't use the app for this account. Either the app registration was deleted, or its supported account types exclude personal accounts
- `AADSTS7000215` (invalid client secret) or `AADSTS7000222` (expired client secret): create a new secret under "Certificates & secrets" and update `.env`, then sign in again so `tokens.json` stores it
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

# Or read the expiry from the token file (it also holds secrets, so don't share its contents)
jq '.expiresAt' ~/.config/microsoft-todo-mcp/tokens.json

# Convert the timestamp to a readable date
date -d @$(($(jq -r '.expiresAt' ~/.config/microsoft-todo-mcp/tokens.json) / 1000))
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
3. Run `pnpm run validate` (format, lint, typecheck, test, build) before submitting
4. Submit a pull request

## License

MIT License - See [LICENSE](LICENSE) file for details

## Acknowledgments

- Fork of [@jhirono/todomcp](https://github.com/jhirono/todomcp)
- Built on [somamcp](https://www.npmjs.com/package/somamcp) and the [Model Context Protocol](https://modelcontextprotocol.io) TypeScript SDK
- Uses [Microsoft Graph API](https://developer.microsoft.com/en-us/graph)

## Support

- [GitHub Issues](https://github.com/jordanburke/microsoft-todo-mcp-server/issues)
- [npm Package](https://www.npmjs.com/package/microsoft-todo-mcp-server)
