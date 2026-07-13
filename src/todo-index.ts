import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js"
import dotenv from "dotenv"
import { z } from "zod"

import { tokenManager } from "./token-manager.js"

// Load environment variables
dotenv.config()

// Log the current working directory
console.error("Current working directory:", process.cwd())

// Microsoft Graph API endpoints
const MS_GRAPH_BASE = "https://graph.microsoft.com/v1.0"
const USER_AGENT = "microsoft-todo-mcp-server/1.0"

// Create server instance
const server = new McpServer({
  name: "mstodo",
  version: "1.0.0",
})

// Every Graph call resolves to one of these — never a bare null. `ok:false`
// carries the real HTTP status + body so handlers can surface the true reason.
// The detail travels *in the result*, not in shared module state, so concurrent
// requests can never clobber each other's error (fixes the old lastGraphError race).
type GraphResult<T> = { ok: true; status: number; data: T } | { ok: false; status: number; error: string }

// Standard failure return for handlers when a Graph call failed. Appends the
// captured HTTP status/body so Claude can act on the real reason.
function graphFail(msg: string, res: { status: number; error: string }) {
  return {
    content: [
      {
        type: "text" as const,
        text: `${msg} — ${res.error}`,
      },
    ],
  }
}

// Return used when we cannot even obtain an access token (before any Graph call).
function authFail() {
  return {
    content: [
      {
        type: "text" as const,
        text: "Failed to authenticate with Microsoft API. Your tokens may have expired — re-run the auth/setup flow, then restart the client.",
      },
    ],
  }
}

// Core Microsoft Graph request. Always resolves to a GraphResult — never throws
// to the caller, never returns a bare null. Handles empty/204 bodies (DELETE and
// some PATCH/POST) so a successful mutation is distinguishable from a failure.
async function graphRequest<T>(url: string, token: string, method = "GET", body?: any): Promise<GraphResult<T>> {
  const hasBody = body !== undefined && (method === "POST" || method === "PATCH" || method === "PUT")
  const headers: Record<string, string> = {
    "User-Agent": USER_AGENT,
    Accept: "application/json",
    Authorization: `Bearer ${token}`,
  }
  if (hasBody) headers["Content-Type"] = "application/json"

  try {
    const options: RequestInit = { method, headers }

    if (hasBody) {
      options.body = JSON.stringify(body)
    }

    console.error(`Making request to: ${method} ${url}`)

    let response = await fetch(url, options)

    // On 401, refresh the token once and retry.
    if (response.status === 401) {
      console.error("Got 401, attempting token refresh...")
      const newToken = await getAccessToken() // triggers refresh via TokenManager
      if (newToken && newToken !== token) {
        headers.Authorization = `Bearer ${newToken}`
        response = await fetch(url, { ...options, headers })
      }
    }

    if (!response.ok) {
      const errorText = await response.text()
      console.error(`HTTP error! status: ${response.status}, body: ${errorText}`)

      if (errorText.includes("MailboxNotEnabledForRESTAPI")) {
        return {
          ok: false,
          status: response.status,
          error:
            "MailboxNotEnabledForRESTAPI — the Microsoft To Do API is not available for this personal " +
            "Microsoft account through Graph. This is a Microsoft limitation, not an auth issue.",
        }
      }

      return {
        ok: false,
        status: response.status,
        error: `HTTP ${response.status}: ${errorText.substring(0, 500)}`,
      }
    }

    // Success. 204 and other empty bodies have nothing to parse; return {} so the
    // caller sees ok:true (a successful DELETE must not look like a failure).
    if (response.status === 204) {
      return { ok: true, status: response.status, data: {} as T }
    }
    const text = await response.text()
    if (!text) {
      return { ok: true, status: response.status, data: {} as T }
    }
    console.error(`Response received: ${text.substring(0, 200)}...`)
    return { ok: true, status: response.status, data: JSON.parse(text) as T }
  } catch (error) {
    console.error("Error making Graph API request:", error)
    return { ok: false, status: 0, error: `Request failed: ${String(error)}` }
  }
}

// Fetch every page of a Graph collection, following @odata.nextLink so callers
// get complete results instead of a silently-truncated first page.
async function graphGetAll<T>(url: string, token: string): Promise<GraphResult<{ value: T[] }>> {
  const all: T[] = []
  let next: string | null = url
  let guard = 0
  while (next) {
    if (++guard > 100) {
      console.error("graphGetAll: pagination guard tripped (>100 pages)")
      break
    }
    const res = await graphRequest<{ value: T[]; "@odata.nextLink"?: string }>(next, token)
    if (!res.ok) return res
    all.push(...(res.data.value || []))
    next = res.data["@odata.nextLink"] || null
  }
  return { ok: true, status: 200, data: { value: all } }
}

// Authentication helper using delegated flow with token manager
async function getAccessToken(): Promise<string | null> {
  try {
    console.error("getAccessToken called")

    // Use the token manager to get tokens (handles all sources and refresh)
    const tokens = await tokenManager.getTokens()

    if (tokens) {
      console.error(`Successfully retrieved valid token`)
      return tokens.accessToken
    }

    console.error("No valid tokens available")
    return null
  } catch (error) {
    console.error("Error getting access token:", error)
    return null
  }
}

// Server configuration type
interface ServerConfig {
  accessToken?: string
  refreshToken?: string
  tokenFilePath?: string
}

// Function to check if the account is a personal Microsoft account
async function isPersonalMicrosoftAccount(): Promise<boolean> {
  try {
    const token = await getAccessToken()
    if (!token) return false

    // Make a request to get user info
    const url = `${MS_GRAPH_BASE}/me`
    const response = await fetch(url, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
      },
    })

    if (!response.ok) {
      console.error(`Error getting user info: ${response.status}`)
      return false
    }

    const userData = await response.json()
    const email = userData.mail || userData.userPrincipalName || ""

    // Check if the email domain indicates a personal account
    const personalDomains = ["outlook.com", "hotmail.com", "live.com", "msn.com", "passport.com"]
    const domain = email.split("@")[1]?.toLowerCase()

    if (domain && personalDomains.some((d) => domain.includes(d))) {
      console.error(`
=================================================================
WARNING: Personal Microsoft Account Detected

Your Microsoft account (${email}) appears to be a personal account.
Microsoft To Do API access is typically not available for personal accounts
through the Microsoft Graph API, only for Microsoft 365 business accounts.

You may encounter the "MailboxNotEnabledForRESTAPI" error when trying to
access To Do lists or tasks. This is a limitation of the Microsoft Graph API,
not an issue with your authentication or this application.

You can still use Microsoft To Do through the web interface or mobile apps,
but API access is restricted for personal accounts.
=================================================================
      `)
      return true
    }

    return false
  } catch (error) {
    console.error("Error checking account type:", error)
    return false
  }
}

// Server tool to check authentication status
server.tool(
  "auth-status",
  "Check if you're authenticated with Microsoft Graph API. Shows current token status and expiration time, and indicates if the token needs to be refreshed.",
  {},
  async () => {
    const tokens = await tokenManager.getTokens()

    if (!tokens) {
      return {
        content: [
          {
            type: "text",
            text: "Not authenticated. Please run 'npx microsoft-todo-mcp-server setup' to authenticate with Microsoft.",
          },
        ],
      }
    }

    const isExpired = Date.now() > tokens.expiresAt
    const expiryTime = new Date(tokens.expiresAt).toLocaleString()

    // Check if it's a personal account
    const isPersonal = await isPersonalMicrosoftAccount()
    let accountMessage = ""

    if (isPersonal) {
      accountMessage =
        "\n\n⚠️ WARNING: You are using a personal Microsoft account. " +
        "Microsoft To Do API access is typically not available for personal accounts " +
        "through the Microsoft Graph API. You may encounter 'MailboxNotEnabledForRESTAPI' errors. " +
        "This is a Microsoft limitation, not an authentication issue."
    }

    if (isExpired) {
      return {
        content: [
          {
            type: "text",
            text: `Authentication expired at ${expiryTime}. Will attempt to refresh when you call any API.${accountMessage}`,
          },
        ],
      }
    } else {
      return {
        content: [
          {
            type: "text",
            text: `Authenticated. Token expires at ${expiryTime}.${accountMessage}`,
          },
        ],
      }
    }
  },
)

interface TaskList {
  id: string
  displayName: string
  isOwner?: boolean
  isShared?: boolean
  wellknownListName?: string // 'none', 'defaultList', 'flaggedEmails', 'unknownFutureValue'
}

interface DateTimeTimeZone {
  dateTime: string
  timeZone: string
}

interface Task {
  id: string
  title: string
  status: string
  importance: string
  isReminderOn?: boolean
  dueDateTime?: DateTimeTimeZone
  startDateTime?: DateTimeTimeZone
  completedDateTime?: DateTimeTimeZone
  reminderDateTime?: DateTimeTimeZone
  body?: {
    content: string
    contentType: string
  }
  categories?: string[]
}

interface ChecklistItem {
  id: string
  displayName: string
  isChecked: boolean
  createdDateTime?: string
}

// Register tools
server.tool(
  "get-task-lists",
  "Get all Microsoft Todo task lists (the top-level containers that organize your tasks). Shows list names, IDs, and indicates default or shared lists.",
  {},
  async () => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return authFail()
      }

      const response = await graphGetAll<TaskList>(`${MS_GRAPH_BASE}/me/todo/lists?$top=100`, token)

      if (!response.ok) {
        return graphFail("Failed to retrieve task lists", response)
      }

      const lists = response.data.value || []
      if (lists.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: "No task lists found.",
            },
          ],
        }
      }

      const formattedLists = lists.map((list) => {
        // Add well-known list name if applicable
        let wellKnownInfo = ""
        if (list.wellknownListName && list.wellknownListName !== "none") {
          if (list.wellknownListName === "defaultList") {
            wellKnownInfo = " (Default Tasks List)"
          } else if (list.wellknownListName === "flaggedEmails") {
            wellKnownInfo = " (Flagged Emails)"
          }
        }

        // Add sharing info if applicable
        let sharingInfo = ""
        if (list.isShared) {
          sharingInfo = list.isOwner ? " (Shared by you)" : " (Shared with you)"
        }

        return `ID: ${list.id}\nName: ${list.displayName}${wellKnownInfo}${sharingInfo}\n---`
      })

      return {
        content: [
          {
            type: "text",
            text: `Your task lists:\n\n${formattedLists.join("\n")}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching task lists: ${error}`,
          },
        ],
      }
    }
  },
)

// Enhanced organized view of task lists
server.tool(
  "get-task-lists-organized",
  "Get all task lists organized into logical folders/categories based on naming patterns, emoji prefixes, and sharing status. Provides a hierarchical view similar to folder organization.",
  {
    includeIds: z.boolean().optional().describe("Include list IDs in output (default: false)"),
    groupBy: z
      .enum(["category", "shared", "type"])
      .optional()
      .describe("Grouping strategy - 'category' (default), 'shared', or 'type'"),
  },
  async ({ includeIds, groupBy }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return authFail()
      }

      const response = await graphGetAll<TaskList>(`${MS_GRAPH_BASE}/me/todo/lists?$top=100`, token)

      if (!response.ok) {
        return graphFail("Failed to retrieve task lists", response)
      }

      const lists = response.data.value || []
      if (lists.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: "No task lists found.",
            },
          ],
        }
      }

      // Group by shared status
      if (groupBy === "shared") {
        const sharedLists = lists.filter((l) => l.isShared)
        const personalLists = lists.filter((l) => !l.isShared)

        let output = "📂 Microsoft To Do Lists - By Sharing Status\n"
        output += "=".repeat(50) + "\n\n"

        output += `👥 Shared Lists (${sharedLists.length})\n`
        sharedLists.forEach((list) => {
          const ownership = list.isOwner ? "Shared by you" : "Shared with you"
          output += `   ├─ ${list.displayName} [${ownership}]\n`
        })

        output += `\n🔒 Personal Lists (${personalLists.length})\n`
        personalLists.forEach((list) => {
          output += `   ├─ ${list.displayName}\n`
        })

        return { content: [{ type: "text", text: output }] }
      }

      // Helper function to organize lists
      const organizeLists = (lists: TaskList[]): { [category: string]: TaskList[] } => {
        const organized: { [category: string]: TaskList[] } = {}

        // Patterns for categorizing lists
        const patterns = {
          archived: /\(([^)]+)\s*-\s*Archived\)$/i,
          archive: /^📦\s*Archive/i,
          shopping: /^🛒/,
          property: /^🏡/,
          family: /^👪/,
          seasonal: /^(🎄|🎉)/,
          work: /^(Work|SBIR)/i,
          travel: /^(🚗|Rangeley)/i,
          reading: /^📰/,
        }

        lists.forEach((list) => {
          // Check archived pattern
          const archiveMatch = list.displayName.match(patterns.archived)
          if (archiveMatch) {
            const category = `📦 Archived - ${archiveMatch[1]}`
            if (!organized[category]) organized[category] = []
            organized[category].push(list)
          }
          // Check archive prefix
          else if (patterns.archive.test(list.displayName)) {
            if (!organized["📦 Archives"]) organized["📦 Archives"] = []
            organized["📦 Archives"].push(list)
          }
          // Check shopping lists
          else if (patterns.shopping.test(list.displayName)) {
            if (!organized["🛒 Shopping Lists"]) organized["🛒 Shopping Lists"] = []
            organized["🛒 Shopping Lists"].push(list)
          }
          // Check property lists
          else if (patterns.property.test(list.displayName)) {
            if (!organized["🏡 Properties"]) organized["🏡 Properties"] = []
            organized["🏡 Properties"].push(list)
          }
          // Check family lists
          else if (patterns.family.test(list.displayName)) {
            if (!organized["👪 Family"]) organized["👪 Family"] = []
            organized["👪 Family"].push(list)
          }
          // Check seasonal lists
          else if (patterns.seasonal.test(list.displayName)) {
            if (!organized["🎉 Seasonal & Events"]) organized["🎉 Seasonal & Events"] = []
            organized["🎉 Seasonal & Events"].push(list)
          }
          // Check work lists
          else if (patterns.work.test(list.displayName)) {
            if (!organized["💼 Work"]) organized["💼 Work"] = []
            organized["💼 Work"].push(list)
          }
          // Check travel lists
          else if (patterns.travel.test(list.displayName)) {
            if (!organized["🚗 Travel & Rangeley"]) organized["🚗 Travel & Rangeley"] = []
            organized["🚗 Travel & Rangeley"].push(list)
          }
          // Check reading lists
          else if (patterns.reading.test(list.displayName)) {
            if (!organized["📚 Reading"]) organized["📚 Reading"] = []
            organized["📚 Reading"].push(list)
          }
          // Special lists
          else if (list.wellknownListName && list.wellknownListName !== "none") {
            if (!organized["⭐ Special Lists"]) organized["⭐ Special Lists"] = []
            organized["⭐ Special Lists"].push(list)
          }
          // Shared lists
          else if (list.isShared) {
            if (!organized["👥 Shared Lists"]) organized["👥 Shared Lists"] = []
            organized["👥 Shared Lists"].push(list)
          }
          // Everything else
          else {
            if (!organized["📋 Other Lists"]) organized["📋 Other Lists"] = []
            organized["📋 Other Lists"].push(list)
          }
        })

        return organized
      }

      // Default: organize by category
      const organized = organizeLists(lists)

      let output = "📂 Microsoft To Do Lists - Organized View\n"
      output += "=".repeat(50) + "\n\n"

      // Sort categories for consistent display
      const sortedCategories = Object.keys(organized).sort((a, b) => {
        // Priority order for categories
        const priority: { [key: string]: number } = {
          "⭐ Special Lists": 1,
          "👥 Shared Lists": 2,
          "💼 Work": 3,
          "👪 Family": 4,
          "🏡 Properties": 5,
          "🛒 Shopping Lists": 6,
          "🚗 Travel & Rangeley": 7,
          "🎉 Seasonal & Events": 8,
          "📚 Reading": 9,
          "📋 Other Lists": 10,
          "📦 Archives": 11,
        }

        // Check if categories start with "📦 Archived -"
        const aIsArchived = a.startsWith("📦 Archived -")
        const bIsArchived = b.startsWith("📦 Archived -")

        if (aIsArchived && !bIsArchived) return 1
        if (!aIsArchived && bIsArchived) return -1
        if (aIsArchived && bIsArchived) return a.localeCompare(b)

        const aPriority = priority[a] || 999
        const bPriority = priority[b] || 999

        if (aPriority !== bPriority) return aPriority - bPriority
        return a.localeCompare(b)
      })

      sortedCategories.forEach((category) => {
        const categoryLists = organized[category]
        output += `${category} (${categoryLists.length})\n`

        categoryLists.forEach((list, index) => {
          const isLast = index === categoryLists.length - 1
          const prefix = isLast ? "└─" : "├─"

          let listInfo = `${prefix} ${list.displayName}`

          // Add metadata
          const metadata: string[] = []
          if (list.wellknownListName === "defaultList") metadata.push("Default")
          if (list.wellknownListName === "flaggedEmails") metadata.push("Flagged Emails")
          if (list.isShared && list.isOwner) metadata.push("Shared by you")
          if (list.isShared && !list.isOwner) metadata.push("Shared with you")

          if (metadata.length > 0) {
            listInfo += ` [${metadata.join(", ")}]`
          }

          output += `   ${listInfo}\n`

          if (!isLast) {
            output += "   │\n"
          }
        })

        output += "\n"
      })

      // Add summary
      const totalLists = Object.values(organized).reduce((sum, l) => sum + l.length, 0)
      const totalCategories = Object.keys(organized).length

      output += "-".repeat(50) + "\n"
      output += `Summary: ${totalLists} lists in ${totalCategories} categories\n`

      if (includeIds) {
        // Add a section with IDs
        output += "\n\n📋 List IDs Reference:\n" + "-".repeat(50) + "\n"
        lists.forEach((list) => {
          output += `${list.displayName}: ${list.id}\n`
        })
      }

      return { content: [{ type: "text", text: output }] }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching organized task lists: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "create-task-list",
  "Create a new task list (top-level container) in Microsoft Todo to help organize your tasks into categories or projects.",
  {
    displayName: z.string().describe("Name of the new task list"),
  },
  async ({ displayName }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return authFail()
      }

      // Prepare the request body
      const requestBody = {
        displayName,
      }

      // Make the API request to create the task list
      const response = await graphRequest<TaskList>(`${MS_GRAPH_BASE}/me/todo/lists`, token, "POST", requestBody)

      if (!response.ok) {
        return graphFail(`Failed to create task list: ${displayName}`, response)
      }

      return {
        content: [
          {
            type: "text",
            text: `Task list created successfully!\nName: ${response.data.displayName}\nID: ${response.data.id}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error creating task list: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "update-task-list",
  "Update the name of an existing task list (top-level container) in Microsoft Todo.",
  {
    listId: z.string().describe("ID of the task list to update"),
    displayName: z.string().describe("New name for the task list"),
  },
  async ({ listId, displayName }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return authFail()
      }

      // Prepare the request body
      const requestBody = {
        displayName,
      }

      // Make the API request to update the task list
      const response = await graphRequest<TaskList>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}`,
        token,
        "PATCH",
        requestBody,
      )

      if (!response.ok) {
        return graphFail(`Failed to update task list with ID: ${listId}`, response)
      }

      return {
        content: [
          {
            type: "text",
            text: `Task list updated successfully!\nNew name: ${response.data.displayName}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error updating task list: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "delete-task-list",
  "Delete a task list (top-level container) from Microsoft Todo. This will remove the list and all tasks within it.",
  {
    listId: z.string().describe("ID of the task list to delete"),
  },
  async ({ listId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return authFail()
      }

      // Make a DELETE request to the Microsoft Graph API
      const url = `${MS_GRAPH_BASE}/me/todo/lists/${listId}`
      console.error(`Deleting task list: ${url}`)

      const result = await graphRequest(url, token, "DELETE")
      if (!result.ok) {
        return graphFail(`Failed to delete task list with ID: ${listId}`, result)
      }

      return {
        content: [
          {
            type: "text",
            text: `Task list with ID: ${listId} was successfully deleted.`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error deleting task list: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "get-tasks",
  "Get tasks from a specific Microsoft Todo list. These are the main todo items that can contain checklist items (subtasks).",
  {
    listId: z.string().describe("ID of the task list"),
    filter: z.string().optional().describe("OData $filter query (e.g., 'status eq \\'completed\\'')"),
    select: z.string().optional().describe("Comma-separated list of properties to include (e.g., 'id,title,status')"),
    orderby: z.string().optional().describe("Property to sort by (e.g., 'createdDateTime desc')"),
    top: z.number().optional().describe("Maximum number of tasks to retrieve"),
    skip: z.number().optional().describe("Number of tasks to skip"),
    count: z.boolean().optional().describe("Whether to include a count of tasks"),
  },
  async ({ listId, filter, select, orderby, top, skip, count }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return authFail()
      }

      // Build the query parameters
      const queryParams = new URLSearchParams()

      if (filter) queryParams.append("$filter", filter)
      if (select) queryParams.append("$select", select)
      if (orderby) queryParams.append("$orderby", orderby)
      if (top !== undefined) queryParams.append("$top", top.toString())
      if (skip !== undefined) queryParams.append("$skip", skip.toString())
      if (count !== undefined) queryParams.append("$count", count.toString())

      // Construct the URL with query parameters
      const queryString = queryParams.toString()
      const url = `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks${queryString ? "?" + queryString : ""}`

      console.error(`Making request to: ${url}`)

      const response = await graphRequest<{
        value: Task[]
        "@odata.count"?: number
        "@odata.nextLink"?: string
      }>(url, token)

      if (!response.ok) {
        return graphFail(`Failed to retrieve tasks for list: ${listId}`, response)
      }

      const tasks: Task[] = [...(response.data.value || [])]

      // Follow pagination so the caller gets the full list, not a truncated first
      // page. When `top` is set, treat it as a hard cap across pages.
      let nextLink = response.data["@odata.nextLink"]
      let guard = 0
      let truncated = false
      while (nextLink && (top === undefined || tasks.length < top)) {
        if (++guard > 100) {
          truncated = true
          break
        }
        const page = await graphRequest<{ value: Task[]; "@odata.nextLink"?: string }>(nextLink, token)
        if (!page.ok) {
          return graphFail(`Failed to retrieve tasks (page ${guard + 1}) for list: ${listId}`, page)
        }
        tasks.push(...(page.data.value || []))
        nextLink = page.data["@odata.nextLink"]
      }
      if (top !== undefined && tasks.length > top) tasks.length = top

      if (tasks.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: `No tasks found in list with ID: ${listId}`,
            },
          ],
        }
      }

      // Format the tasks based on available properties
      const formattedTasks = tasks.map((task) => {
        // Default format
        let taskInfo = `ID: ${task.id}\nTitle: ${task.title}`

        // Add status if available
        if (task.status) {
          const status = task.status === "completed" ? "✓" : "○"
          taskInfo = `${status} ${taskInfo}`
        }

        // Add due date if available
        if (task.dueDateTime) {
          taskInfo += `\nDue: ${new Date(task.dueDateTime.dateTime).toLocaleDateString()}`
        }

        // Add importance if available
        if (task.importance) {
          taskInfo += `\nImportance: ${task.importance}`
        }

        // Add categories if available
        if (task.categories && task.categories.length > 0) {
          taskInfo += `\nCategories: ${task.categories.join(", ")}`
        }

        // Add body content if available and not empty
        if (task.body && task.body.content && task.body.content.trim() !== "") {
          const previewLength = 50
          const contentPreview =
            task.body.content.length > previewLength
              ? task.body.content.substring(0, previewLength) + "..."
              : task.body.content
          taskInfo += `\nDescription: ${contentPreview}`
        }

        return `${taskInfo}\n---`
      })

      // Add count information if requested and available
      let countInfo = ""
      if (count && response.data["@odata.count"] !== undefined) {
        countInfo = `Total count: ${response.data["@odata.count"]}\n\n`
      }
      if (truncated) {
        countInfo += `⚠️ Results truncated at ${tasks.length} tasks (100-page pagination guard). Narrow with a $filter to see the rest.\n\n`
      }

      return {
        content: [
          {
            type: "text",
            text: `Tasks in list ${listId}:\n\n${countInfo}${formattedTasks.join("\n")}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching tasks: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "create-task",
  "Create a new task in a specific Microsoft Todo list. A task is the main todo item that can have a title, description, due date, and other properties.",
  {
    listId: z.string().describe("ID of the task list"),
    title: z.string().describe("Title of the task"),
    body: z.string().optional().describe("Description or body content of the task"),
    dueDateTime: z.string().optional().describe("Due date in ISO format (e.g., 2023-12-31T23:59:59Z)"),
    startDateTime: z.string().optional().describe("Start date in ISO format (e.g., 2023-12-31T23:59:59Z)"),
    importance: z.enum(["low", "normal", "high"]).optional().describe("Task importance"),
    isReminderOn: z.boolean().optional().describe("Whether to enable reminder for this task"),
    reminderDateTime: z.string().optional().describe("Reminder date and time in ISO format"),
    status: z
      .enum(["notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"])
      .optional()
      .describe("Status of the task"),
    categories: z.array(z.string()).optional().describe("Categories associated with the task"),
  },
  async ({
    listId,
    title,
    body,
    dueDateTime,
    startDateTime,
    importance,
    isReminderOn,
    reminderDateTime,
    status,
    categories,
  }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return authFail()
      }

      // Construct the task body with all supported properties
      const taskBody: any = { title }

      // Add optional properties if provided
      if (body) {
        taskBody.body = {
          content: body,
          contentType: "text",
        }
      }

      if (dueDateTime) {
        taskBody.dueDateTime = {
          dateTime: dueDateTime,
          timeZone: "UTC",
        }
      }

      if (startDateTime) {
        taskBody.startDateTime = {
          dateTime: startDateTime,
          timeZone: "UTC",
        }
      }

      if (importance) {
        taskBody.importance = importance
      }

      if (isReminderOn !== undefined) {
        taskBody.isReminderOn = isReminderOn
      }

      if (reminderDateTime) {
        taskBody.reminderDateTime = {
          dateTime: reminderDateTime,
          timeZone: "UTC",
        }
      }

      if (status) {
        taskBody.status = status
      }

      if (categories && categories.length > 0) {
        taskBody.categories = categories
      }

      const response = await graphRequest<Task>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks`,
        token,
        "POST",
        taskBody,
      )

      if (!response.ok) {
        return graphFail(`Failed to create task in list: ${listId}`, response)
      }

      return {
        content: [
          {
            type: "text",
            text: `Task created successfully!\nID: ${response.data.id}\nTitle: ${response.data.title}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error creating task: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "update-task",
  "Update an existing task in Microsoft Todo. Allows changing any properties of the task including title, due date, importance, etc.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task to update"),
    title: z.string().optional().describe("New title of the task"),
    body: z.string().optional().describe("New description or body content of the task"),
    dueDateTime: z.string().optional().describe("New due date in ISO format (e.g., 2023-12-31T23:59:59Z)"),
    startDateTime: z.string().optional().describe("New start date in ISO format (e.g., 2023-12-31T23:59:59Z)"),
    importance: z.enum(["low", "normal", "high"]).optional().describe("New task importance"),
    isReminderOn: z.boolean().optional().describe("Whether to enable reminder for this task"),
    reminderDateTime: z.string().optional().describe("New reminder date and time in ISO format"),
    status: z
      .enum(["notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"])
      .optional()
      .describe("New status of the task"),
    categories: z.array(z.string()).optional().describe("New categories associated with the task"),
  },
  async ({
    listId,
    taskId,
    title,
    body,
    dueDateTime,
    startDateTime,
    importance,
    isReminderOn,
    reminderDateTime,
    status,
    categories,
  }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return authFail()
      }

      // Construct the task update body with all provided properties
      const taskBody: any = {}

      // Add optional properties if provided
      if (title !== undefined) {
        taskBody.title = title
      }

      if (body !== undefined) {
        taskBody.body = {
          content: body,
          contentType: "text",
        }
      }

      if (dueDateTime !== undefined) {
        if (dueDateTime === "") {
          // Remove the due date by setting it to null
          taskBody.dueDateTime = null
        } else {
          taskBody.dueDateTime = {
            dateTime: dueDateTime,
            timeZone: "UTC",
          }
        }
      }

      if (startDateTime !== undefined) {
        if (startDateTime === "") {
          // Remove the start date by setting it to null
          taskBody.startDateTime = null
        } else {
          taskBody.startDateTime = {
            dateTime: startDateTime,
            timeZone: "UTC",
          }
        }
      }

      if (importance !== undefined) {
        taskBody.importance = importance
      }

      if (isReminderOn !== undefined) {
        taskBody.isReminderOn = isReminderOn
      }

      if (reminderDateTime !== undefined) {
        if (reminderDateTime === "") {
          // Remove the reminder date by setting it to null
          taskBody.reminderDateTime = null
        } else {
          taskBody.reminderDateTime = {
            dateTime: reminderDateTime,
            timeZone: "UTC",
          }
        }
      }

      if (status !== undefined) {
        taskBody.status = status
      }

      if (categories !== undefined) {
        taskBody.categories = categories
      }

      // Make sure we have at least one property to update
      if (Object.keys(taskBody).length === 0) {
        return {
          content: [
            {
              type: "text",
              text: "No properties provided for update. Please specify at least one property to change.",
            },
          ],
        }
      }

      const response = await graphRequest<Task>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`,
        token,
        "PATCH",
        taskBody,
      )

      if (!response.ok) {
        return graphFail(`Failed to update task with ID: ${taskId} in list: ${listId}`, response)
      }

      return {
        content: [
          {
            type: "text",
            text: `Task updated successfully!\nID: ${response.data.id}\nTitle: ${response.data.title}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error updating task: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "delete-task",
  "Delete a task from a Microsoft Todo list. This will remove the task and all its checklist items (subtasks).",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task to delete"),
  },
  async ({ listId, taskId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return authFail()
      }

      // Make a DELETE request to the Microsoft Graph API
      const url = `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`
      console.error(`Deleting task: ${url}`)

      const result = await graphRequest(url, token, "DELETE")
      if (!result.ok) {
        return graphFail(`Failed to delete task with ID: ${taskId} from list: ${listId}`, result)
      }

      return {
        content: [
          {
            type: "text",
            text: `Task with ID: ${taskId} was successfully deleted from list: ${listId}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error deleting task: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "get-checklist-items",
  "Get checklist items (subtasks) for a specific task. Checklist items are smaller steps or components that belong to a parent task.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
  },
  async ({ listId, taskId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return authFail()
      }

      // Fetch the task first to get its title
      const taskResponse = await graphRequest<Task>(`${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`, token)

      const taskTitle = taskResponse.ok ? taskResponse.data.title : "Unknown Task"

      // Fetch the checklist items
      const response = await graphRequest<{ value: ChecklistItem[] }>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems`,
        token,
      )

      if (!response.ok) {
        return graphFail(`Failed to retrieve checklist items for task: ${taskId}`, response)
      }

      const items = response.data.value || []
      if (items.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: `No checklist items found for task "${taskTitle}" (ID: ${taskId})`,
            },
          ],
        }
      }

      const formattedItems = items.map((item) => {
        const status = item.isChecked ? "✓" : "○"
        let itemInfo = `${status} ${item.displayName} (ID: ${item.id})`

        // Add creation date if available
        if (item.createdDateTime) {
          const createdDate = new Date(item.createdDateTime).toLocaleString()
          itemInfo += `\nCreated: ${createdDate}`
        }

        return itemInfo
      })

      return {
        content: [
          {
            type: "text",
            text: `Checklist items for task "${taskTitle}" (ID: ${taskId}):\n\n${formattedItems.join("\n\n")}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching checklist items: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "create-checklist-item",
  "Create a new checklist item (subtask) for a task. Checklist items help break down a task into smaller, manageable steps.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    displayName: z.string().describe("Text content of the checklist item"),
    isChecked: z.boolean().optional().describe("Whether the item is checked off"),
  },
  async ({ listId, taskId, displayName, isChecked }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return authFail()
      }

      // Prepare the request body
      const requestBody: any = {
        displayName,
      }

      if (isChecked !== undefined) {
        requestBody.isChecked = isChecked
      }

      // Make the API request to create the checklist item
      const response = await graphRequest<ChecklistItem>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems`,
        token,
        "POST",
        requestBody,
      )

      if (!response.ok) {
        return graphFail(`Failed to create checklist item for task: ${taskId}`, response)
      }

      return {
        content: [
          {
            type: "text",
            text: `Checklist item created successfully!\nContent: ${response.data.displayName}\nID: ${response.data.id}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error creating checklist item: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "update-checklist-item",
  "Update an existing checklist item (subtask). Allows changing the text content or completion status of the subtask.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    checklistItemId: z.string().describe("ID of the checklist item to update"),
    displayName: z.string().optional().describe("New text content of the checklist item"),
    isChecked: z.boolean().optional().describe("Whether the item is checked off"),
  },
  async ({ listId, taskId, checklistItemId, displayName, isChecked }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return authFail()
      }

      // Prepare the update body, including only the fields that are provided
      const requestBody: any = {}

      if (displayName !== undefined) {
        requestBody.displayName = displayName
      }

      if (isChecked !== undefined) {
        requestBody.isChecked = isChecked
      }

      // Make sure we have at least one property to update
      if (Object.keys(requestBody).length === 0) {
        return {
          content: [
            {
              type: "text",
              text: "No properties provided for update. Please specify either displayName or isChecked.",
            },
          ],
        }
      }

      // Make the API request to update the checklist item
      const response = await graphRequest<ChecklistItem>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems/${checklistItemId}`,
        token,
        "PATCH",
        requestBody,
      )

      if (!response.ok) {
        return graphFail(`Failed to update checklist item with ID: ${checklistItemId}`, response)
      }

      const statusText = response.data.isChecked ? "Checked" : "Not checked"

      return {
        content: [
          {
            type: "text",
            text: `Checklist item updated successfully!\nContent: ${response.data.displayName}\nStatus: ${statusText}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error updating checklist item: ${error}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "delete-checklist-item",
  "Delete a checklist item (subtask) from a task. This removes just the specific subtask, not the parent task.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    checklistItemId: z.string().describe("ID of the checklist item to delete"),
  },
  async ({ listId, taskId, checklistItemId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return authFail()
      }

      // Make a DELETE request to the Microsoft Graph API
      const url = `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems/${checklistItemId}`
      console.error(`Deleting checklist item: ${url}`)

      const result = await graphRequest(url, token, "DELETE")
      if (!result.ok) {
        return graphFail(`Failed to delete checklist item with ID: ${checklistItemId} from task: ${taskId}`, result)
      }

      return {
        content: [
          {
            type: "text",
            text: `Checklist item with ID: ${checklistItemId} was successfully deleted from task: ${taskId}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error deleting checklist item: ${error}`,
          },
        ],
      }
    }
  },
)

// Bulk archive completed tasks
server.tool(
  "archive-completed-tasks",
  "Move completed tasks older than a specified number of days from one list to another (archive) list. Useful for cleaning up active lists while preserving historical tasks.",
  {
    sourceListId: z.string().describe("ID of the source list to archive tasks from"),
    targetListId: z.string().describe("ID of the target archive list"),
    olderThanDays: z
      .number()
      .min(0)
      .default(90)
      .describe("Archive tasks completed more than this many days ago (default: 90)"),
    dryRun: z
      .boolean()
      .optional()
      .default(false)
      .describe("If true, only preview what would be archived without making changes"),
  },
  async ({ sourceListId, targetListId, olderThanDays, dryRun }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return authFail()
      }

      // Calculate cutoff date
      const cutoffDate = new Date()
      cutoffDate.setDate(cutoffDate.getDate() - olderThanDays)

      // Get all completed tasks from source list (paginated so nothing is missed).
      const tasksResponse = await graphGetAll<Task>(
        `${MS_GRAPH_BASE}/me/todo/lists/${sourceListId}/tasks?$filter=status eq 'completed'&$top=100`,
        token,
      )

      if (!tasksResponse.ok) {
        return graphFail("Failed to retrieve tasks from source list", tasksResponse)
      }

      // Filter tasks older than cutoff
      const tasksToArchive = tasksResponse.data.value.filter((task) => {
        if (!task.completedDateTime?.dateTime) return false
        const completedDate = new Date(task.completedDateTime.dateTime)
        return completedDate < cutoffDate
      })

      if (tasksToArchive.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: `No completed tasks found older than ${olderThanDays} days.`,
            },
          ],
        }
      }

      if (dryRun) {
        // Preview mode - just show what would be archived
        let preview = `📋 Archive Preview\n`
        preview += `Would archive ${tasksToArchive.length} tasks completed before ${cutoffDate.toLocaleDateString()}\n\n`

        tasksToArchive.forEach((task) => {
          const completedDate = task.completedDateTime?.dateTime
            ? new Date(task.completedDateTime.dateTime).toLocaleDateString()
            : "Unknown"
          preview += `- ${task.title} (completed: ${completedDate})\n`
        })

        return { content: [{ type: "text", text: preview }] }
      }

      // Copy each completed task into the target list, then delete the original —
      // but only after the copy is confirmed, so a failure never loses data.
      let successCount = 0
      const failedTasks: string[] = []

      for (const task of tasksToArchive) {
        const created = await graphRequest<Task>(`${MS_GRAPH_BASE}/me/todo/lists/${targetListId}/tasks`, token, "POST", {
          title: task.title,
          status: "completed",
          body: task.body,
          importance: task.importance,
          completedDateTime: task.completedDateTime,
          dueDateTime: task.dueDateTime,
          reminderDateTime: task.reminderDateTime,
          categories: task.categories,
        })
        if (!created.ok) {
          failedTasks.push(`${task.title} (copy failed: ${created.error})`)
          continue
        }

        const deleted = await graphRequest(
          `${MS_GRAPH_BASE}/me/todo/lists/${sourceListId}/tasks/${task.id}`,
          token,
          "DELETE",
        )
        if (!deleted.ok) {
          // Copy succeeded but original could not be removed — surface the duplicate
          // rather than pretending the move was clean.
          failedTasks.push(`${task.title} (copied OK but original NOT deleted — duplicate left: ${deleted.error})`)
          continue
        }
        successCount++
      }

      let result = `📦 Archive Complete\n`
      result += `Successfully archived ${successCount} of ${tasksToArchive.length} tasks\n`
      result += `Tasks completed before ${cutoffDate.toLocaleDateString()} were moved.\n`

      if (failedTasks.length > 0) {
        result += `\n⚠️ ${failedTasks.length} task(s) had problems:\n`
        failedTasks.forEach((title) => {
          result += `- ${title}\n`
        })
      }

      return { content: [{ type: "text", text: result }] }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error archiving tasks: ${error}`,
          },
        ],
      }
    }
  },
)

// Main function to start the server
export async function startServer(config?: ServerConfig): Promise<void> {
  try {
    // Note: Token management is now handled by the TokenManager class
    // Config options are kept for backward compatibility but not used

    // Check if using a personal Microsoft account and show warning if needed
    await isPersonalMicrosoftAccount()

    // Start the server
    const transport = new StdioServerTransport()
    await server.connect(transport)

    console.error("Server started and listening")
  } catch (error) {
    console.error("Error starting server:", error)
    throw error
  }
}

// Main entry point when executed directly
if (import.meta.url === `file://${process.argv[1]}`) {
  startServer().catch((error) => {
    console.error("Fatal error in main():", error)
    process.exit(1)
  })
}
