import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js"
import { z } from "zod"
import { readFileSync, writeFileSync, existsSync } from "fs"
import { join } from "path"
import dotenv from "dotenv"
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

// Helper function for making Microsoft Graph API requests
async function makeGraphRequest<T>(url: string, token: string, method = "GET", body?: any): Promise<T | null> {
  const headers = {
    "User-Agent": USER_AGENT,
    Accept: "application/json",
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
  }

  try {
    const options: RequestInit = {
      method,
      headers,
    }

    if (body && (method === "POST" || method === "PATCH")) {
      options.body = JSON.stringify(body)
    }

    console.error(`Making request to: ${url}`)
    console.error(
      `Request options: ${JSON.stringify({
        method,
        headers: {
          ...headers,
          Authorization: "Bearer [REDACTED]",
        },
      })}`,
    )

    let response = await fetch(url, options)

    // If we get a 401, try to refresh the token and retry once
    if (response.status === 401) {
      console.error("Got 401, attempting token refresh...")
      const newToken = await getAccessToken() // This will trigger refresh
      if (newToken && newToken !== token) {
        // Retry with new token
        headers.Authorization = `Bearer ${newToken}`
        response = await fetch(url, { ...options, headers })
      }
    }

    if (!response.ok) {
      const errorText = await response.text()
      console.error(`HTTP error! status: ${response.status}, body: ${errorText}`)

      // Check for the specific MailboxNotEnabledForRESTAPI error
      if (errorText.includes("MailboxNotEnabledForRESTAPI")) {
        console.error(`
=================================================================
ERROR: MailboxNotEnabledForRESTAPI

The Microsoft To Do API is not available for personal Microsoft accounts 
(outlook.com, hotmail.com, live.com, etc.) through the Graph API.

This is a limitation of the Microsoft Graph API, not an authentication issue.
Microsoft only allows To Do API access for Microsoft 365 business accounts.

You can still use Microsoft To Do through the web interface or mobile apps,
but API access is restricted for personal accounts.
=================================================================
        `)

        throw new Error(
          "Microsoft To Do API is not available for personal Microsoft accounts. See console for details.",
        )
      }

      throw new Error(`HTTP error! status: ${response.status}, body: ${errorText}`)
    }

    // Handle DELETE requests which may return empty responses
    if (method === "DELETE") {
      console.error(`DELETE request successful: ${response.status}`)
      return {} as T // Return empty object for successful DELETE
    }

    const data = await response.json()
    console.error(`Response received: ${JSON.stringify(data).substring(0, 200)}...`)
    return data as T
  } catch (error) {
    console.error("Error making Graph API request:", error)
    return null
  }
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

interface Task {
  id: string
  title: string
  status: string
  importance: string
  createdDateTime?: string
  lastModifiedDateTime?: string
  bodyLastModifiedDateTime?: string
  hasAttachments?: boolean
  dueDateTime?: {
    dateTime: string
    timeZone: string
  }
  startDateTime?: {
    dateTime: string
    timeZone: string
  }
  completedDateTime?: {
    dateTime: string
    timeZone: string
  }
  reminderDateTime?: {
    dateTime: string
    timeZone: string
  }
  body?: {
    content: string
    contentType: string
    lastModifiedDateTime?: string
  }
  categories?: string[]
  isReminderOn?: boolean
  recurrence?: {
    pattern: {
      type: string
      interval: number
      month?: number
      dayOfMonth?: number
      daysOfWeek?: string[]
      firstDayOfWeek?: string
      index?: string
    }
    range: {
      type: string
      startDate: string
      endDate?: string
      recurrenceTimeZone?: string
      numberOfOccurrences?: number
    }
  }
  linkedResources?: {
    webUrl?: string
    applicationName?: string
    displayName?: string
    externalId?: string
  }[]
}

interface TaskAttachment {
  id: string
  name: string
  size?: number
  contentType?: string
  lastModifiedDateTime?: string
  contentBytes?: string // Base64 encoded content
}

interface CreateTaskBody {
  title: string
  body?: {
    content: string
    contentType: string
  }
  dueDateTime?: {
    dateTime: string
    timeZone: string
  }
  startDateTime?: {
    dateTime: string
    timeZone: string
  }
  completedDateTime?: {
    dateTime: string
    timeZone: string
  }
  reminderDateTime?: {
    dateTime: string
    timeZone: string
  }
  importance?: string
  isReminderOn?: boolean
  status?: string
  categories?: string[]
  createdDateTime?: string
  lastModifiedDateTime?: string
  bodyLastModifiedDateTime?: string
  recurrence?: {
    pattern: {
      type: string
      interval: number
      month?: number
      dayOfMonth?: number
      daysOfWeek?: string[]
      firstDayOfWeek?: string
      index?: string
    }
    range: {
      type: string
      startDate: string
      endDate?: string
      recurrenceTimeZone?: string
      numberOfOccurrences?: number
    }
  }
  linkedResources?: {
    webUrl?: string
    applicationName?: string
    displayName?: string
    externalId?: string
  }[]
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
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      const response = await makeGraphRequest<{ value: TaskList[] }>(`${MS_GRAPH_BASE}/me/todo/lists`, token)

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to retrieve task lists",
            },
          ],
        }
      }

      const lists = response.value || []
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
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      const response = await makeGraphRequest<{ value: TaskList[] }>(`${MS_GRAPH_BASE}/me/todo/lists`, token)

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to retrieve task lists",
            },
          ],
        }
      }

      const lists = response.value || []
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
          let placed = false

          // Check archived pattern
          const archiveMatch = list.displayName.match(patterns.archived)
          if (archiveMatch) {
            const category = `📦 Archived - ${archiveMatch[1]}`
            if (!organized[category]) organized[category] = []
            organized[category].push(list)
            placed = true
          }
          // Check archive prefix
          else if (patterns.archive.test(list.displayName)) {
            if (!organized["📦 Archives"]) organized["📦 Archives"] = []
            organized["📦 Archives"].push(list)
            placed = true
          }
          // Check shopping lists
          else if (patterns.shopping.test(list.displayName)) {
            if (!organized["🛒 Shopping Lists"]) organized["🛒 Shopping Lists"] = []
            organized["🛒 Shopping Lists"].push(list)
            placed = true
          }
          // Check property lists
          else if (patterns.property.test(list.displayName)) {
            if (!organized["🏡 Properties"]) organized["🏡 Properties"] = []
            organized["🏡 Properties"].push(list)
            placed = true
          }
          // Check family lists
          else if (patterns.family.test(list.displayName)) {
            if (!organized["👪 Family"]) organized["👪 Family"] = []
            organized["👪 Family"].push(list)
            placed = true
          }
          // Check seasonal lists
          else if (patterns.seasonal.test(list.displayName)) {
            if (!organized["🎉 Seasonal & Events"]) organized["🎉 Seasonal & Events"] = []
            organized["🎉 Seasonal & Events"].push(list)
            placed = true
          }
          // Check work lists
          else if (patterns.work.test(list.displayName)) {
            if (!organized["💼 Work"]) organized["💼 Work"] = []
            organized["💼 Work"].push(list)
            placed = true
          }
          // Check travel lists
          else if (patterns.travel.test(list.displayName)) {
            if (!organized["🚗 Travel & Rangeley"]) organized["🚗 Travel & Rangeley"] = []
            organized["🚗 Travel & Rangeley"].push(list)
            placed = true
          }
          // Check reading lists
          else if (patterns.reading.test(list.displayName)) {
            if (!organized["📚 Reading"]) organized["📚 Reading"] = []
            organized["📚 Reading"].push(list)
            placed = true
          }
          // Special lists
          else if (list.wellknownListName && list.wellknownListName !== "none") {
            if (!organized["⭐ Special Lists"]) organized["⭐ Special Lists"] = []
            organized["⭐ Special Lists"].push(list)
            placed = true
          }
          // Shared lists (only if not already placed)
          else if (list.isShared && !placed) {
            if (!organized["👥 Shared Lists"]) organized["👥 Shared Lists"] = []
            organized["👥 Shared Lists"].push(list)
            placed = true
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
          const metadata = []
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
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Prepare the request body
      const requestBody = {
        displayName,
      }

      // Make the API request to create the task list
      const response = await makeGraphRequest<TaskList>(`${MS_GRAPH_BASE}/me/todo/lists`, token, "POST", requestBody)

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to create task list: ${displayName}`,
            },
          ],
        }
      }

      return {
        content: [
          {
            type: "text",
            text: `Task list created successfully!\nName: ${response.displayName}\nID: ${response.id}`,
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
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Prepare the request body
      const requestBody = {
        displayName,
      }

      // Make the API request to update the task list
      const response = await makeGraphRequest<TaskList>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}`,
        token,
        "PATCH",
        requestBody,
      )

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to update task list with ID: ${listId}`,
            },
          ],
        }
      }

      return {
        content: [
          {
            type: "text",
            text: `Task list updated successfully!\nNew name: ${response.displayName}`,
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
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Make a DELETE request to the Microsoft Graph API
      const url = `${MS_GRAPH_BASE}/me/todo/lists/${listId}`
      console.error(`Deleting task list: ${url}`)

      // The DELETE method doesn't return a response body, so we expect null
      await makeGraphRequest<null>(url, token, "DELETE")

      // If we get here, the delete was successful (204 No Content)
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
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
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

      const response = await makeGraphRequest<{ value: Task[]; "@odata.count"?: number }>(url, token)

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to retrieve tasks for list: ${listId}`,
            },
          ],
        }
      }

      const tasks = response.value || []
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
      if (count && response["@odata.count"] !== undefined) {
        countInfo = `Total count: ${response["@odata.count"]}\n\n`
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
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
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

      const response = await makeGraphRequest<Task>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks`,
        token,
        "POST",
        taskBody,
      )

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to create task in list: ${listId}`,
            },
          ],
        }
      }

      return {
        content: [
          {
            type: "text",
            text: `Task created successfully!\nID: ${response.id}\nTitle: ${response.title}`,
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
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
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

      const response = await makeGraphRequest<Task>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`,
        token,
        "PATCH",
        taskBody,
      )

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to update task with ID: ${taskId} in list: ${listId}`,
            },
          ],
        }
      }

      return {
        content: [
          {
            type: "text",
            text: `Task updated successfully!\nID: ${response.id}\nTitle: ${response.title}`,
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
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Make a DELETE request to the Microsoft Graph API
      const url = `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`
      console.error(`Deleting task: ${url}`)

      // The DELETE method doesn't return a response body, so we expect null
      await makeGraphRequest<null>(url, token, "DELETE")

      // If we get here, the delete was successful (204 No Content)
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
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Fetch the task first to get its title
      const taskResponse = await makeGraphRequest<Task>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`,
        token,
      )

      const taskTitle = taskResponse ? taskResponse.title : "Unknown Task"

      // Fetch the checklist items
      const response = await makeGraphRequest<{ value: ChecklistItem[] }>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems`,
        token,
      )

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to retrieve checklist items for task: ${taskId}`,
            },
          ],
        }
      }

      const items = response.value || []
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
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Prepare the request body
      const requestBody: any = {
        displayName,
      }

      if (isChecked !== undefined) {
        requestBody.isChecked = isChecked
      }

      // Make the API request to create the checklist item
      const response = await makeGraphRequest<ChecklistItem>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems`,
        token,
        "POST",
        requestBody,
      )

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to create checklist item for task: ${taskId}`,
            },
          ],
        }
      }

      return {
        content: [
          {
            type: "text",
            text: `Checklist item created successfully!\nContent: ${response.displayName}\nID: ${response.id}`,
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
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
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
      const response = await makeGraphRequest<ChecklistItem>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems/${checklistItemId}`,
        token,
        "PATCH",
        requestBody,
      )

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to update checklist item with ID: ${checklistItemId}`,
            },
          ],
        }
      }

      const statusText = response.isChecked ? "Checked" : "Not checked"

      return {
        content: [
          {
            type: "text",
            text: `Checklist item updated successfully!\nContent: ${response.displayName}\nStatus: ${statusText}`,
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
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Make a DELETE request to the Microsoft Graph API
      const url = `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems/${checklistItemId}`
      console.error(`Deleting checklist item: ${url}`)

      // The DELETE method doesn't return a response body, so we expect null
      await makeGraphRequest<null>(url, token, "DELETE")

      // If we get here, the delete was successful (204 No Content)
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
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Calculate cutoff date
      const cutoffDate = new Date()
      cutoffDate.setDate(cutoffDate.getDate() - olderThanDays)

      // Get all completed tasks from source list
      const tasksResponse = await makeGraphRequest<{ value: Task[] }>(
        `${MS_GRAPH_BASE}/me/todo/lists/${sourceListId}/tasks?$filter=status eq 'completed'`,
        token,
      )

      if (!tasksResponse || !tasksResponse.value) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to retrieve tasks from source list",
            },
          ],
        }
      }

      // Filter tasks older than cutoff
      const tasksToArchive = tasksResponse.value.filter((task) => {
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

      // Actually archive the tasks
      let successCount = 0
      let failedTasks: string[] = []

      for (const task of tasksToArchive) {
        try {
          // Create task in target list
          const createResponse = await makeGraphRequest(
            `${MS_GRAPH_BASE}/me/todo/lists/${targetListId}/tasks`,
            token,
            "POST",
            {
              title: task.title,
              status: "completed",
              body: task.body,
              importance: task.importance,
              completedDateTime: task.completedDateTime,
              dueDateTime: task.dueDateTime,
              reminderDateTime: task.reminderDateTime,
              categories: task.categories,
            },
          )

          if (createResponse) {
            // Delete from source list
            await makeGraphRequest(`${MS_GRAPH_BASE}/me/todo/lists/${sourceListId}/tasks/${task.id}`, token, "DELETE")
            successCount++
          } else {
            failedTasks.push(task.title)
          }
        } catch (error) {
          failedTasks.push(task.title)
        }
      }

      let result = `📦 Archive Complete\n`
      result += `Successfully archived ${successCount} of ${tasksToArchive.length} tasks\n`
      result += `Tasks completed before ${cutoffDate.toLocaleDateString()} were moved.\n`

      if (failedTasks.length > 0) {
        result += `\n⚠️ Failed to archive ${failedTasks.length} tasks:\n`
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

// Test tool to explore Graph API for hidden properties
server.tool(
  "test-graph-api-exploration",
  "Test various Graph API queries to discover hidden properties or endpoints for folder/group organization in Microsoft To Do.",
  {
    testType: z.enum(["odata-select", "odata-expand", "headers", "extensions", "all"]).describe("Type of test to run"),
  },
  async ({ testType }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      let results = "🔍 Graph API Exploration Results\n" + "=".repeat(50) + "\n\n"

      // Test 1: Try with $select=* to get all properties
      if (testType === "odata-select" || testType === "all") {
        results += "📊 Test 1: Using $select=* to retrieve all properties\n"
        try {
          const response = await makeGraphRequest<any>(`${MS_GRAPH_BASE}/me/todo/lists?$select=*`, token)
          if (response && response.value && response.value.length > 0) {
            const firstList = response.value[0]
            const properties = Object.keys(firstList)
            results += `Found ${properties.length} properties: ${properties.join(", ")}\n`

            // Show full first list as example
            results += "\nExample list object:\n"
            results += JSON.stringify(firstList, null, 2).substring(0, 1000) + "...\n"
          }
        } catch (error) {
          results += `Error: ${error}\n`
        }
        results += "\n"
      }

      // Test 2: Try various $expand options
      if (testType === "odata-expand" || testType === "all") {
        results += "📊 Test 2: Using $expand to retrieve related data\n"
        const expandOptions = [
          "extensions",
          "singleValueExtendedProperties",
          "multiValueExtendedProperties",
          "openExtensions",
          "parent",
          "children",
          "folder",
          "parentFolder",
          "group",
          "category",
        ]

        for (const expand of expandOptions) {
          try {
            const response = await makeGraphRequest<any>(
              `${MS_GRAPH_BASE}/me/todo/lists?$expand=${expand}&$top=1`,
              token,
            )
            if (response && response.value) {
              results += `✓ $expand=${expand}: Success - `
              if (response.value.length > 0 && response.value[0][expand]) {
                results += `Found data!\n`
                results += JSON.stringify(response.value[0][expand], null, 2).substring(0, 500) + "...\n"
              } else {
                results += `No additional data returned\n`
              }
            }
          } catch (error: any) {
            results += `✗ $expand=${expand}: ${error.message || "Failed"}\n`
          }
        }
        results += "\n"
      }

      // Test 3: Check response headers for additional info
      if (testType === "headers" || testType === "all") {
        results += "📊 Test 3: Checking response headers\n"
        try {
          const response = await fetch(`${MS_GRAPH_BASE}/me/todo/lists`, {
            headers: {
              Authorization: `Bearer ${token}`,
              Accept: "application/json",
              Prefer: "return=representation",
            },
          })

          results += "Response headers:\n"
          response.headers.forEach((value, key) => {
            results += `${key}: ${value}\n`
          })
        } catch (error) {
          results += `Error: ${error}\n`
        }
        results += "\n"
      }

      // Test 4: Try extensions endpoint
      if (testType === "extensions" || testType === "all") {
        results += "📊 Test 4: Checking for extensions\n"
        try {
          const listsResponse = await makeGraphRequest<{ value: TaskList[] }>(
            `${MS_GRAPH_BASE}/me/todo/lists?$top=1`,
            token,
          )

          if (listsResponse && listsResponse.value && listsResponse.value.length > 0) {
            const listId = listsResponse.value[0].id

            // Try to get extensions
            try {
              const extResponse = await makeGraphRequest<any>(
                `${MS_GRAPH_BASE}/me/todo/lists/${listId}/extensions`,
                token,
              )
              results += `Extensions found: ${JSON.stringify(extResponse, null, 2)}\n`
            } catch (error: any) {
              results += `No extensions endpoint: ${error.message}\n`
            }
          }
        } catch (error) {
          results += `Error: ${error}\n`
        }
        results += "\n"
      }

      // Test 5: Check if there's a separate folders or groups endpoint
      if (testType === "all") {
        results += "📊 Test 5: Checking for folder/group endpoints\n"
        const endpoints = [
          "/me/todo/folders",
          "/me/todo/groups",
          "/me/todo/listGroups",
          "/me/todo/listFolders",
          "/me/todo/categories",
        ]

        for (const endpoint of endpoints) {
          try {
            const response = await makeGraphRequest<any>(`${MS_GRAPH_BASE}${endpoint}`, token)
            results += `✓ ${endpoint}: Found! Response: ${JSON.stringify(response).substring(0, 200)}...\n`
          } catch (error: any) {
            results += `✗ ${endpoint}: Not found (${error.message || "Failed"})\n`
          }
        }
      }

      results += "\n" + "=".repeat(50) + "\n"
      results += "Analysis complete. Check results above for any discovered properties or endpoints."

      return {
        content: [
          {
            type: "text",
            text: results,
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error during Graph API exploration: ${error}`,
          },
        ],
      }
    }
  },
)

// Move task tool with full metadata and attachment preservation
server.tool(
  "move-task",
  "Move a task from one list to another, preserving checklist items and most metadata. Tasks with attachments cannot be moved. Creation timestamps cannot be preserved due to API limitations.",
  {
    sourceListId: z.string().describe("ID of the source task list"),
    sourceTaskId: z.string().describe("ID of the task to move"),
    targetListId: z.string().describe("ID of the target task list"),
  },
  async ({ sourceListId, sourceTaskId, targetListId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        }
      }

      // Get the original task with all details
      const originalTask = await makeGraphRequest<Task>(
        `${MS_GRAPH_BASE}/me/todo/lists/${sourceListId}/tasks/${sourceTaskId}`,
        token,
      )

      if (!originalTask) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to retrieve task: ${sourceTaskId}`,
            },
          ],
        }
      }

      // Check if task has attachments - reject if it does
      if (originalTask.hasAttachments === true) {
        return {
          content: [
            {
              type: "text",
              text: `Cannot move task "${originalTask.title}" because it has attachments. Tasks with attachments cannot be moved between lists.`,
            },
          ],
        }
      }

      // Get checklist items
      const checklistResponse = await makeGraphRequest<{ value: ChecklistItem[] }>(
        `${MS_GRAPH_BASE}/me/todo/lists/${sourceListId}/tasks/${sourceTaskId}/checklistItems`,
        token,
      )
      const checklistItems = checklistResponse?.value || []

      // Create the new task in the target list with all metadata
      const newTaskBody: CreateTaskBody = {
        title: originalTask.title,
      }

      // Copy all available properties, attempting to preserve timestamps
      if (originalTask.body) {
        newTaskBody.body = {
          content: originalTask.body.content,
          contentType: originalTask.body.contentType || "text",
        }
      }

      if (originalTask.dueDateTime) newTaskBody.dueDateTime = originalTask.dueDateTime
      if (originalTask.startDateTime) newTaskBody.startDateTime = originalTask.startDateTime
      if (originalTask.importance) newTaskBody.importance = originalTask.importance
      if (originalTask.isReminderOn !== undefined) newTaskBody.isReminderOn = originalTask.isReminderOn
      if (originalTask.reminderDateTime) newTaskBody.reminderDateTime = originalTask.reminderDateTime
      if (originalTask.status) newTaskBody.status = originalTask.status
      if (originalTask.categories) newTaskBody.categories = originalTask.categories
      if (originalTask.recurrence) newTaskBody.recurrence = originalTask.recurrence
      if (originalTask.linkedResources) newTaskBody.linkedResources = originalTask.linkedResources

      // Attempt to preserve creation date (may not work due to API limitations)
      if (originalTask.createdDateTime) {
        newTaskBody.createdDateTime = originalTask.createdDateTime
      }

      // Attempt to preserve modification dates
      if (originalTask.lastModifiedDateTime) {
        newTaskBody.lastModifiedDateTime = originalTask.lastModifiedDateTime
      }

      if (originalTask.bodyLastModifiedDateTime) {
        newTaskBody.bodyLastModifiedDateTime = originalTask.bodyLastModifiedDateTime
      }

      // Handle completed date
      if (originalTask.completedDateTime) {
        newTaskBody.completedDateTime = originalTask.completedDateTime
      }

      const newTask = await makeGraphRequest<Task>(
        `${MS_GRAPH_BASE}/me/todo/lists/${targetListId}/tasks`,
        token,
        "POST",
        newTaskBody,
      )

      if (!newTask) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to create task in target list",
            },
          ],
        }
      }

      // Copy checklist items to the new task
      for (const item of checklistItems) {
        await makeGraphRequest(
          `${MS_GRAPH_BASE}/me/todo/lists/${targetListId}/tasks/${newTask.id}/checklistItems`,
          token,
          "POST",
          {
            displayName: item.displayName,
            isChecked: item.isChecked,
          },
        )
      }

      // Delete the original task
      const deleteResult = await makeGraphRequest(
        `${MS_GRAPH_BASE}/me/todo/lists/${sourceListId}/tasks/${sourceTaskId}`,
        token,
        "DELETE",
      )

      const checklistCount = checklistItems.length

      // Check if timestamps were actually preserved
      const timestampsPreserved =
        newTask.createdDateTime === originalTask.createdDateTime &&
        newTask.lastModifiedDateTime === originalTask.lastModifiedDateTime

      return {
        content: [
          {
            type: "text",
            text:
              `✅ Successfully moved task "${originalTask.title}"\n\n` +
              `📋 Details:\n` +
              `• Task ID: ${newTask.id}\n` +
              `• Checklist items moved: ${checklistCount}\n` +
              `• Original task deleted: ${deleteResult !== null ? "✅ Yes" : "❌ No"}\n` +
              `• Timestamps preserved: ${timestampsPreserved ? "✅ Yes" : "❌ No"}\n\n` +
              (timestampsPreserved
                ? `✅ All metadata successfully preserved!`
                : `⚠️ Note: Timestamps were not preserved. Original creation: ${originalTask.createdDateTime}, New creation: ${newTask.createdDateTime}.`),
          },
        ],
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error moving task: ${error}`,
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
