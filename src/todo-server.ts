// The MCP server and its Microsoft To Do tools, built for a given credential source. The
// stdio entry point (todo-index.ts) passes the token manager's; the HTTP entry point passes
// each signed-in caller's own.
import { readFileSync } from "fs"
import { IO, Option } from "functype"
import { dirname, join } from "path"
import type { AuthProvider, SessionAuth } from "somamcp"
import { createServer } from "somamcp/node"
import { fileURLToPath } from "url"
import { z } from "zod"

import type { CredentialSource } from "./credentials.js"
import { graphErrorText, graphRequest, graphRequestAll, MS_GRAPH_BASE, runOrNull } from "./graph-client.js"

// The published version, read from package.json (one level up from dist/ and src/), so
// serverInfo reports what is installed.
const PACKAGE_VERSION: `${number}.${number}.${number}` = JSON.parse(
  readFileSync(join(dirname(fileURLToPath(import.meta.url)), "..", "package.json"), "utf8"),
).version

type ToolReply = { content: Array<{ type: "text"; text: string }> }
const reply = (text: string): ToolReply => ({ content: [{ type: "text", text }] })
const NOT_AUTHENTICATED = "Failed to authenticate with Microsoft API"

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
  dueDateTime?: {
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
  }
  categories?: string[]
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
      startDate?: string
      endDate?: string
      numberOfOccurrences?: number
      recurrenceTimeZone?: string
    }
  }
}

// Zod schemas for recurrence input. The MCP exposes a friendly nested shape;
// see https://learn.microsoft.com/en-us/graph/api/resources/recurrencepattern
const recurrencePatternSchema = z.object({
  type: z
    .enum(["daily", "weekly", "absoluteMonthly", "relativeMonthly", "absoluteYearly", "relativeYearly"])
    .describe(
      "Recurrence pattern type. daily | weekly | absoluteMonthly (e.g. 15th of every N months) | relativeMonthly (e.g. 2nd Tuesday of every N months) | absoluteYearly | relativeYearly.",
    ),
  interval: z
    .number()
    .int()
    .min(1)
    .max(255)
    .describe("How often the pattern repeats (1 = every period, 2 = every second, ...)."),
  month: z.number().int().min(1).max(12).optional().describe("Month for absoluteYearly / relativeYearly (1-12)."),
  dayOfMonth: z.number().int().min(1).max(31).optional().describe("Day of month for absoluteMonthly / absoluteYearly."),
  daysOfWeek: z
    .array(z.enum(["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"]))
    .optional()
    .describe("Days of week for weekly / relativeMonthly / relativeYearly."),
  firstDayOfWeek: z
    .enum(["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"])
    .optional()
    .describe("Which day is the start of the week (defaults to sunday)."),
  index: z
    .enum(["first", "second", "third", "fourth", "last"])
    .optional()
    .describe("Ordinal position for relativeMonthly / relativeYearly (e.g. 'second' Tuesday)."),
})

const recurrenceRangeSchema = z.object({
  type: z
    .enum(["endDate", "noEnd", "numbered"])
    .describe(
      "How long the recurrence repeats: endDate (stops at endDate), noEnd (forever), numbered (stops after N occurrences).",
    ),
  startDate: z
    .string()
    .optional()
    .describe("Required for endDate/numbered; ISO date YYYY-MM-DD. Defaults to the task due date if omitted."),
  endDate: z.string().optional().describe("Required when type=endDate; ISO date YYYY-MM-DD."),
  numberOfOccurrences: z
    .number()
    .int()
    .min(1)
    .optional()
    .describe("Required when type=numbered; how many occurrences before stopping."),
  recurrenceTimeZone: z
    .string()
    .optional()
    .describe(
      "IANA timezone for the recurrence. In one live test Microsoft To Do stored 'Asia/Kolkata' as UTC, and it keeps only the due date's day, not its time.",
    ),
})

const recurrenceSchema = z
  .object({
    pattern: recurrencePatternSchema,
    range: recurrenceRangeSchema,
  })
  .describe(
    "Recurrence rule. Omit to leave the task one-off. To stop an existing recurring task, set recurrence to null via update-task.",
  )

type RecurrenceInput = z.infer<typeof recurrenceSchema>

// Which optional fields each pattern and range type requires and accepts, per
// https://learn.microsoft.com/en-us/graph/api/resources/recurrencepattern and
// .../recurrencerange. Graph answers a missing required field with a 400 that names no
// field, so the server checks before sending. Against a live account, Graph silently
// dropped a field the type does not use (dayOfMonth on weekly, firstDayOfWeek on
// relativeMonthly) instead of the documented error, so those are refused too: the
// caller would otherwise believe the field took effect.
const PATTERN_FIELDS: Record<string, { required: string[]; allowed: string[] }> = {
  daily: { required: [], allowed: [] },
  weekly: { required: ["daysOfWeek"], allowed: ["daysOfWeek", "firstDayOfWeek"] },
  absoluteMonthly: { required: ["dayOfMonth"], allowed: ["dayOfMonth"] },
  relativeMonthly: { required: ["daysOfWeek"], allowed: ["daysOfWeek", "index"] },
  absoluteYearly: { required: ["dayOfMonth", "month"], allowed: ["dayOfMonth", "month"] },
  relativeYearly: { required: ["daysOfWeek", "month"], allowed: ["daysOfWeek", "month", "index"] },
}
const RANGE_FIELDS: Record<string, { required: string[]; allowed: string[] }> = {
  noEnd: { required: [], allowed: [] },
  endDate: { required: ["endDate"], allowed: ["endDate"] },
  numbered: { required: ["numberOfOccurrences"], allowed: ["numberOfOccurrences"] },
}
const PATTERN_OPTIONAL = ["month", "dayOfMonth", "daysOfWeek", "firstDayOfWeek", "index"] as const
const RANGE_OPTIONAL = ["endDate", "numberOfOccurrences"] as const

// The first rule the recurrence breaks, as a message for the caller, or null if none.
function recurrenceProblem({ pattern, range }: RecurrenceInput): string | null {
  const check = (
    kind: string,
    type: string,
    value: Record<string, unknown>,
    fields: { required: string[]; allowed: string[] },
    optional: readonly string[],
  ): string | null => {
    const missing = fields.required.filter(
      (f) => value[f] === undefined || (Array.isArray(value[f]) && !value[f].length),
    )
    if (missing.length) return `A ${type} ${kind} needs ${missing.join(" and ")}.`
    const extra = optional.filter((f) => value[f] !== undefined && !fields.allowed.includes(f))
    if (extra.length) return `A ${type} ${kind} does not use ${extra.join(" or ")}; leave it out.`
    return null
  }
  return (
    check("pattern", pattern.type, pattern, PATTERN_FIELDS[pattern.type], PATTERN_OPTIONAL) ??
    check("range", range.type, range, RANGE_FIELDS[range.type], RANGE_OPTIONAL)
  )
}

// The Graph body for a checked recurrence. range.startDate defaults to the due date's
// calendar day. A weekly pattern gets firstDayOfWeek, which the docs list as required
// for weekly; Graph defaulted it to sunday in live testing, so this matches either way.
function toGraphRecurrence({ pattern, range }: RecurrenceInput, dueISO: string): RecurrenceInput {
  return {
    pattern: pattern.type === "weekly" ? { firstDayOfWeek: "sunday", ...pattern } : { ...pattern },
    range: { ...range, startDate: range.startDate ?? dueISO.substring(0, 10) },
  }
}

// A stored recurrence as one readable line, e.g. "every 2 weekly on friday (range: numbered, 5x, ends 2027-01-15)".
// Graph fills every field on read (index "first", endDate "0001-01-01", zeros), so only the
// fields the pattern type uses are shown.
function formatRecurrence({ pattern, range }: NonNullable<Task["recurrence"]>): string {
  const every = pattern.interval > 1 ? `every ${pattern.interval}` : "every"
  let text = `${every} ${pattern.type
    .replace(/([A-Z])/g, " $1")
    .trim()
    .toLowerCase()}`
  const relative = pattern.type === "relativeMonthly" || pattern.type === "relativeYearly"
  if ((pattern.type === "weekly" || relative) && pattern.daysOfWeek?.length) {
    text += ` on ${relative && pattern.index ? `the ${pattern.index} ` : ""}${pattern.daysOfWeek.join("/")}`
  }
  if ((pattern.type === "absoluteMonthly" || pattern.type === "absoluteYearly") && pattern.dayOfMonth) {
    text += ` on day ${pattern.dayOfMonth}`
  }
  if ((pattern.type === "absoluteYearly" || pattern.type === "relativeYearly") && pattern.month) {
    text += ` of month ${pattern.month}`
  }
  let rangeText = range.type
  if (range.type === "endDate" && range.endDate) rangeText += ` until ${range.endDate}`
  if (range.type === "numbered" && range.numberOfOccurrences) {
    rangeText += `, ${range.numberOfOccurrences}x`
    // Graph computes the last occurrence's date for a numbered range; show it when present.
    if (range.endDate && range.endDate !== "0001-01-01") rangeText += `, ends ${range.endDate}`
  }
  return `${text} (range: ${rangeText})`
}

interface ChecklistItem {
  id: string
  displayName: string
  isChecked: boolean
  createdDateTime?: string
}

// The "List IDs Reference" section get-task-lists-organized appends with includeIds.
const listIdsReference = (lists: TaskList[]): string =>
  "\n\n📋 List IDs Reference:\n" +
  "-".repeat(50) +
  "\n" +
  lists.map((list) => `${list.displayName}: ${list.id}\n`).join("")

export type TodoServerOptions<S extends SessionAuth> = {
  /** Where each tool call gets its Graph token. */
  readonly credentials: CredentialSource<S>
  /** Authentication for the HTTP endpoint, such as somamcp's Azure provider. */
  readonly auth?: AuthProvider<S>
}

/** A server with every Microsoft To Do tool, reading Graph tokens from `credentials`. */
export function createTodoServer<S extends SessionAuth>({ credentials, auth }: TodoServerOptions<S>) {
  const server = createServer<S>({
    name: "mstodo",
    version: PACKAGE_VERSION,
    ...(auth ? { auth } : {}),
  })

  // Server tool to check authentication status
  server.addTool({
    name: "auth-status",
    description:
      "Check if you're authenticated with Microsoft Graph API. Shows current token status and expiration time, and indicates if the token needs to be refreshed.",
    parameters: z.object({}),
    execute: async (_args, context) => reply(await credentials.status(context)),
  })

  // Register tools
  server.addTool({
    name: "get-task-lists",
    description:
      "Get all Microsoft Todo task lists (the top-level containers that organize your tasks). Shows list names, IDs, and indicates default or shared lists.",
    parameters: z.object({}),
    execute: async (_args, context) => {
      try {
        const token = (await credentials.graphAuth(context)).orNull()
        if (token === null) return reply(NOT_AUTHENTICATED)

        const fetched = await graphRequestAll<TaskList>(`${MS_GRAPH_BASE}/me/todo/lists`, token).run()
        if (fetched.isLeft()) return reply("Failed to retrieve task lists")

        const lists = fetched.value
        if (lists.length === 0) {
          return reply("No task lists found.")
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

        return reply(`Your task lists:\n\n${formattedLists.join("\n")}`)
      } catch (error) {
        return reply(`Error fetching task lists: ${error}`)
      }
    },
  })

  // Enhanced organized view of task lists
  server.addTool({
    name: "get-task-lists-organized",
    description:
      "Get all task lists organized into logical folders/categories based on naming patterns, emoji prefixes, and sharing status. Provides a hierarchical view similar to folder organization.",
    parameters: z.object({
      includeIds: z.boolean().optional().describe("Include list IDs in output (default: false)"),
      groupBy: z
        .enum(["category", "shared", "type"])
        .optional()
        .describe("Grouping strategy - 'category' (default), 'shared', or 'type'"),
    }),
    execute: async ({ includeIds, groupBy }, context) => {
      try {
        const token = (await credentials.graphAuth(context)).orNull()
        if (token === null) return reply(NOT_AUTHENTICATED)

        const fetched = await graphRequestAll<TaskList>(`${MS_GRAPH_BASE}/me/todo/lists`, token).run()
        if (fetched.isLeft()) return reply("Failed to retrieve task lists")

        const lists = fetched.value
        if (lists.length === 0) {
          return reply("No task lists found.")
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

          if (includeIds) output += listIdsReference(lists)

          return reply(output)
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

        if (includeIds) output += listIdsReference(lists)

        return reply(output)
      } catch (error) {
        return reply(`Error fetching organized task lists: ${error}`)
      }
    },
  })

  server.addTool({
    name: "create-task-list",
    description:
      "Create a new task list (top-level container) in Microsoft Todo to help organize your tasks into categories or projects.",
    parameters: z.object({
      displayName: z.string().describe("Name of the new task list"),
    }),
    execute: async ({ displayName }, context) => {
      try {
        const token = (await credentials.graphAuth(context)).orNull()
        if (token === null) return reply(NOT_AUTHENTICATED)

        return (
          await graphRequest<TaskList>(`${MS_GRAPH_BASE}/me/todo/lists`, token, "POST", { displayName }).run()
        ).fold(
          () => reply(`Failed to create task list: ${displayName}`),
          (list) => reply(`Task list created successfully!\nName: ${list.displayName}\nID: ${list.id}`),
        )
      } catch (error) {
        return reply(`Error creating task list: ${error}`)
      }
    },
  })

  server.addTool({
    name: "update-task-list",
    description: "Update the name of an existing task list (top-level container) in Microsoft Todo.",
    parameters: z.object({
      listId: z.string().describe("ID of the task list to update"),
      displayName: z.string().describe("New name for the task list"),
    }),
    execute: async ({ listId, displayName }, context) => {
      try {
        const token = (await credentials.graphAuth(context)).orNull()
        if (token === null) return reply(NOT_AUTHENTICATED)

        return (
          await graphRequest<TaskList>(`${MS_GRAPH_BASE}/me/todo/lists/${listId}`, token, "PATCH", {
            displayName,
          }).run()
        ).fold(
          () => reply(`Failed to update task list with ID: ${listId}`),
          (list) => reply(`Task list updated successfully!\nNew name: ${list.displayName}`),
        )
      } catch (error) {
        return reply(`Error updating task list: ${error}`)
      }
    },
  })

  server.addTool({
    name: "delete-task-list",
    description:
      "Delete a task list (top-level container) from Microsoft Todo. This will remove the list and all tasks within it.",
    parameters: z.object({
      listId: z.string().describe("ID of the task list to delete"),
    }),
    execute: async ({ listId }, context) => {
      try {
        const token = (await credentials.graphAuth(context)).orNull()
        if (token === null) return reply(NOT_AUTHENTICATED)

        // Make a DELETE request to the Microsoft Graph API
        const url = `${MS_GRAPH_BASE}/me/todo/lists/${listId}`
        console.error(`Deleting task list: ${url}`)

        return (await graphRequest<void>(url, token, "DELETE").run()).fold(
          (error) => reply(`Failed to delete task list with ID: ${listId}: ${graphErrorText(error)}`),
          () => reply(`Task list with ID: ${listId} was successfully deleted.`),
        )
      } catch (error) {
        return reply(`Error deleting task list: ${error}`)
      }
    },
  })

  server.addTool({
    name: "get-tasks",
    description:
      "Get tasks from a specific Microsoft Todo list. These are the main todo items that can contain checklist items (subtasks).",
    parameters: z.object({
      listId: z.string().describe("ID of the task list"),
      filter: z.string().optional().describe("OData $filter query (e.g., 'status eq \\'completed\\'')"),
      select: z
        .string()
        .optional()
        .describe(
          "Comma-separated subset of the shown fields to include: status, dueDateTime, importance, categories, recurrence, body (e.g., 'status,dueDateTime'). id and title are always included; other names are ignored.",
        ),
      orderby: z
        .string()
        .optional()
        .describe("OData $orderby, e.g. 'createdDateTime desc'. Graph rejects 'title' on To Do tasks."),
      top: z.number().optional().describe("Maximum number of tasks to retrieve"),
      skip: z.number().optional().describe("Number of tasks to skip"),
      count: z
        .boolean()
        .optional()
        .describe(
          "Print how many tasks this call returned. Graph ignores $count on To Do tasks, so it is not a total for the list.",
        ),
    }),
    execute: async ({ listId, filter, select, orderby, top, skip, count }, context) => {
      try {
        const token = (await credentials.graphAuth(context)).orNull()
        if (token === null) return reply(NOT_AUTHENTICATED)

        // Build the query parameters
        const queryParams = new URLSearchParams()

        if (filter) queryParams.append("$filter", filter)
        // Deliberately not forwarded as $select: /me/todo/lists/{id}/tasks answers any
        // $select with HTTP 400 (reproduced with and without $filter and $top, for every
        // field combination tried, against a personal account). Forwarding it turned a
        // valid call into a hard failure, so the parameter is applied as a client-side
        // projection below instead — the caller still gets the narrower result it asked for.
        if (orderby) queryParams.append("$orderby", orderby)
        if (top !== undefined) queryParams.append("$top", top.toString())
        if (skip !== undefined) queryParams.append("$skip", skip.toString())
        if (count !== undefined) queryParams.append("$count", count.toString())

        // Construct the URL with query parameters
        const queryString = queryParams.toString()
        const url = `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks${queryString ? "?" + queryString : ""}`

        const fetched = await graphRequestAll<Task>(url, token, top).run()
        if (fetched.isLeft()) {
          return reply(`Failed to retrieve tasks for list: ${listId} | Graph said: ${graphErrorText(fetched.value)}`)
        }

        const tasks = fetched.value
        if (tasks.length === 0) {
          return reply(`No tasks found in list with ID: ${listId}`)
        }

        // Client-side stand-in for $select (see above). No list means everything. ID and Title
        // always print; status, dueDateTime, importance, categories, recurrence and body print only when
        // named. Any other name is ignored.
        const wanted = select
          ? new Set(
              select
                .split(",")
                .map((f) => f.trim())
                .filter(Boolean),
            )
          : null
        const want = (field: string) => !wanted || wanted.has(field)

        // Format the tasks based on available properties
        const formattedTasks = tasks.map((task) => {
          // Default format
          let taskInfo = `ID: ${task.id}\nTitle: ${task.title}`

          // Add status if available
          if (want("status") && task.status) {
            const status = task.status === "completed" ? "✓" : "○"
            taskInfo = `${status} ${taskInfo}`
          }

          // Add due date if available
          if (want("dueDateTime") && task.dueDateTime) {
            taskInfo += `\nDue: ${new Date(task.dueDateTime.dateTime).toLocaleDateString()}`
          }

          // Add importance if available
          if (want("importance") && task.importance) {
            taskInfo += `\nImportance: ${task.importance}`
          }

          // Add categories if available
          if (want("categories") && task.categories && task.categories.length > 0) {
            taskInfo += `\nCategories: ${task.categories.join(", ")}`
          }

          // Add recurrence if available
          if (want("recurrence") && task.recurrence) {
            taskInfo += `\nRepeats: ${formatRecurrence(task.recurrence)}`
          }

          // Add body content if available and not empty
          if (want("body") && task.body && task.body.content && task.body.content.trim() !== "") {
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
        if (count) {
          countInfo = `Total count: ${tasks.length}\n\n`
        }

        return reply(`Tasks in list ${listId}:\n\n${countInfo}${formattedTasks.join("\n")}`)
      } catch (error) {
        return reply(`Error fetching tasks: ${error}`)
      }
    },
  })

  server.addTool({
    name: "create-task",
    description:
      "Create a new task in a specific Microsoft Todo list. A task is the main todo item that can have a title, description, due date, and other properties. Pass `recurrence` to make the task repeat (e.g. monthly bill payments, birthdays, weekly reviews).",
    parameters: z.object({
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
      recurrence: recurrenceSchema
        .optional()
        .describe(
          "Recurrence rule. Omit for a one-off task. Requires `dueDateTime` to be set (the recurrence derives its first occurrence from it).",
        ),
    }),
    execute: async (
      {
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
        recurrence,
      },
      context,
    ) => {
      try {
        const token = (await credentials.graphAuth(context)).orNull()
        if (token === null) return reply(NOT_AUTHENTICATED)

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

        if (recurrence) {
          const problem = recurrenceProblem(recurrence)
          if (problem) {
            return reply(`Recurrence not set: ${problem}`)
          }
          if (!dueDateTime) {
            return reply(
              "Cannot create a recurring task without a dueDateTime — Microsoft Graph derives the first occurrence from the due date.",
            )
          }
          taskBody.recurrence = toGraphRecurrence(recurrence, dueDateTime)
        }

        return (
          await graphRequest<Task>(`${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks`, token, "POST", taskBody).run()
        ).fold(
          (error) => reply(`Failed to create task in list: ${listId}: ${graphErrorText(error)}`),
          (created) => {
            const recurrenceInfo = created.recurrence ? `\nRepeats: ${formatRecurrence(created.recurrence)}` : ""
            return reply(`Task created successfully!\nID: ${created.id}\nTitle: ${created.title}${recurrenceInfo}`)
          },
        )
      } catch (error) {
        return reply(`Error creating task: ${error}`)
      }
    },
  })

  server.addTool({
    name: "update-task",
    description:
      "Update an existing task in Microsoft Todo. Allows changing any properties of the task including title, due date, importance, and recurrence. To stop a recurring task, pass `recurrence: null`. Pass a new `recurrence` object to change or add recurrence.",
    parameters: z.object({
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
      recurrence: z
        .union([recurrenceSchema, z.null()])
        .optional()
        .describe(
          "New recurrence rule for the task. Pass an object to set/change recurrence. Pass `null` to make the task one-off (clears any existing recurrence). Omit to leave recurrence untouched.",
        ),
    }),
    execute: async (
      {
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
        recurrence,
      },
      context,
    ) => {
      try {
        const token = (await credentials.graphAuth(context)).orNull()
        if (token === null) return reply(NOT_AUTHENTICATED)

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

        if (recurrence !== undefined) {
          if (recurrence === null) {
            // Caller asked to clear any existing recurrence
            taskBody.recurrence = null
          } else {
            const problem = recurrenceProblem(recurrence)
            if (problem) {
              return reply(`Recurrence not set: ${problem}`)
            }
            // Graph derives the first occurrence from the due date, and range.startDate
            // defaults to it. Use the due date this call sets, or else the one the task
            // already has. Clearing the due date in the same call leaves none.
            let dueISO: string | undefined
            if (dueDateTime !== undefined) {
              dueISO = dueDateTime || undefined
            } else {
              const existing = await graphRequest<Task>(
                `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`,
                token,
              ).run()
              if (existing.isLeft()) {
                return reply(`Failed to read the task's current due date: ${graphErrorText(existing.value)}`)
              }
              dueISO = existing.value.dueDateTime?.dateTime
            }
            if (!dueISO) {
              return reply(
                "Cannot set recurrence without a dueDateTime — pass dueDateTime in this update or set one on the task first.",
              )
            }
            taskBody.recurrence = toGraphRecurrence(recurrence, dueISO)
          }
        }

        // Make sure we have at least one property to update
        if (Object.keys(taskBody).length === 0) {
          return reply("No properties provided for update. Please specify at least one property to change.")
        }

        return (
          await graphRequest<Task>(
            `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`,
            token,
            "PATCH",
            taskBody,
          ).run()
        ).fold(
          (error) => reply(`Failed to update task with ID: ${taskId} in list: ${listId}: ${graphErrorText(error)}`),
          (updated) => reply(`Task updated successfully!\nID: ${updated.id}\nTitle: ${updated.title}`),
        )
      } catch (error) {
        return reply(`Error updating task: ${error}`)
      }
    },
  })

  server.addTool({
    name: "delete-task",
    description:
      "Delete a task from a Microsoft Todo list. This will remove the task and all its checklist items (subtasks).",
    parameters: z.object({
      listId: z.string().describe("ID of the task list"),
      taskId: z.string().describe("ID of the task to delete"),
    }),
    execute: async ({ listId, taskId }, context) => {
      try {
        const token = (await credentials.graphAuth(context)).orNull()
        if (token === null) return reply(NOT_AUTHENTICATED)

        // Make a DELETE request to the Microsoft Graph API
        const url = `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`
        console.error(`Deleting task: ${url}`)

        return (await graphRequest<void>(url, token, "DELETE").run()).fold(
          (error) => reply(`Failed to delete task with ID: ${taskId} from list: ${listId}: ${graphErrorText(error)}`),
          () => reply(`Task with ID: ${taskId} was successfully deleted from list: ${listId}`),
        )
      } catch (error) {
        return reply(`Error deleting task: ${error}`)
      }
    },
  })

  server.addTool({
    name: "get-checklist-items",
    description:
      "Get checklist items (subtasks) for a specific task. Checklist items are smaller steps or components that belong to a parent task.",
    parameters: z.object({
      listId: z.string().describe("ID of the task list"),
      taskId: z.string().describe("ID of the task"),
    }),
    execute: async ({ listId, taskId }, context) => {
      try {
        const token = (await credentials.graphAuth(context)).orNull()
        if (token === null) return reply(NOT_AUTHENTICATED)

        // Fetch the task first to get its title
        const taskTitle = Option(
          await runOrNull(graphRequest<Task>(`${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`, token)),
        )
          .map((task) => task.title)
          .orElse("Unknown Task")

        // Fetch the checklist items
        const fetched = await graphRequestAll<ChecklistItem>(
          `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems`,
          token,
        ).run()
        if (fetched.isLeft()) return reply(`Failed to retrieve checklist items for task: ${taskId}`)

        const items = fetched.value
        if (items.length === 0) {
          return reply(`No checklist items found for task "${taskTitle}" (ID: ${taskId})`)
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

        return reply(`Checklist items for task "${taskTitle}" (ID: ${taskId}):\n\n${formattedItems.join("\n\n")}`)
      } catch (error) {
        return reply(`Error fetching checklist items: ${error}`)
      }
    },
  })

  server.addTool({
    name: "create-checklist-item",
    description:
      "Create a new checklist item (subtask) for a task. Checklist items help break down a task into smaller, manageable steps.",
    parameters: z.object({
      listId: z.string().describe("ID of the task list"),
      taskId: z.string().describe("ID of the task"),
      displayName: z.string().describe("Text content of the checklist item"),
      isChecked: z.boolean().optional().describe("Whether the item is checked off"),
    }),
    execute: async ({ listId, taskId, displayName, isChecked }, context) => {
      try {
        const token = (await credentials.graphAuth(context)).orNull()
        if (token === null) return reply(NOT_AUTHENTICATED)

        // Prepare the request body
        const requestBody: any = {
          displayName,
        }

        if (isChecked !== undefined) {
          requestBody.isChecked = isChecked
        }

        return (
          await graphRequest<ChecklistItem>(
            `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems`,
            token,
            "POST",
            requestBody,
          ).run()
        ).fold(
          () => reply(`Failed to create checklist item for task: ${taskId}`),
          (item) => reply(`Checklist item created successfully!\nContent: ${item.displayName}\nID: ${item.id}`),
        )
      } catch (error) {
        return reply(`Error creating checklist item: ${error}`)
      }
    },
  })

  server.addTool({
    name: "update-checklist-item",
    description:
      "Update an existing checklist item (subtask). Allows changing the text content or completion status of the subtask.",
    parameters: z.object({
      listId: z.string().describe("ID of the task list"),
      taskId: z.string().describe("ID of the task"),
      checklistItemId: z.string().describe("ID of the checklist item to update"),
      displayName: z.string().optional().describe("New text content of the checklist item"),
      isChecked: z.boolean().optional().describe("Whether the item is checked off"),
    }),
    execute: async ({ listId, taskId, checklistItemId, displayName, isChecked }, context) => {
      try {
        const token = (await credentials.graphAuth(context)).orNull()
        if (token === null) return reply(NOT_AUTHENTICATED)

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
          return reply("No properties provided for update. Please specify either displayName or isChecked.")
        }

        return (
          await graphRequest<ChecklistItem>(
            `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems/${checklistItemId}`,
            token,
            "PATCH",
            requestBody,
          ).run()
        ).fold(
          () => reply(`Failed to update checklist item with ID: ${checklistItemId}`),
          (item) =>
            reply(
              `Checklist item updated successfully!\nContent: ${item.displayName}\nStatus: ${item.isChecked ? "Checked" : "Not checked"}`,
            ),
        )
      } catch (error) {
        return reply(`Error updating checklist item: ${error}`)
      }
    },
  })

  server.addTool({
    name: "delete-checklist-item",
    description:
      "Delete a checklist item (subtask) from a task. This removes just the specific subtask, not the parent task.",
    parameters: z.object({
      listId: z.string().describe("ID of the task list"),
      taskId: z.string().describe("ID of the task"),
      checklistItemId: z.string().describe("ID of the checklist item to delete"),
    }),
    execute: async ({ listId, taskId, checklistItemId }, context) => {
      try {
        const token = (await credentials.graphAuth(context)).orNull()
        if (token === null) return reply(NOT_AUTHENTICATED)

        // Make a DELETE request to the Microsoft Graph API
        const url = `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems/${checklistItemId}`
        console.error(`Deleting checklist item: ${url}`)

        return (await graphRequest<void>(url, token, "DELETE").run()).fold(
          (error) =>
            reply(
              `Failed to delete checklist item with ID: ${checklistItemId} from task: ${taskId}: ${graphErrorText(error)}`,
            ),
          () => reply(`Checklist item with ID: ${checklistItemId} was successfully deleted from task: ${taskId}`),
        )
      } catch (error) {
        return reply(`Error deleting checklist item: ${error}`)
      }
    },
  })

  // Bulk archive completed tasks
  server.addTool({
    name: "archive-completed-tasks",
    description:
      "Move completed tasks older than a specified number of days from one list to another (archive) list. Useful for cleaning up active lists while preserving historical tasks.",
    parameters: z.object({
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
    }),
    execute: async ({ sourceListId, targetListId, olderThanDays, dryRun }, context) => {
      try {
        const token = (await credentials.graphAuth(context)).orNull()
        if (token === null) return reply(NOT_AUTHENTICATED)

        // Calculate cutoff date
        const cutoffDate = new Date()
        cutoffDate.setDate(cutoffDate.getDate() - olderThanDays)

        // Get all completed tasks from source list
        const completed = await graphRequestAll<Task>(
          `${MS_GRAPH_BASE}/me/todo/lists/${sourceListId}/tasks?$filter=status eq 'completed'`,
          token,
        ).run()
        if (completed.isLeft()) return reply("Failed to retrieve tasks from source list")

        // Filter tasks older than cutoff
        const tasksToArchive = completed.value.filter((task) => {
          if (!task.completedDateTime?.dateTime) return false
          const completedDate = new Date(task.completedDateTime.dateTime)
          return completedDate < cutoffDate
        })

        if (tasksToArchive.length === 0) {
          return reply(`No completed tasks found older than ${olderThanDays} days.`)
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

          return reply(preview)
        }

        // One task at a time: copy it to the target list, then delete it from the source.
        type Outcome = "archived" | "notCopied" | "notRemoved"
        const archiveOne = (task: Task): IO<never, never, Outcome> =>
          graphRequest(`${MS_GRAPH_BASE}/me/todo/lists/${targetListId}/tasks`, token, "POST", {
            title: task.title,
            status: "completed",
            body: task.body,
            importance: task.importance,
            completedDateTime: task.completedDateTime,
            dueDateTime: task.dueDateTime,
            reminderDateTime: task.reminderDateTime,
            categories: task.categories,
          })
            .map((): Outcome => "notRemoved")
            .flatMap(() =>
              graphRequest<void>(`${MS_GRAPH_BASE}/me/todo/lists/${sourceListId}/tasks/${task.id}`, token, "DELETE")
                .map((): Outcome => "archived")
                .recover<Outcome>("notRemoved"),
            )
            .recover<Outcome>("notCopied")

        const outcomes = (await IO.forEach(tasksToArchive, archiveOne).run()).orElse(
          tasksToArchive.map((): Outcome => "notCopied"),
        )
        const titlesWith = (outcome: Outcome) =>
          tasksToArchive.filter((_, i) => outcomes[i] === outcome).map((task) => task.title)
        const successCount = titlesWith("archived").length
        const failedTasks = titlesWith("notCopied")
        const notRemoved = titlesWith("notRemoved")

        let result = `📦 Archive Complete\n`
        result += `Successfully archived ${successCount} of ${tasksToArchive.length} tasks\n`
        result += `Tasks completed before ${cutoffDate.toLocaleDateString()} were moved.\n`

        if (failedTasks.length > 0) {
          result += `\n⚠️ Failed to archive ${failedTasks.length} tasks:\n`
          result += failedTasks.map((title) => `- ${title}\n`).join("")
        }

        if (notRemoved.length > 0) {
          result += `\n⚠️ Copied ${notRemoved.length} tasks to the archive list but could not delete them from the source list, so they are now in both:\n`
          result += notRemoved.map((title) => `- ${title}\n`).join("")
        }

        return reply(result)
      } catch (error) {
        return reply(`Error archiving tasks: ${error}`)
      }
    },
  })

  // Test tool to explore Graph API for hidden properties
  server.addTool({
    name: "test-graph-api-exploration",
    description:
      "Test various Graph API queries to discover hidden properties or endpoints for folder/group organization in Microsoft To Do.",
    parameters: z.object({
      testType: z
        .enum(["odata-select", "odata-expand", "headers", "extensions", "all"])
        .describe("Type of test to run"),
    }),
    execute: async ({ testType }, context) => {
      try {
        const token = (await credentials.graphAuth(context)).orNull()
        if (token === null) return reply(NOT_AUTHENTICATED)

        let results = "🔍 Graph API Exploration Results\n" + "=".repeat(50) + "\n\n"

        type Listing = { value?: Array<Record<string, unknown>> }
        // Each probe reports Graph's reason when the request fails, instead of treating the
        // failure as an empty success.
        const probe = <T = Listing>(path: string) => graphRequest<T>(`${MS_GRAPH_BASE}${path}`, token).run()

        // Test 1: Try with $select=* to get all properties
        if (testType === "odata-select" || testType === "all") {
          results += "📊 Test 1: Using $select=* to retrieve all properties\n"
          results += (await probe("/me/todo/lists?$select=*")).fold(
            (error) => `Error: ${graphErrorText(error)}\n`,
            (response) =>
              Option(response.value?.[0]).fold(
                () => "",
                (firstList) => {
                  const properties = Object.keys(firstList)
                  return (
                    `Found ${properties.length} properties: ${properties.join(", ")}\n` +
                    "\nExample list object:\n" +
                    JSON.stringify(firstList, null, 2).substring(0, 1000) +
                    "...\n"
                  )
                },
              ),
          )
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
            results += (await probe(`/me/todo/lists?$expand=${expand}&$top=1`)).fold(
              (error) => `✗ $expand=${expand}: ${graphErrorText(error)}\n`,
              (response) => {
                const data = response.value?.[0]?.[expand]
                return data
                  ? `✓ $expand=${expand}: Success - Found data!\n${JSON.stringify(data, null, 2).substring(0, 500)}...\n`
                  : `✓ $expand=${expand}: Success - No additional data returned\n`
              },
            )
          }
          results += "\n"
        }

        // Test 3: Check response headers for additional info
        if (testType === "headers" || testType === "all") {
          results += "📊 Test 3: Checking response headers\n"
          try {
            // A raw fetch on purpose: this test reports the response headers themselves.
            const response = await fetch(`${MS_GRAPH_BASE}/me/todo/lists`, {
              headers: {
                Authorization: `Bearer ${token.accessToken}`,
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
          const lists = await probe<{ value?: TaskList[] }>("/me/todo/lists?$top=1")
          if (lists.isLeft()) {
            results += `Error: ${graphErrorText(lists.value)}\n`
          } else {
            const listId = lists.value.value?.[0]?.id
            if (listId) {
              results += (await probe<unknown>(`/me/todo/lists/${listId}/extensions`)).fold(
                (error) => `No extensions endpoint: ${graphErrorText(error)}\n`,
                (extensions) => `Extensions found: ${JSON.stringify(extensions, null, 2)}\n`,
              )
            }
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
            results += (await probe<unknown>(endpoint)).fold(
              (error) => `✗ ${endpoint}: Not found (${graphErrorText(error)})\n`,
              (response) => `✓ ${endpoint}: Found! Response: ${JSON.stringify(response).substring(0, 200)}...\n`,
            )
          }
        }

        results += "\n" + "=".repeat(50) + "\n"
        results += "Analysis complete. Check results above for any discovered properties or endpoints."

        return reply(results)
      } catch (error) {
        return reply(`Error during Graph API exploration: ${error}`)
      }
    },
  })

  return server
}
