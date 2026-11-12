// Task-list tools: get-task-lists, get-task-lists-organized, create/update/delete-task-list.
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it } from "vitest"

import {
  graphError,
  type Harness,
  json,
  noContent,
  resetNetwork,
  restoreNetwork,
  startServer,
} from "./test-support/graph-harness.js"

let h: Harness

beforeAll(async () => {
  h = await startServer()
})
afterAll(() => h.cleanup())
beforeEach(() => resetNetwork(h.net))
afterEach(() => restoreNetwork(h.net))

const LISTS = [
  { id: "L-default", displayName: "Tasks", wellknownListName: "defaultList", isOwner: true, isShared: false },
  { id: "L-flagged", displayName: "Flagged email", wellknownListName: "flaggedEmails", isOwner: true, isShared: false },
  { id: "L-shared-mine", displayName: "Groceries", wellknownListName: "none", isOwner: true, isShared: true },
  { id: "L-shared-theirs", displayName: "Team board", wellknownListName: "none", isOwner: false, isShared: true },
  { id: "L-plain", displayName: "Ideas", wellknownListName: "none", isOwner: true, isShared: false },
]

describe("get-task-lists", () => {
  it("prints every list with its ID, well-known name and sharing", async () => {
    h.net.on("GET", "/me/todo/lists", json({ value: LISTS }))

    const { text, isError } = await h.callTool("get-task-lists")

    expect(isError).toBe(false)
    expect(text).toBe(
      [
        "Your task lists:",
        "",
        "ID: L-default\nName: Tasks (Default Tasks List)\n---",
        "ID: L-flagged\nName: Flagged email (Flagged Emails)\n---",
        "ID: L-shared-mine\nName: Groceries (Shared by you)\n---",
        "ID: L-shared-theirs\nName: Team board (Shared with you)\n---",
        "ID: L-plain\nName: Ideas\n---",
      ].join("\n"),
    )
  })

  it("follows @odata.nextLink to the last page", async () => {
    h.net
      .on(
        "GET",
        "/me/todo/lists",
        json({
          value: LISTS.slice(0, 2),
          "@odata.nextLink": `${"https://graph.microsoft.com/v1.0"}/me/todo/lists?$skiptoken=p2`,
        }),
      )
      .on("GET", "/me/todo/lists?$skiptoken=p2", json({ value: LISTS.slice(2) }))

    const { text } = await h.callTool("get-task-lists")

    expect(h.net.graphRequests()).toEqual(["GET /me/todo/lists", "GET /me/todo/lists?$skiptoken=p2"])
    for (const list of LISTS) expect(text).toContain(`ID: ${list.id}`)
  })

  it("fails the whole call when a later page fails, rather than returning part of the lists", async () => {
    h.net
      .on(
        "GET",
        "/me/todo/lists",
        json({
          value: LISTS.slice(0, 2),
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/todo/lists?$skiptoken=p2",
        }),
      )
      .on("GET", "/me/todo/lists?$skiptoken=p2", graphError(503, "serviceUnavailable", "try later"))

    expect((await h.callTool("get-task-lists")).text).toBe("Failed to retrieve task lists")
  })

  it("says so when there are no lists", async () => {
    h.net.on("GET", "/me/todo/lists", json({ value: [] }))

    expect((await h.callTool("get-task-lists")).text).toBe("No task lists found.")
  })

  it("reports a Graph failure without Graph's reason", async () => {
    h.net.on("GET", "/me/todo/lists", graphError(500, "generalException", "boom"))

    expect((await h.callTool("get-task-lists")).text).toBe("Failed to retrieve task lists")
  })
})

describe("get-task-lists-organized", () => {
  const CATEGORISED = [
    { id: "a1", displayName: "Tasks", wellknownListName: "defaultList" },
    { id: "a2", displayName: "🛒 Costco" },
    { id: "a3", displayName: "🛒 Hardware" },
    { id: "a4", displayName: "Work Projects" },
    { id: "a5", displayName: "SBIR grant" },
    { id: "a6", displayName: "🏡 Lake house" },
    { id: "a7", displayName: "👪 Kids" },
    { id: "a8", displayName: "🎄 Holidays" },
    { id: "a9", displayName: "🚗 Road trip" },
    { id: "a10", displayName: "📰 Articles" },
    { id: "a11", displayName: "📦 Archive 2024" },
    { id: "a12", displayName: "Old stuff (Work - Archived)" },
    { id: "a13", displayName: "Book club", isShared: true, isOwner: false },
    { id: "a14", displayName: "Misc" },
  ]

  // "Archived - Work  (1)" has two spaces: the archived pattern's capture keeps the space
  // before " - Archived". Pinned as it is today.
  it("groups lists into categories in a fixed order, archived groups last", async () => {
    h.net.on("GET", "/me/todo/lists", json({ value: CATEGORISED }))

    const { text } = await h.callTool("get-task-lists-organized")

    expect(text).toMatchInlineSnapshot(`
      "📂 Microsoft To Do Lists - Organized View
      ==================================================

      ⭐ Special Lists (1)
         └─ Tasks [Default]

      👥 Shared Lists (1)
         └─ Book club [Shared with you]

      💼 Work (2)
         ├─ Work Projects
         │
         └─ SBIR grant

      👪 Family (1)
         └─ 👪 Kids

      🏡 Properties (1)
         └─ 🏡 Lake house

      🛒 Shopping Lists (2)
         ├─ 🛒 Costco
         │
         └─ 🛒 Hardware

      🚗 Travel & Rangeley (1)
         └─ 🚗 Road trip

      🎉 Seasonal & Events (1)
         └─ 🎄 Holidays

      📚 Reading (1)
         └─ 📰 Articles

      📋 Other Lists (1)
         └─ Misc

      📦 Archives (1)
         └─ 📦 Archive 2024

      📦 Archived - Work  (1)
         └─ Old stuff (Work - Archived)

      --------------------------------------------------
      Summary: 14 lists in 12 categories
      "
    `)
  })

  it("appends an ID reference with includeIds", async () => {
    h.net.on("GET", "/me/todo/lists", json({ value: CATEGORISED.slice(0, 2) }))

    const { text } = await h.callTool("get-task-lists-organized", { includeIds: true })

    expect(text).toContain("📋 List IDs Reference:\n" + "-".repeat(50) + "\nTasks: a1\n🛒 Costco: a2\n")
  })

  it("groups by sharing status with groupBy=shared", async () => {
    h.net.on("GET", "/me/todo/lists", json({ value: LISTS }))

    const { text } = await h.callTool("get-task-lists-organized", { groupBy: "shared" })

    expect(text).toBe(
      "📂 Microsoft To Do Lists - By Sharing Status\n" +
        "=".repeat(50) +
        "\n\n" +
        "👥 Shared Lists (2)\n" +
        "   ├─ Groceries [Shared by you]\n" +
        "   ├─ Team board [Shared with you]\n" +
        "\n🔒 Personal Lists (3)\n" +
        "   ├─ Tasks\n" +
        "   ├─ Flagged email\n" +
        "   ├─ Ideas\n",
    )
  })

  it("KNOWN ISSUE: groupBy=shared prints no list IDs, even with includeIds", async () => {
    h.net.on("GET", "/me/todo/lists", json({ value: LISTS }))

    const { text } = await h.callTool("get-task-lists-organized", { groupBy: "shared", includeIds: true })

    expect(text).not.toContain("L-default")
  })

  it("treats groupBy=type like the default category view", async () => {
    h.net.on("GET", "/me/todo/lists", json({ value: CATEGORISED }))
    const byCategory = (await h.callTool("get-task-lists-organized")).text

    expect((await h.callTool("get-task-lists-organized", { groupBy: "type" })).text).toBe(byCategory)
  })

  it("reports a Graph failure and an empty account", async () => {
    h.net.on("GET", "/me/todo/lists", graphError(500, "generalException", "boom"))
    expect((await h.callTool("get-task-lists-organized")).text).toBe("Failed to retrieve task lists")

    h.net.on("GET", "/me/todo/lists", json({ value: [] }))
    expect((await h.callTool("get-task-lists-organized")).text).toBe("No task lists found.")
  })
})

describe("create-task-list", () => {
  it("POSTs the name and prints the new list", async () => {
    h.net.on("POST", "/me/todo/lists", json({ id: "L-new", displayName: "Garden" }, 201))

    const { text } = await h.callTool("create-task-list", { displayName: "Garden" })

    expect(h.net.calls[0].body).toEqual({ displayName: "Garden" })
    expect(text).toBe("Task list created successfully!\nName: Garden\nID: L-new")
  })

  it("reports a failure", async () => {
    h.net.on("POST", "/me/todo/lists", graphError(400, "invalidRequest", "bad name"))

    expect((await h.callTool("create-task-list", { displayName: "Garden" })).text).toBe(
      "Failed to create task list: Garden",
    )
  })
})

describe("update-task-list", () => {
  it("PATCHes the list's name", async () => {
    h.net.on("PATCH", "/me/todo/lists/L1", json({ id: "L1", displayName: "Renamed" }))

    const { text } = await h.callTool("update-task-list", { listId: "L1", displayName: "Renamed" })

    expect(h.net.graphRequests()).toEqual(["PATCH /me/todo/lists/L1"])
    expect(h.net.calls[0].body).toEqual({ displayName: "Renamed" })
    expect(text).toBe("Task list updated successfully!\nNew name: Renamed")
  })

  it("reports a failure", async () => {
    h.net.on("PATCH", "/me/todo/lists/L1", graphError(404, "itemNotFound", "gone"))

    expect((await h.callTool("update-task-list", { listId: "L1", displayName: "x" })).text).toBe(
      "Failed to update task list with ID: L1",
    )
  })
})

describe("delete-task-list", () => {
  it("DELETEs the list and reports success on 204", async () => {
    h.net.on("DELETE", "/me/todo/lists/L1", noContent())

    const { text } = await h.callTool("delete-task-list", { listId: "L1" })

    expect(h.net.graphRequests()).toEqual(["DELETE /me/todo/lists/L1"])
    expect(text).toBe("Task list with ID: L1 was successfully deleted.")
  })

  it("reports Graph's refusal instead of success", async () => {
    h.net.on("DELETE", "/me/todo/lists/missing", graphError(404, "itemNotFound", "The list was not found"))

    expect((await h.callTool("delete-task-list", { listId: "missing" })).text).toBe(
      'Failed to delete task list with ID: missing: HTTP error! status: 404, body: {"error":{"code":"itemNotFound","message":"The list was not found"}}',
    )
  })
})
