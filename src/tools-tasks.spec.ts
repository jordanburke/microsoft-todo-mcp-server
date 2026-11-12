// Task tools: get-tasks, create-task, update-task, delete-task. Recurrence has its own spec
// (recurrence.spec.ts); this one covers everything else those tools send and print.
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it } from "vitest"

import {
  GRAPH,
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
afterEach(() => restoreNetwork())

const TASKS = "/me/todo/lists/L1/tasks"

const task = (id: string, extra: Record<string, unknown> = {}) => ({
  id,
  title: `Task ${id}`,
  status: "notStarted",
  importance: "normal",
  ...extra,
})

describe("get-tasks", () => {
  it("prints each task's fields", async () => {
    h.net.on(
      "GET",
      TASKS,
      json({
        value: [
          task("T1", {
            status: "completed",
            importance: "high",
            dueDateTime: { dateTime: "2026-12-15T00:00:00.0000000", timeZone: "UTC" },
            categories: ["Home", "Bills"],
            body: { content: "Short note", contentType: "text" },
          }),
          task("T2", { body: { content: "x".repeat(60), contentType: "text" } }),
          task("T3", { body: { content: "   ", contentType: "text" }, categories: [] }),
        ],
      }),
    )

    const { text } = await h.callTool("get-tasks", { listId: "L1" })

    expect(text).toBe(
      [
        "Tasks in list L1:",
        "",
        "✓ ID: T1\nTitle: Task T1\nDue: 12/15/2026\nImportance: high\nCategories: Home, Bills\nDescription: Short note\n---",
        `○ ID: T2\nTitle: Task T2\nImportance: normal\nDescription: ${"x".repeat(50)}...\n---`,
        "○ ID: T3\nTitle: Task T3\nImportance: normal\n---",
      ].join("\n"),
    )
  })

  it("forwards filter, orderby, top, skip and count as OData parameters, but not select", async () => {
    h.net.on("GET", /^\/me\/todo\/lists\/L1\/tasks\?/, json({ value: [task("T1")] }))

    await h.callTool("get-tasks", {
      listId: "L1",
      filter: "status eq 'completed'",
      select: "status",
      orderby: "createdDateTime desc",
      top: 5,
      skip: 2,
      count: true,
    })

    expect(h.net.graphRequests()).toEqual([
      `GET ${TASKS}?%24filter=status+eq+%27completed%27&%24orderby=createdDateTime+desc&%24top=5&%24skip=2&%24count=true`,
    ])
  })

  it("sends no query string when no options are given", async () => {
    h.net.on("GET", TASKS, json({ value: [task("T1")] }))

    await h.callTool("get-tasks", { listId: "L1" })

    expect(h.net.graphRequests()).toEqual([`GET ${TASKS}`])
  })

  it("applies select on the client: id and title always, the named fields only", async () => {
    h.net.on("GET", TASKS, json({ value: [task("T1", { status: "completed", categories: ["A"] })] }))

    const { text } = await h.callTool("get-tasks", { listId: "L1", select: " categories , bogus" })

    expect(text).toBe("Tasks in list L1:\n\nID: T1\nTitle: Task T1\nCategories: A\n---")
  })

  it("follows nextLink across pages and stops once top tasks are collected", async () => {
    h.net
      .on(
        "GET",
        `${TASKS}?%24top=3`,
        json({ value: [task("T1"), task("T2")], "@odata.nextLink": `${GRAPH}${TASKS}?$skip=2` }),
      )
      .on(
        "GET",
        `${TASKS}?$skip=2`,
        json({ value: [task("T3"), task("T4")], "@odata.nextLink": `${GRAPH}${TASKS}?$skip=4` }),
      )

    const { text } = await h.callTool("get-tasks", { listId: "L1", top: 3 })

    expect(h.net.graphRequests()).toEqual([`GET ${TASKS}?%24top=3`, `GET ${TASKS}?$skip=2`])
    expect(text).toContain("ID: T3")
    expect(text).not.toContain("ID: T4")
  })

  it("gives up after 100 pages and says why", async () => {
    let page = 0
    h.net.on("GET", /^\/me\/todo\/lists\/L1\/tasks/, () =>
      json({ value: [task(`T${++page}`)], "@odata.nextLink": `${GRAPH}${TASKS}?$skip=${page}` }),
    )

    const { text } = await h.callTool("get-tasks", { listId: "L1" })

    expect(h.net.calls).toHaveLength(100)
    expect(text).toBe(
      "Failed to retrieve tasks for list: L1 | Graph said: stopped after 100 pages without reaching the end of the collection",
    )
  })

  it("KNOWN ISSUE: count prints the number returned, not Graph's total", async () => {
    h.net.on("GET", /^\/me\/todo\/lists\/L1\/tasks/, json({ value: [task("T1"), task("T2")], "@odata.count": 40 }))

    const { text } = await h.callTool("get-tasks", { listId: "L1", count: true })

    expect(text).toContain("Total count: 2\n\n")
  })

  it("KNOWN ISSUE: offers orderby=title, which Graph rejects", async () => {
    h.net.on(
      "GET",
      `${TASKS}?%24orderby=title`,
      graphError(400, "BadRequest", "The query specified in the URI is not valid."),
    )

    const { text } = await h.callTool("get-tasks", { listId: "L1", orderby: "title" })

    expect(text).toBe(
      'Failed to retrieve tasks for list: L1 | Graph said: HTTP error! status: 400, body: {"error":{"code":"BadRequest","message":"The query specified in the URI is not valid."}}',
    )
  })

  it("says so when the list is empty", async () => {
    h.net.on("GET", TASKS, json({ value: [] }))

    expect((await h.callTool("get-tasks", { listId: "L1" })).text).toBe("No tasks found in list with ID: L1")
  })
})

describe("create-task", () => {
  it("sends every field in Graph's shape, with UTC date-times", async () => {
    h.net.on("POST", TASKS, json(task("T9", { title: "Pay rent" }), 201))

    const { text } = await h.callTool("create-task", {
      listId: "L1",
      title: "Pay rent",
      body: "Transfer from checking",
      dueDateTime: "2031-01-01T09:00:00Z",
      startDateTime: "2030-12-30T09:00:00Z",
      importance: "high",
      isReminderOn: false,
      reminderDateTime: "2030-12-31T09:00:00Z",
      status: "inProgress",
      categories: ["Bills"],
    })

    expect(h.net.calls[0].headers["content-type"]).toBe("application/json")
    expect(h.net.calls[0].body).toEqual({
      title: "Pay rent",
      body: { content: "Transfer from checking", contentType: "text" },
      dueDateTime: { dateTime: "2031-01-01T09:00:00Z", timeZone: "UTC" },
      startDateTime: { dateTime: "2030-12-30T09:00:00Z", timeZone: "UTC" },
      importance: "high",
      isReminderOn: false,
      reminderDateTime: { dateTime: "2030-12-31T09:00:00Z", timeZone: "UTC" },
      status: "inProgress",
      categories: ["Bills"],
    })
    expect(text).toBe("Task created successfully!\nID: T9\nTitle: Pay rent")
  })

  it("sends only the title when nothing else is given, and drops an empty category list", async () => {
    h.net.on("POST", TASKS, json(task("T9"), 201))

    await h.callTool("create-task", { listId: "L1", title: "Bare", categories: [] })

    expect(h.net.calls[0].body).toEqual({ title: "Bare" })
  })

  it("passes Graph's reason through on failure", async () => {
    h.net.on("POST", TASKS, graphError(400, "invalidRequest", "title too long"))

    expect((await h.callTool("create-task", { listId: "L1", title: "x" })).text).toBe(
      'Failed to create task in list: L1: HTTP error! status: 400, body: {"error":{"code":"invalidRequest","message":"title too long"}}',
    )
  })
})

describe("update-task", () => {
  it("sends only the fields given", async () => {
    h.net.on("PATCH", `${TASKS}/T1`, json(task("T1", { title: "New title" })))

    const { text } = await h.callTool("update-task", {
      listId: "L1",
      taskId: "T1",
      title: "New title",
      body: "Updated note",
      importance: "low",
      isReminderOn: true,
      status: "completed",
      categories: [],
      dueDateTime: "2031-02-01T00:00:00Z",
      startDateTime: "2031-01-15T00:00:00Z",
      reminderDateTime: "2031-01-31T08:00:00Z",
    })

    expect(h.net.graphRequests()).toEqual([`PATCH ${TASKS}/T1`])
    expect(h.net.calls[0].body).toEqual({
      title: "New title",
      body: { content: "Updated note", contentType: "text" },
      importance: "low",
      isReminderOn: true,
      status: "completed",
      categories: [],
      dueDateTime: { dateTime: "2031-02-01T00:00:00Z", timeZone: "UTC" },
      startDateTime: { dateTime: "2031-01-15T00:00:00Z", timeZone: "UTC" },
      reminderDateTime: { dateTime: "2031-01-31T08:00:00Z", timeZone: "UTC" },
    })
    expect(text).toBe("Task updated successfully!\nID: T1\nTitle: New title")
  })

  it("clears a date when it is passed as an empty string", async () => {
    h.net.on("PATCH", `${TASKS}/T1`, json(task("T1")))

    await h.callTool("update-task", {
      listId: "L1",
      taskId: "T1",
      dueDateTime: "",
      startDateTime: "",
      reminderDateTime: "",
    })

    expect(h.net.calls[0].body).toEqual({ dueDateTime: null, startDateTime: null, reminderDateTime: null })
  })

  it("refuses an update with nothing to change, without calling Graph", async () => {
    const { text } = await h.callTool("update-task", { listId: "L1", taskId: "T1" })

    expect(text).toBe("No properties provided for update. Please specify at least one property to change.")
    expect(h.net.calls).toHaveLength(0)
  })

  it("passes Graph's reason through on failure", async () => {
    h.net.on("PATCH", `${TASKS}/T1`, graphError(404, "itemNotFound", "gone"))

    expect((await h.callTool("update-task", { listId: "L1", taskId: "T1", title: "x" })).text).toBe(
      'Failed to update task with ID: T1 in list: L1: HTTP error! status: 404, body: {"error":{"code":"itemNotFound","message":"gone"}}',
    )
  })
})

describe("delete-task", () => {
  it("DELETEs the task and reports success on 204", async () => {
    h.net.on("DELETE", `${TASKS}/T1`, noContent())

    const { text } = await h.callTool("delete-task", { listId: "L1", taskId: "T1" })

    expect(h.net.graphRequests()).toEqual([`DELETE ${TASKS}/T1`])
    expect(text).toBe("Task with ID: T1 was successfully deleted from list: L1")
  })

  it("KNOWN ISSUE: reports success even when Graph refuses the delete", async () => {
    h.net.on("DELETE", `${TASKS}/missing`, graphError(404, "itemNotFound", "The task was not found"))

    expect((await h.callTool("delete-task", { listId: "L1", taskId: "missing" })).text).toBe(
      "Task with ID: missing was successfully deleted from list: L1",
    )
  })
})
