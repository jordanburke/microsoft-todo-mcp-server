// archive-completed-tasks and test-graph-api-exploration.
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it, vi } from "vitest"

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
beforeEach(() => {
  resetNetwork(h.net)
  // Only Date is faked, so the server's own timers and promises run normally.
  vi.useFakeTimers({ toFake: ["Date"] })
  vi.setSystemTime(new Date("2026-11-11T12:00:00Z"))
})
afterEach(() => {
  vi.useRealTimers()
  restoreNetwork()
})

// Sent as a raw string, not URL-encoded: the archive tool builds this URL by hand.
const COMPLETED = "/me/todo/lists/SRC/tasks?$filter=status eq 'completed'"

const done = (id: string, completed: string | undefined, extra: Record<string, unknown> = {}) => ({
  id,
  title: `Done ${id}`,
  status: "completed",
  importance: "normal",
  ...(completed ? { completedDateTime: { dateTime: completed, timeZone: "UTC" } } : {}),
  ...extra,
})

// 2026-11-11 minus 90 days is 2026-08-13.
const OLD = done("old", "2026-08-01T10:00:00.0000000", {
  body: { content: "notes", contentType: "text" },
  categories: ["Bills"],
  dueDateTime: { dateTime: "2026-07-31T00:00:00.0000000", timeZone: "UTC" },
})
const RECENT = done("recent", "2026-11-01T10:00:00.0000000")
const UNDATED = done("undated", undefined)

describe("archive-completed-tasks", () => {
  it("previews with dryRun and changes nothing", async () => {
    h.net.on("GET", COMPLETED, json({ value: [OLD, RECENT, UNDATED] }))

    const { text } = await h.callTool("archive-completed-tasks", {
      sourceListId: "SRC",
      targetListId: "DST",
      dryRun: true,
    })

    expect(h.net.graphRequests()).toEqual([`GET ${COMPLETED}`])
    expect(text).toBe(
      "📋 Archive Preview\nWould archive 1 tasks completed before 8/13/2026\n\n- Done old (completed: 8/1/2026)\n",
    )
  })

  it("copies each old task to the target, then deletes it from the source", async () => {
    h.net
      .on("GET", COMPLETED, json({ value: [OLD, RECENT] }))
      .on("POST", "/me/todo/lists/DST/tasks", json({ id: "copy" }, 201))
      .on("DELETE", "/me/todo/lists/SRC/tasks/old", noContent())

    const { text } = await h.callTool("archive-completed-tasks", { sourceListId: "SRC", targetListId: "DST" })

    expect(h.net.graphRequests()).toEqual([
      `GET ${COMPLETED}`,
      "POST /me/todo/lists/DST/tasks",
      "DELETE /me/todo/lists/SRC/tasks/old",
    ])
    expect(h.net.calls[1].body).toEqual({
      title: "Done old",
      status: "completed",
      body: { content: "notes", contentType: "text" },
      importance: "normal",
      completedDateTime: { dateTime: "2026-08-01T10:00:00.0000000", timeZone: "UTC" },
      dueDateTime: { dateTime: "2026-07-31T00:00:00.0000000", timeZone: "UTC" },
      categories: ["Bills"],
    })
    expect(text).toBe(
      "📦 Archive Complete\nSuccessfully archived 1 of 1 tasks\nTasks completed before 8/13/2026 were moved.\n",
    )
  })

  it("honours olderThanDays", async () => {
    h.net.on("GET", COMPLETED, json({ value: [OLD, RECENT] }))

    const { text } = await h.callTool("archive-completed-tasks", {
      sourceListId: "SRC",
      targetListId: "DST",
      olderThanDays: 5,
      dryRun: true,
    })

    expect(text).toContain("Would archive 2 tasks completed before 11/6/2026")
  })

  it("lists the tasks it could not copy, and leaves them in the source", async () => {
    const OTHER = done("other", "2026-07-01T10:00:00.0000000")
    h.net
      .on("GET", COMPLETED, json({ value: [OLD, OTHER] }))
      .on("POST", "/me/todo/lists/DST/tasks", (call) =>
        (call.body as { title: string }).title === "Done old" ? json({ id: "copy" }, 201) : graphError(400, "x", "no"),
      )
      .on("DELETE", "/me/todo/lists/SRC/tasks/old", noContent())

    const { text } = await h.callTool("archive-completed-tasks", { sourceListId: "SRC", targetListId: "DST" })

    expect(h.net.graphRequests()).not.toContain("DELETE /me/todo/lists/SRC/tasks/other")
    expect(text).toContain("Successfully archived 1 of 2 tasks")
    expect(text).toContain("⚠️ Failed to archive 1 tasks:\n- Done other\n")
  })

  it("KNOWN ISSUE: counts a task as archived even when deleting it from the source fails", async () => {
    h.net
      .on("GET", COMPLETED, json({ value: [OLD] }))
      .on("POST", "/me/todo/lists/DST/tasks", json({ id: "copy" }, 201))
      .on("DELETE", "/me/todo/lists/SRC/tasks/old", graphError(403, "accessDenied", "no"))

    const { text } = await h.callTool("archive-completed-tasks", { sourceListId: "SRC", targetListId: "DST" })

    expect(text).toContain("Successfully archived 1 of 1 tasks")
  })

  it("says so when nothing is old enough", async () => {
    h.net.on("GET", COMPLETED, json({ value: [RECENT, UNDATED] }))

    expect((await h.callTool("archive-completed-tasks", { sourceListId: "SRC", targetListId: "DST" })).text).toBe(
      "No completed tasks found older than 90 days.",
    )
  })

  it("reports a failure to read the source list", async () => {
    h.net.on("GET", COMPLETED, graphError(404, "itemNotFound", "no list"))

    expect((await h.callTool("archive-completed-tasks", { sourceListId: "SRC", targetListId: "DST" })).text).toBe(
      "Failed to retrieve tasks from source list",
    )
  })
})

describe("test-graph-api-exploration", () => {
  it("odata-select lists the first list's properties", async () => {
    h.net.on("GET", "/me/todo/lists?$select=*", json({ value: [{ id: "L1", displayName: "Tasks" }] }))

    const { text } = await h.callTool("test-graph-api-exploration", { testType: "odata-select" })

    expect(text).toContain("Found 2 properties: id, displayName")
    expect(text.endsWith("Analysis complete. Check results above for any discovered properties or endpoints.")).toBe(
      true,
    )
  })

  it("odata-expand tries ten expansions and reports each", async () => {
    h.net.on("GET", /^\/me\/todo\/lists\?\$expand=/, json({ value: [{ id: "L1" }] }))

    const { text } = await h.callTool("test-graph-api-exploration", { testType: "odata-expand" })

    expect(h.net.calls).toHaveLength(10)
    expect(text).toContain("✓ $expand=extensions: Success - No additional data returned")
  })

  it("headers prints the response headers", async () => {
    h.net.on("GET", "/me/todo/lists", json({ value: [] }))

    const { text } = await h.callTool("test-graph-api-exploration", { testType: "headers" })

    expect(h.net.calls[0].headers.prefer).toBe("return=representation")
    expect(text).toContain("content-type: application/json")
  })

  it("extensions reads the first list's extensions", async () => {
    h.net
      .on("GET", "/me/todo/lists?$top=1", json({ value: [{ id: "L1" }] }))
      .on("GET", "/me/todo/lists/L1/extensions", json({ value: [] }))

    const { text } = await h.callTool("test-graph-api-exploration", { testType: "extensions" })

    expect(h.net.graphRequests()).toEqual(["GET /me/todo/lists?$top=1", "GET /me/todo/lists/L1/extensions"])
    expect(text).toContain('Extensions found: {\n  "value": []\n}')
  })

  it("KNOWN ISSUE: all reports every missing folder endpoint as Found, because failures return null", async () => {
    h.net.route(() => graphError(404, "itemNotFound", "nope"))

    const { text } = await h.callTool("test-graph-api-exploration", { testType: "all" })

    expect(text).toContain("✓ /me/todo/folders: Found! Response: null...")
  })
})
