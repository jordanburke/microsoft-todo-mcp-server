import { Client } from "@modelcontextprotocol/sdk/client/index.js"
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js"
import { mkdtempSync, rmSync } from "fs"
import { tmpdir } from "os"
import { join } from "path"
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it, vi } from "vitest"

const GRAPH = "https://graph.microsoft.com/v1.0/me/todo/lists/L1/tasks"

type GraphCall = { url: string; method: string; body?: Record<string, unknown> }

let calls: GraphCall[] = []
let respond: (call: GraphCall) => unknown = () => ({})
let status = 200
let client: Client
let dir: string
let originalCwd: string

const jwtExpiringIn = (seconds: number): string =>
  `header.${Buffer.from(JSON.stringify({ exp: Math.floor(Date.now() / 1000) + seconds })).toString("base64url")}.sig`

const callTool = async (name: string, args: Record<string, unknown>): Promise<string> => {
  const result = await client.callTool({ name, arguments: args })
  return (result.content as Array<{ text: string }>)[0].text
}

const monthly15 = {
  pattern: { type: "absoluteMonthly", interval: 1, dayOfMonth: 15 },
  range: { type: "noEnd" },
}

beforeAll(async () => {
  // The token manager is a module singleton, so its environment must be in place
  // before todo-index is imported. ./tokens.json in this repo is a real credential.
  dir = mkdtempSync(join(tmpdir(), "mstodo-recurrence-"))
  originalCwd = process.cwd()
  process.chdir(dir)
  process.env.MSTODO_TOKEN_FILE = join(dir, "tokens.json")
  process.env.MS_TODO_ACCESS_TOKEN = jwtExpiringIn(3600)
  process.env.MS_TODO_REFRESH_TOKEN = "refresh"

  const { server } = await import("./todo-index.js")
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair()
  await server.connect(serverTransport)
  client = new Client({ name: "spec", version: "0" })
  await client.connect(clientTransport)
})

afterAll(async () => {
  await client.close()
  process.chdir(originalCwd)
  rmSync(dir, { recursive: true, force: true })
  delete process.env.MSTODO_TOKEN_FILE
  delete process.env.MS_TODO_ACCESS_TOKEN
  delete process.env.MS_TODO_REFRESH_TOKEN
})

beforeEach(() => {
  calls = []
  status = 200
  vi.spyOn(console, "error").mockImplementation(() => {})
  vi.stubGlobal(
    "fetch",
    vi.fn(async (url: string, init?: RequestInit) => {
      const call: GraphCall = {
        url,
        method: init?.method ?? "GET",
        body: init?.body ? JSON.parse(String(init.body)) : undefined,
      }
      calls.push(call)
      return Response.json(respond(call), { status })
    }),
  )
})

afterEach(() => {
  vi.unstubAllGlobals()
  vi.restoreAllMocks()
})

describe("create-task recurrence", () => {
  it("sends the rule, with range.startDate defaulted from the due date", async () => {
    respond = ({ body }) => ({ id: "T1", title: "Rent", recurrence: body?.recurrence })

    const text = await callTool("create-task", {
      listId: "L1",
      title: "Rent",
      dueDateTime: "2026-11-15T09:00:00Z",
      recurrence: monthly15,
    })

    const post = calls.find((c) => c.method === "POST")
    expect(post?.body?.recurrence).toEqual({
      pattern: { type: "absoluteMonthly", interval: 1, dayOfMonth: 15 },
      range: { type: "noEnd", startDate: "2026-11-15" },
    })
    expect(text).toContain("Repeats: every absolute monthly on day 15 (range: noEnd)")
  })

  it("refuses a rule without a due date and sends nothing", async () => {
    const text = await callTool("create-task", { listId: "L1", title: "Rent", recurrence: monthly15 })

    expect(text).toContain("without a dueDateTime")
    expect(calls).toHaveLength(0)
  })
})

describe("update-task recurrence", () => {
  it("uses the task's existing due date when the call does not set one", async () => {
    respond = ({ method }) =>
      method === "GET"
        ? { id: "T1", title: "Rent", dueDateTime: { dateTime: "2026-12-15T00:00:00.0000000", timeZone: "UTC" } }
        : { id: "T1", title: "Rent" }

    await callTool("update-task", { listId: "L1", taskId: "T1", recurrence: monthly15 })

    expect(calls.map((c) => `${c.method} ${c.url}`)).toEqual([`GET ${GRAPH}/T1`, `PATCH ${GRAPH}/T1`])
    expect((calls[1].body?.recurrence as { range: { startDate: string } }).range.startDate).toBe("2026-12-15")
  })

  it("refuses a rule when the task has no due date", async () => {
    respond = () => ({ id: "T1", title: "Rent" })

    const text = await callTool("update-task", { listId: "L1", taskId: "T1", recurrence: monthly15 })

    expect(text).toContain("without a dueDateTime")
    expect(calls.map((c) => c.method)).toEqual(["GET"])
  })

  it("clears recurrence with null", async () => {
    respond = () => ({ id: "T1", title: "Rent" })

    await callTool("update-task", { listId: "L1", taskId: "T1", recurrence: null })

    expect(calls).toHaveLength(1)
    expect(calls[0].method).toBe("PATCH")
    expect(calls[0].body).toEqual({ recurrence: null })
  })
})

describe("get-tasks recurrence display", () => {
  beforeEach(() => {
    respond = () => ({ value: [{ id: "T1", title: "Rent", status: "notStarted", recurrence: monthly15 }] })
  })

  it("shows the rule by default", async () => {
    expect(await callTool("get-tasks", { listId: "L1" })).toContain("Repeats: every absolute monthly on day 15")
  })

  it("hides it when select leaves it out, and shows it when select names it", async () => {
    expect(await callTool("get-tasks", { listId: "L1", select: "status" })).not.toContain("Repeats:")
    expect(await callTool("get-tasks", { listId: "L1", select: "recurrence" })).toContain("Repeats:")
  })
})

describe("recurrence rules the server refuses", () => {
  const create = (recurrence: Record<string, unknown>): Promise<string> =>
    callTool("create-task", { listId: "L1", title: "Rent", dueDateTime: "2026-11-15T09:00:00Z", recurrence })

  it.each([
    [
      "absoluteMonthly without dayOfMonth",
      { type: "absoluteMonthly", interval: 1 },
      { type: "noEnd" },
      "needs dayOfMonth",
    ],
    [
      "absoluteYearly without month",
      { type: "absoluteYearly", interval: 1, dayOfMonth: 4 },
      { type: "noEnd" },
      "needs month",
    ],
    [
      "weekly without daysOfWeek",
      { type: "weekly", interval: 1, daysOfWeek: [] },
      { type: "noEnd" },
      "needs daysOfWeek",
    ],
    [
      "weekly with dayOfMonth",
      { type: "weekly", interval: 1, daysOfWeek: ["monday"], dayOfMonth: 3 },
      { type: "noEnd" },
      "does not use dayOfMonth",
    ],
    ["endDate range without endDate", { type: "daily", interval: 1 }, { type: "endDate" }, "needs endDate"],
    [
      "numbered range without a count",
      { type: "daily", interval: 1 },
      { type: "numbered" },
      "needs numberOfOccurrences",
    ],
    [
      "noEnd range with an endDate",
      { type: "daily", interval: 1 },
      { type: "noEnd", endDate: "2027-01-01" },
      "does not use endDate",
    ],
  ])("refuses %s before sending anything", async (_name, pattern, range, message) => {
    const text = await create({ pattern, range })

    expect(text).toContain("Recurrence not set")
    expect(text).toContain(message)
    expect(calls).toHaveLength(0)
  })

  it("refuses on update-task before reading the task", async () => {
    const text = await callTool("update-task", {
      listId: "L1",
      taskId: "T1",
      recurrence: { pattern: { type: "absoluteMonthly", interval: 1 }, range: { type: "noEnd" } },
    })

    expect(text).toContain("needs dayOfMonth")
    expect(calls).toHaveLength(0)
  })

  it("gives a weekly rule firstDayOfWeek, which the docs list as required", async () => {
    respond = ({ body }) => ({ id: "T1", title: "Standup", recurrence: body?.recurrence })

    await create({ pattern: { type: "weekly", interval: 1, daysOfWeek: ["monday"] }, range: { type: "noEnd" } })

    expect((calls[0].body?.recurrence as { pattern: unknown }).pattern).toEqual({
      type: "weekly",
      interval: 1,
      daysOfWeek: ["monday"],
      firstDayOfWeek: "sunday",
    })
  })

  it("passes Graph's reason through when Graph still refuses", async () => {
    status = 400
    respond = () => ({ error: { code: "invalidRequest", message: "Invalid recurrence pattern" } })

    const text = await create(monthly15)

    expect(text).toContain("Failed to create task in list: L1")
    expect(text).toContain("Invalid recurrence pattern")
  })
})

describe("Repeats line on what Graph returns", () => {
  // Graph fills every field on read, whatever the type; these shapes are from a live account.
  const stored = (pattern: Record<string, unknown>, range: Record<string, unknown>) => ({
    pattern: { month: 0, dayOfMonth: 0, daysOfWeek: [], firstDayOfWeek: "sunday", index: "first", ...pattern },
    range: {
      startDate: "2026-11-20",
      endDate: "0001-01-01",
      recurrenceTimeZone: "UTC",
      numberOfOccurrences: 0,
      ...range,
    },
  })

  it.each([
    [
      { type: "weekly", interval: 2, daysOfWeek: ["friday"] },
      { type: "numbered", endDate: "2027-01-15", numberOfOccurrences: 5 },
      "every 2 weekly on friday (range: numbered, 5x, ends 2027-01-15)",
    ],
    [
      { type: "relativeMonthly", interval: 1, daysOfWeek: ["tuesday"], index: "second" },
      { type: "noEnd" },
      "every relative monthly on the second tuesday (range: noEnd)",
    ],
    [
      { type: "absoluteYearly", interval: 1, dayOfMonth: 4, month: 9 },
      { type: "endDate", endDate: "2030-09-04" },
      "every absolute yearly on day 4 of month 9 (range: endDate until 2030-09-04)",
    ],
  ])("renders %j", async (pattern, range, expected) => {
    respond = () => ({ value: [{ id: "T1", title: "x", recurrence: stored(pattern, range) }] })

    expect(await callTool("get-tasks", { listId: "L1", select: "recurrence" })).toContain(`Repeats: ${expected}`)
  })
})
