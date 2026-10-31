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
      return Response.json(respond(call))
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
    expect(text).toContain("Repeats: absoluteMonthly")
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
