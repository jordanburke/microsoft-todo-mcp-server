// The tool surface an MCP client sees: names, descriptions, input schemas and annotations.
// The snapshot makes any change to it, intended or not, show up in review.
import { afterAll, beforeAll, describe, expect, it } from "vitest"

import { type Harness, startServer } from "./test-support/graph-harness.js"

let h: Harness

beforeAll(async () => {
  h = await startServer()
})

afterAll(() => h.cleanup())

describe("tools/list", () => {
  it("lists the 16 To Do tools plus somamcp's info tool", async () => {
    const { result } = await h.rpc("tools/list")
    const names = (result as { tools: Array<{ name: string }> }).tools.map((t) => t.name)

    expect(names.sort()).toEqual(
      [
        "archive-completed-tasks",
        "auth-status",
        "create-checklist-item",
        "create-task",
        "create-task-list",
        "delete-checklist-item",
        "delete-task",
        "delete-task-list",
        "get-checklist-items",
        "get-task-lists",
        "get-task-lists-organized",
        "get-tasks",
        "info",
        "test-graph-api-exploration",
        "update-checklist-item",
        "update-task",
        "update-task-list",
      ].sort(),
    )
  })

  it("keeps every tool's description and input schema", async () => {
    const { result } = await h.rpc("tools/list")
    const tools = (result as { tools: Array<{ name: string }> }).tools
      .filter((t) => t.name !== "info")
      .sort((a, b) => a.name.localeCompare(b.name))

    expect(tools).toMatchSnapshot()
  })
})

describe("argument validation", () => {
  it("answers an unknown tool with JSON-RPC -32602", async () => {
    const { error } = await h.rpc("tools/call", { name: "no-such-tool", arguments: {} })

    expect(error?.code).toBe(-32602)
  })

  it("answers a missing required argument with an isError result naming the field", async () => {
    const result = await h.callTool("get-tasks", {})

    expect(result.isError).toBe(true)
    expect(result.text).toContain("listId")
  })

  it("rejects a value outside an enum", async () => {
    const result = await h.callTool("create-task", { listId: "L1", title: "x", importance: "urgent" })

    expect(result.isError).toBe(true)
    expect(result.text).toContain("importance")
  })
})
