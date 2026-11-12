// auth-status, and what every tool says when there are no credentials at all.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import {
  type Harness,
  json,
  jwtExpiringIn,
  resetNetwork,
  restoreNetwork,
  startServer,
  type TokenSetup,
} from "./test-support/graph-harness.js"

let h: Harness

const start = async (tokens: TokenSetup): Promise<void> => {
  h = await startServer(tokens)
  resetNetwork(h.net)
}

beforeEach(() => {
  vi.useFakeTimers({ toFake: ["Date"] })
  vi.setSystemTime(new Date("2026-11-11T12:00:00Z"))
})
afterEach(() => {
  vi.useRealTimers()
  restoreNetwork(h.net)
  h.cleanup()
})

describe("auth-status", () => {
  it("reports a live token's expiry (the JWT's exp less five minutes), without a Graph request", async () => {
    await start({ kind: "env", accessToken: jwtExpiringIn(3600) })

    expect((await h.callTool("auth-status")).text).toBe("Authenticated. Token expires at 11/11/2026, 12:55:00 PM.")
    expect(h.net.calls).toHaveLength(0)
  })

  it("does not warn personal accounts, which have full To Do API access", async () => {
    await start({ kind: "env", accessToken: jwtExpiringIn(3600) })

    expect((await h.callTool("auth-status")).text).not.toContain("WARNING")
  })

  it("reports an expired token it could not renew", async () => {
    await start({ kind: "env", accessToken: jwtExpiringIn(-600) })
    h.net.onTokenRefresh(json({ error: "invalid_grant" }, 400))

    expect((await h.callTool("auth-status")).text).toBe(
      "Authentication expired at 11/11/2026, 11:45:00 AM. Will attempt to refresh when you call any API.",
    )
  })

  it("says how to sign in when there are no tokens", async () => {
    await start({ kind: "none" })

    expect((await h.callTool("auth-status")).text).toBe(
      "Not authenticated. Please run 'mstodo-setup' (or 'npx -p microsoft-todo-mcp-server mstodo-setup') to authenticate with Microsoft.",
    )
    expect(h.net.calls).toHaveLength(0)
  })
})

describe("every tool without credentials", () => {
  const CALLS: Array<[string, Record<string, unknown>]> = [
    ["get-task-lists", {}],
    ["get-task-lists-organized", {}],
    ["create-task-list", { displayName: "x" }],
    ["update-task-list", { listId: "L1", displayName: "x" }],
    ["delete-task-list", { listId: "L1" }],
    ["get-tasks", { listId: "L1" }],
    ["create-task", { listId: "L1", title: "x" }],
    ["update-task", { listId: "L1", taskId: "T1", title: "x" }],
    ["delete-task", { listId: "L1", taskId: "T1" }],
    ["get-checklist-items", { listId: "L1", taskId: "T1" }],
    ["create-checklist-item", { listId: "L1", taskId: "T1", displayName: "x" }],
    ["update-checklist-item", { listId: "L1", taskId: "T1", checklistItemId: "C1", isChecked: true }],
    ["delete-checklist-item", { listId: "L1", taskId: "T1", checklistItemId: "C1" }],
    ["archive-completed-tasks", { sourceListId: "A", targetListId: "B" }],
    ["test-graph-api-exploration", { testType: "headers" }],
  ]

  it.each(CALLS)("%s says it could not authenticate and sends nothing", async (name, args) => {
    await start({ kind: "none" })

    expect((await h.callTool(name, args)).text).toBe("Failed to authenticate with Microsoft API")
    expect(h.net.calls).toHaveLength(0)
  })
})
