// auth-status, and what every tool says when there are no credentials at all.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import {
  graphError,
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
  it("reports a live token's expiry (the JWT's exp less five minutes)", async () => {
    await start({ kind: "env", accessToken: jwtExpiringIn(3600) })
    h.net.on("GET", "/me", json({ userPrincipalName: "someone@contoso.com" }))

    expect((await h.callTool("auth-status")).text).toBe("Authenticated. Token expires at 11/11/2026, 12:55:00 PM.")
  })

  it("KNOWN ISSUE: warns every personal account that To Do may not work, though it does", async () => {
    await start({ kind: "env", accessToken: jwtExpiringIn(3600) })
    h.net.on("GET", "/me", json({ mail: "someone@outlook.com" }))

    const { text } = await h.callTool("auth-status")

    expect(text).toMatch(
      /^Authenticated\. Token expires at .*\n\n⚠️ WARNING: You are using a personal Microsoft account\./,
    )
  })

  it.each(["hotmail.com", "live.com", "msn.com", "passport.com"])("treats %s as personal", async (domain) => {
    await start({ kind: "env", accessToken: jwtExpiringIn(3600) })
    h.net.on("GET", "/me", json({ userPrincipalName: `someone@${domain}` }))

    expect((await h.callTool("auth-status")).text).toContain("⚠️ WARNING")
  })

  it("shows no warning when /me cannot be read", async () => {
    await start({ kind: "env", accessToken: jwtExpiringIn(3600) })
    h.net.on("GET", "/me", graphError(403, "Forbidden", "no"))

    expect((await h.callTool("auth-status")).text).not.toContain("WARNING")
  })

  it("reports an expired token it could not renew", async () => {
    await start({ kind: "env", accessToken: jwtExpiringIn(-600) })
    h.net
      .onTokenRefresh(json({ error: "invalid_grant" }, 400))
      .on("GET", "/me", graphError(401, "InvalidAuthenticationToken", "expired"))

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
