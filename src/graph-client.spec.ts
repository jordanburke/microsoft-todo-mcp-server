// What every tool shares through the Graph client: request headers, the 401 refresh and
// retry, and how failures reach the caller. Each test starts a fresh server, because a
// refresh changes the token manager's state.
import { readFileSync } from "fs"
import { join } from "path"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import {
  graphError,
  type Harness,
  json,
  LOGIN,
  resetNetwork,
  restoreNetwork,
  startServer,
} from "./test-support/graph-harness.js"

let h: Harness

beforeEach(async () => {
  h = await startServer({ kind: "env", accessToken: "opaque-seed", refreshToken: "seed-refresh" })
  resetNetwork(h.net)
})
afterEach(() => {
  restoreNetwork(h.net)
  h.cleanup()
})

const LISTS = "/me/todo/lists"
const ok = json({ value: [{ id: "L1", displayName: "Tasks" }] })
const renewed = (accessToken: string) =>
  json({ access_token: accessToken, refresh_token: "rotated-refresh", expires_in: 3600 })

describe("request headers", () => {
  it("sends the bearer token, a user agent and JSON headers", async () => {
    h.net.on("GET", LISTS, ok)

    await h.callTool("get-task-lists")

    expect(h.net.calls[0].headers).toMatchObject({
      "user-agent": "microsoft-todo-mcp-server/1.0",
      accept: "application/json",
      authorization: "Bearer opaque-seed",
    })
  })

  it("sends a JSON body on POST", async () => {
    h.net.on("POST", LISTS, json({ id: "L9", displayName: "New" }, 201))

    await h.callTool("create-task-list", { displayName: "New" })

    expect(h.net.calls[0].headers["content-type"]).toBe("application/json")
    expect(h.net.calls[0].rawBody).toBe('{"displayName":"New"}')
  })

  it("sends no body on GET or DELETE", async () => {
    h.net.on("GET", LISTS, ok).on("DELETE", `${LISTS}/L1`, () => new Response(null, { status: 204 }))

    await h.callTool("get-task-lists")
    await h.callTool("delete-task-list", { listId: "L1" })

    expect(h.net.calls.map((c) => c.rawBody)).toEqual([undefined, undefined])
  })
})

describe("401 from Graph", () => {
  it("refreshes the token once and retries the request with the new one", async () => {
    h.net
      .on("GET", LISTS, (call) =>
        call.headers.authorization === "Bearer fresh-token"
          ? ok
          : graphError(401, "InvalidAuthenticationToken", "expired"),
      )
      .onTokenRefresh(renewed("fresh-token"))

    const { text } = await h.callTool("get-task-lists")

    expect(h.net.calls.map((c) => `${c.method} ${c.url}`)).toEqual([
      "GET https://graph.microsoft.com/v1.0/me/todo/lists",
      `POST ${LOGIN}/organizations/oauth2/v2.0/token`,
      "GET https://graph.microsoft.com/v1.0/me/todo/lists",
    ])
    expect(Object.fromEntries(new URLSearchParams(h.net.calls[1].rawBody))).toEqual({
      client_id: "spec-client-id",
      client_secret: "spec-client-secret",
      refresh_token: "seed-refresh",
      grant_type: "refresh_token",
      scope: "offline_access Tasks.Read Tasks.ReadWrite Tasks.Read.Shared Tasks.ReadWrite.Shared User.Read",
    })
    const form = h.net.calls[1]
    expect(form.formBody || form.headers["content-type"]?.startsWith("application/x-www-form-urlencoded")).toBe(true)
    expect(text).toContain("ID: L1")
  })

  it("keeps using the renewed token, and saves it with the client credentials", async () => {
    h.net
      .on("GET", LISTS, (call) =>
        call.headers.authorization === "Bearer fresh-token" ? ok : graphError(401, "InvalidAuthenticationToken", "x"),
      )
      .onTokenRefresh(renewed("fresh-token"))

    await h.callTool("get-task-lists")
    const second = await h.callTool("get-task-lists")

    expect(second.text).toContain("ID: L1")
    expect(h.net.calls.at(-1)?.headers.authorization).toBe("Bearer fresh-token")
    expect(h.net.calls).toHaveLength(4)
    expect(JSON.parse(readFileSync(join(h.dir, "tokens.json"), "utf8"))).toMatchObject({
      accessToken: "fresh-token",
      refreshToken: "rotated-refresh",
      clientId: "spec-client-id",
      clientSecret: "spec-client-secret",
      tenantId: "organizations",
    })
  })

  it("does not retry when the refresh fails", async () => {
    h.net
      .on("GET", LISTS, graphError(401, "InvalidAuthenticationToken", "x"))
      .onTokenRefresh(json({ error: "invalid_grant" }, 400))

    const { text } = await h.callTool("get-task-lists")

    expect(h.net.calls).toHaveLength(2)
    expect(text).toBe("Failed to retrieve task lists")
  })

  it("resends the request body on the retry", async () => {
    h.net
      .on("POST", `${LISTS}/L1/tasks`, (call) =>
        call.headers.authorization === "Bearer fresh-token"
          ? json({ id: "T1", title: "Pay rent" }, 201)
          : graphError(401, "InvalidAuthenticationToken", "expired"),
      )
      .onTokenRefresh(renewed("fresh-token"))

    const { text } = await h.callTool("create-task", { listId: "L1", title: "Pay rent", importance: "high" })

    expect(h.net.graphRequests()).toEqual(["POST /me/todo/lists/L1/tasks", "POST /me/todo/lists/L1/tasks"])
    expect(h.net.calls[2].body).toEqual(h.net.calls[0].body)
    expect(h.net.calls[2].body).toEqual({ title: "Pay rent", importance: "high" })
    expect(text).toBe("Task created successfully!\nID: T1\nTitle: Pay rent")
  })

  it("does not loop when the retried request is rejected too", async () => {
    h.net.on("GET", LISTS, graphError(401, "InvalidAuthenticationToken", "x")).onTokenRefresh(renewed("also-bad"))

    const { text } = await h.callTool("get-task-lists")

    expect(h.net.calls).toHaveLength(3)
    expect(text).toBe("Failed to retrieve task lists")
  })
})

describe("failures", () => {
  it("explains MailboxNotEnabledForRESTAPI instead of passing Graph's body through", async () => {
    h.net.on(
      "POST",
      `${LISTS}/L1/tasks`,
      graphError(401, "MailboxNotEnabledForRESTAPI", "REST API is not yet supported for this mailbox."),
    )
    h.net.onTokenRefresh(json({ error: "invalid_grant" }, 400))

    expect((await h.callTool("create-task", { listId: "L1", title: "x" })).text).toBe(
      "Failed to create task in list: L1: Microsoft has not enabled the To Do API for this account's mailbox (MailboxNotEnabledForRESTAPI).",
    )
  })

  it("reports a network error as the failure reason", async () => {
    h.net.route(() => () => {
      throw new TypeError("fetch failed")
    })

    expect((await h.callTool("create-task", { listId: "L1", title: "x" })).text).toBe(
      "Failed to create task in list: L1: fetch failed",
    )
  })

  it("reports a 2xx reply that is not JSON as a failure", async () => {
    h.net.on("POST", `${LISTS}/L1/tasks`, new Response("<html>gateway</html>", { status: 200 }))

    const { text } = await h.callTool("create-task", { listId: "L1", title: "x" })

    // The reason's wording comes from the JSON parser, so only its presence is pinned.
    expect(text).toMatch(/^Failed to create task in list: L1: \S/)
  })

  it("logs the failure reason to stderr even when the reply does not show it", async () => {
    h.net.on("GET", LISTS, graphError(500, "generalException", "boom"))

    expect((await h.callTool("get-task-lists")).text).toBe("Failed to retrieve task lists")

    const logged = vi.mocked(console.error).mock.calls.map((args) => args.join(" "))
    expect(logged.some((line) => line.includes("status: 500") && line.includes("generalException"))).toBe(true)
  })

  it("leaves a request's query, which can hold a caller's $filter, out of the log", async () => {
    h.net.on("GET", /^\/me\/todo\/lists\/L1\/tasks\?/, graphError(400, "badRequest", "nope"))

    await h.callTool("get-tasks", { listId: "L1", filter: "title eq 'secret plans'" })

    const logged = vi
      .mocked(console.error)
      .mock.calls.map((args) => args.join(" "))
      .join("\n")
    expect(logged).toContain("Graph request failed: GET https://graph.microsoft.com/v1.0/me/todo/lists/L1/tasks:")
    expect(logged).not.toContain("secret")
  })

  it("reports every Graph failure as an ordinary result, not an isError one", async () => {
    h.net
      .on("GET", LISTS, graphError(500, "generalException", "boom"))
      .on("POST", `${LISTS}/L1/tasks`, graphError(400, "invalidRequest", "bad"))
      .on("DELETE", `${LISTS}/L1`, graphError(404, "itemNotFound", "gone"))

    const results = [
      await h.callTool("get-task-lists"),
      await h.callTool("create-task", { listId: "L1", title: "x" }),
      await h.callTool("delete-task-list", { listId: "L1" }),
    ]

    expect(results.map((r) => r.isError)).toEqual([false, false, false])
  })
})

describe("IDs", () => {
  // Real Graph IDs are long base64 strings with "-", "_" and a trailing "=".
  const LIST = "AQMkADAwATM0MDAAMS1iNTcwLWI2NTEtMDACLTAwCgAuAAADiN8-_x=="
  const TASK = "AQMkADAwATM0MDAAMS1iNTcwLWI2NTEtMDACLTAwCgBGAAADiN8_-y=="

  it("puts them in the path as they are", async () => {
    h.net.on("PATCH", `${LISTS}/${LIST}/tasks/${TASK}`, json({ id: TASK, title: "x" }))

    await h.callTool("update-task", { listId: LIST, taskId: TASK, title: "x" })

    expect(h.net.calls[0].url).toBe(`https://graph.microsoft.com/v1.0/me/todo/lists/${LIST}/tasks/${TASK}`)
  })
})
