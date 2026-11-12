// The usual install: tokens in a token file, no environment pair. Driven through the tools,
// so it holds whatever the token manager is built from.
import { readFileSync } from "fs"
import { join } from "path"
import { afterEach, beforeEach, describe, expect, it } from "vitest"

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

const start = async (tokens: Record<string, unknown>): Promise<void> => {
  h = await startServer({ kind: "file", tokens })
  resetNetwork(h.net)
}

afterEach(() => {
  restoreNetwork(h.net)
  h.cleanup()
})

const LISTS = "/me/todo/lists"
const ok = json({ value: [{ id: "L1", displayName: "Tasks" }] })
const fileTokens = () => JSON.parse(readFileSync(join(h.dir, "tokens.json"), "utf8"))
const STORED_CREDENTIALS = { clientId: "file-client", clientSecret: "file-secret", tenantId: "contoso-tenant" }

describe("a live token in the file", () => {
  beforeEach(() => start({ accessToken: "file-token", refreshToken: "file-refresh", expiresAt: Date.now() + 3600_000 }))

  it("is sent as is", async () => {
    h.net.on("GET", LISTS, ok)

    await h.callTool("get-task-lists")

    expect(h.net.calls.map((c) => c.headers.authorization)).toEqual(["Bearer file-token"])
  })
})

describe("an expired token in the file", () => {
  beforeEach(() =>
    start({ accessToken: "stale", refreshToken: "file-refresh", expiresAt: Date.now() - 1000, ...STORED_CREDENTIALS }),
  )

  it("is renewed with the stored credentials and tenant before the request", async () => {
    h.net
      .onTokenRefresh(json({ access_token: "renewed", refresh_token: "rotated", expires_in: 3600 }))
      .on("GET", LISTS, ok)

    const { text } = await h.callTool("get-task-lists")

    expect(h.net.calls.map((c) => `${c.method} ${c.url}`)).toEqual([
      `POST ${LOGIN}/contoso-tenant/oauth2/v2.0/token`,
      "GET https://graph.microsoft.com/v1.0/me/todo/lists",
    ])
    const form = new URLSearchParams(h.net.calls[0].rawBody)
    expect([form.get("client_id"), form.get("client_secret"), form.get("refresh_token")]).toEqual([
      "file-client",
      "file-secret",
      "file-refresh",
    ])
    expect(h.net.calls[1].headers.authorization).toBe("Bearer renewed")
    expect(fileTokens()).toMatchObject({ accessToken: "renewed", refreshToken: "rotated", ...STORED_CREDENTIALS })
    expect(text).toContain("ID: L1")
  })

  it("is still sent when renewal fails, so Graph's 401 is what the caller sees", async () => {
    h.net
      .onTokenRefresh(json({ error: "invalid_grant" }, 400))
      .on("GET", LISTS, graphError(401, "InvalidAuthenticationToken", "expired"))

    const { text } = await h.callTool("get-task-lists")

    expect(h.net.calls.filter((c) => c.host === "graph.microsoft.com")[0].headers.authorization).toBe("Bearer stale")
    expect(text).toBe("Failed to retrieve task lists")
  })
})

describe("a 401 for a token the file says is live", () => {
  beforeEach(() =>
    start({
      accessToken: "revoked",
      refreshToken: "file-refresh",
      expiresAt: Date.now() + 3600_000,
      ...STORED_CREDENTIALS,
    }),
  )

  it("renews and retries once", async () => {
    h.net
      .on("GET", LISTS, (call) =>
        call.headers.authorization === "Bearer renewed" ? ok : graphError(401, "InvalidAuthenticationToken", "revoked"),
      )
      .onTokenRefresh(json({ access_token: "renewed", refresh_token: "rotated", expires_in: 3600 }))

    const { text } = await h.callTool("get-task-lists")

    expect(h.net.calls.map((c) => c.method)).toEqual(["GET", "POST", "GET"])
    expect(fileTokens().accessToken).toBe("renewed")
    expect(text).toContain("ID: L1")
  })
})
