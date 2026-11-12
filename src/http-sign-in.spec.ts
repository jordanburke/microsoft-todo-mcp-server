// A whole sign-in through HTTP mode's real OAuth provider (somamcp's createAzureProvider), with
// only Microsoft faked: register, authorize, consent, Microsoft's callback, the code redemption,
// tool calls with the server's token, a refresh, and the refresh token's rotation. The live run
// on 2026-11-12 (see CLAUDE.md) proved the same against the real Microsoft; this keeps it proved
// across somamcp upgrades.
import { createHash, randomBytes } from "crypto"
import { mkdtempSync, rmSync } from "fs"
import { tmpdir } from "os"
import { join } from "path"
import { afterEach, beforeEach, describe, expect, it } from "vitest"

import { readHttpConfig } from "./http-config.js"
import { createHttpServer } from "./http-server.js"
import { type Call, FakeNetwork, json, resetNetwork, restoreNetwork } from "./test-support/graph-harness.js"

const BASE = "https://todo.example.com"
const REDIRECT = "http://localhost:5577/cb"
const env = {
  MSTODO_HTTP_BASE_URL: BASE,
  CLIENT_ID: "app-client",
  CLIENT_SECRET: "app-secret",
  TENANT_ID: "consumers",
  JWT_SIGNING_KEY: "j".repeat(48),
  TOKEN_ENCRYPTION_KEY: "e".repeat(48),
}

const b64url = (value: string | Buffer): string => Buffer.from(value).toString("base64url")
// Microsoft's id token. somamcp reads it straight from the token response over TLS and does not
// check its signature (OpenID Connect Core §3.1.3.7), so an unsigned one stands in.
const idToken = (claims: Record<string, unknown>): string =>
  `${b64url(JSON.stringify({ alg: "none" }))}.${b64url(JSON.stringify(claims))}.sig`

const USER = {
  oid: "00000000-0000-0000-0000-0000000000a1",
  tid: "9188040d-6c67-4c5b-b112-36a304b66dad",
  preferred_username: "someone@outlook.com",
  iss: "https://login.microsoftonline.com/9188040d-6c67-4c5b-b112-36a304b66dad/v2.0",
}

type Server = ReturnType<typeof createHttpServer>["server"]

/** Everything a client does up to a redeemable code, with Microsoft answering as `claims`. */
const signIn = async (server: Server, claims: Record<string, unknown>) => {
  const verifier = b64url(randomBytes(32))
  const registered = await server.fetch(
    new Request(`${BASE}/oauth/register`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ redirect_uris: [REDIRECT], token_endpoint_auth_method: "none" }),
    }),
  )
  const { client_id: clientId } = await registered.json()

  const authorize = new URL(`${BASE}/oauth/authorize`)
  for (const [k, v] of Object.entries({
    response_type: "code",
    client_id: clientId,
    redirect_uri: REDIRECT,
    code_challenge: b64url(createHash("sha256").update(verifier).digest()),
    code_challenge_method: "S256",
    state: "client-state",
    resource: `${BASE}/mcp`,
  }))
    authorize.searchParams.set(k, v)
  const consentPage = await (await server.fetch(new Request(authorize))).text()
  const transactionId = /name="transaction_id" value="([^"]+)"/.exec(consentPage)?.[1] ?? ""
  // Post where the form says, as a browser would: 2.7 fixed a double slash in this URL.
  const consentAction = /<form method="post" action="([^"]+)"/.exec(consentPage)?.[1] ?? ""

  const approved = await server.fetch(
    new Request(new URL(consentAction.replace(/&amp;/g, "&"), BASE), {
      method: "POST",
      headers: { "content-type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ transaction_id: transactionId, action: "approve" }),
    }),
  )
  const toMicrosoft = new URL(approved.headers.get("location") ?? "")

  microsoftClaims = claims
  const callback = await server.fetch(
    new Request(`${BASE}/oauth/callback?code=ms-code&state=${toMicrosoft.searchParams.get("state")}`),
  )
  return { registered, consentPage, approved, toMicrosoft, callback, clientId, verifier }
}

const token = async (server: Server, fields: Record<string, string>) => {
  const response = await server.fetch(
    new Request(`${BASE}/oauth/token`, {
      method: "POST",
      headers: { "content-type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams(fields),
    }),
  )
  return { status: response.status, body: await response.json() }
}

const toolCall = async (server: Server, bearer: string, name: string) => {
  const response = await server.fetch(
    new Request(`${BASE}/mcp`, {
      method: "POST",
      headers: {
        "content-type": "application/json",
        accept: "application/json, text/event-stream",
        authorization: `Bearer ${bearer}`,
      },
      body: JSON.stringify({ jsonrpc: "2.0", id: 1, method: "tools/call", params: { name, arguments: {} } }),
    }),
  )
  const line = (await response.text()).split("\n").find((l) => l.startsWith("data: "))
  return { status: response.status, text: line ? JSON.parse(line.slice(6)).result.content[0].text : "" }
}

let microsoftClaims: Record<string, unknown> = USER
let issued = 0

// Microsoft's token endpoint: a fresh Graph token per exchange, and the id token for the sign-in.
const microsoftTokens = (call: Call) => {
  const grant = new URLSearchParams(call.rawBody).get("grant_type")
  issued += 1
  return json({
    access_token: `graph-token-${issued}`,
    refresh_token: `ms-refresh-${issued}`,
    expires_in: 3600,
    token_type: "Bearer",
    scope: "openid profile email User.Read Tasks.Read Tasks.ReadWrite Tasks.Read.Shared Tasks.ReadWrite.Shared",
    ...(grant === "authorization_code" ? { id_token: idToken(microsoftClaims) } : {}),
  })
}

describe("a sign-in through HTTP mode's OAuth provider", () => {
  const net = new FakeNetwork()
  let store: string
  let http: ReturnType<typeof createHttpServer>
  const start = (extra: Record<string, string> = {}) => {
    http = createHttpServer(
      readHttpConfig({ ...env, MSTODO_HTTP_STORE: store, ...extra }).orThrow(new Error("expected a config")),
    )
  }

  beforeEach(() => {
    store = mkdtempSync(join(tmpdir(), "mstodo-sign-in-"))
    issued = 0
    resetNetwork(net)
    net.onTokenRefresh(microsoftTokens)
    net.on("GET", "/me/todo/lists", (call) =>
      json({ value: [{ id: `L-for-${call.headers.authorization?.slice("Bearer ".length)}`, displayName: "Tasks" }] }),
    )
  })
  afterEach(() => {
    http.provider.close()
    rmSync(store, { recursive: true, force: true })
    restoreNetwork(net)
  })

  it("signs a user in, calls tools with their Microsoft token, and refreshes", async () => {
    start()
    const flow = await signIn(http.server, USER)

    expect(flow.registered.status).toBe(201)
    expect(flow.consentPage).toContain('name="action" value="approve"')
    expect(flow.toMicrosoft.origin + flow.toMicrosoft.pathname).toBe(
      "https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize",
    )
    expect(flow.toMicrosoft.searchParams.get("redirect_uri")).toBe(`${BASE}/oauth/callback`)
    expect(flow.toMicrosoft.searchParams.get("scope")?.split(" ")).toEqual(
      expect.arrayContaining(["openid", "offline_access", "Tasks.ReadWrite"]),
    )

    const back = new URL(flow.callback.headers.get("location") ?? "")
    expect(back.origin + back.pathname).toBe(REDIRECT)
    expect(back.searchParams.get("state")).toBe("client-state")

    const redeemed = await token(http.server, {
      grant_type: "authorization_code",
      code: back.searchParams.get("code") ?? "",
      redirect_uri: REDIRECT,
      code_verifier: flow.verifier,
      client_id: flow.clientId,
    })
    expect(redeemed.status).toBe(200)
    // The client gets the server's own tokens, never Microsoft's.
    expect(JSON.stringify(redeemed.body)).not.toMatch(/graph-token|ms-refresh/)

    expect((await toolCall(http.server, redeemed.body.access_token, "auth-status")).text).toMatch(
      /^Signed in as someone@outlook\.com\. The Microsoft token expires at /,
    )
    expect((await toolCall(http.server, redeemed.body.access_token, "get-task-lists")).text).toContain(
      "ID: L-for-graph-token-1",
    )

    const refreshed = await token(http.server, {
      grant_type: "refresh_token",
      refresh_token: redeemed.body.refresh_token,
      client_id: flow.clientId,
    })
    expect(refreshed.status).toBe(200)
    expect((await toolCall(http.server, refreshed.body.access_token, "get-task-lists")).text).toContain(
      "ID: L-for-graph-token-2",
    )

    // The spent refresh token is refused, as Microsoft's own rotation is.
    const reused = await token(http.server, {
      grant_type: "refresh_token",
      refresh_token: redeemed.body.refresh_token,
      client_id: flow.clientId,
    })
    expect(reused.status).toBe(400)
  })

  it("posts consent to a single-slash /oauth/consent", async () => {
    start()
    const flow = await signIn(http.server, USER)

    expect(flow.consentPage).toMatch(/<form method="post" action="[^"]*(?<!\/)\/oauth\/consent"/)
    expect(flow.approved.status).toBe(302)
  })

  it("refuses PKCE's plain method at authorize", async () => {
    start()
    const registered = await http.server.fetch(
      new Request(`${BASE}/oauth/register`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ redirect_uris: [REDIRECT], token_endpoint_auth_method: "none" }),
      }),
    )
    const { client_id: clientId } = await registered.json()
    const authorize = new URL(`${BASE}/oauth/authorize`)
    for (const [k, v] of Object.entries({
      response_type: "code",
      client_id: clientId,
      redirect_uri: REDIRECT,
      code_challenge: "a".repeat(43),
      code_challenge_method: "plain",
      state: "s",
    }))
      authorize.searchParams.set(k, v)

    const response = await http.server.fetch(new Request(authorize))
    const back = new URL(response.headers.get("location") ?? "")

    expect(back.origin + back.pathname).toBe(REDIRECT)
    expect(back.searchParams.get("error")).toBe("invalid_request")
  })

  it("keeps a code usable after a redemption with the wrong verifier", async () => {
    start()
    const flow = await signIn(http.server, USER)
    const code = new URL(flow.callback.headers.get("location") ?? "").searchParams.get("code") ?? ""
    const redeem = (verifier: string) =>
      token(http.server, {
        grant_type: "authorization_code",
        code,
        redirect_uri: REDIRECT,
        code_verifier: verifier,
        client_id: flow.clientId,
      })

    expect((await redeem(b64url(randomBytes(32)))).status).toBe(400)
    expect((await redeem(flow.verifier)).status).toBe(200)
  })

  it("refuses a sign-in the allow-list does not name, back on the client's redirect URI", async () => {
    start({ MSTODO_HTTP_ALLOWED_USERS: "someone-else@outlook.com" })
    const flow = await signIn(http.server, USER)

    const back = new URL(flow.callback.headers.get("location") ?? "")
    expect(back.origin + back.pathname).toBe(REDIRECT)
    expect(back.searchParams.get("error")).toBe("access_denied")
    expect(back.searchParams.get("code")).toBeNull()
  })
})
