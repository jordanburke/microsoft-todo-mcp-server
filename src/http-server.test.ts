import { createHash } from "node:crypto"
import { createServer, request as httpRequest, type Server } from "node:http"

import { afterEach, describe, expect, it } from "vitest"

import { createRemoteMcpApplication, type HttpServerConfig, loadHttpServerConfig } from "./http-server.js"

const publicOrigin = "https://todo.example.com"
const resourceUrl = `${publicOrigin}/mcp`
const redirectUri = "https://chatgpt.example.test/oauth/callback"
const password = "correct horse battery staple"

describe("remote MCP HTTP server", () => {
  let server: Server | undefined

  afterEach(async () => {
    if (server) {
      await new Promise<void>((resolve, reject) => {
        server?.close((error) => (error ? reject(error) : resolve()))
      })
      server = undefined
    }
  })

  it("rejects unsafe public configuration", () => {
    expect(() =>
      loadHttpServerConfig({
        MCP_PUBLIC_URL: "http://todo.example.com",
        MCP_OAUTH_PASSWORD: password,
      }),
    ).toThrow("must use HTTPS")

    expect(() =>
      loadHttpServerConfig({
        MCP_PUBLIC_URL: "https://todo.example.com/base",
        MCP_OAUTH_PASSWORD: password,
      }),
    ).toThrow("without a path")
  })

  it("completes discovery, PKCE authorization, refresh, and an MCP request", async () => {
    const config: HttpServerConfig = {
      host: "127.0.0.1",
      password,
      port: 43123,
      publicUrl: new URL(publicOrigin),
      trustProxy: 0,
    }
    const { app } = createRemoteMcpApplication(config)
    server = createServer(app)
    await new Promise<void>((resolve) => server?.listen(0, "127.0.0.1", resolve))
    const address = server.address()
    if (!address || typeof address === "string") throw new Error("Test server did not bind to TCP")
    const requestOrigin = `http://127.0.0.1:${address.port}`

    const proxiedHealthResponse = await getWithHost(requestOrigin, "/healthz", "todo.example.com")
    expect(proxiedHealthResponse.status).toBe(200)

    const invalidHostResponse = await getWithHost(requestOrigin, "/healthz", "untrusted.example.com")
    expect(invalidHostResponse.status).toBe(403)
    expect(JSON.parse(invalidHostResponse.body)).toMatchObject({
      error: { message: "Invalid Host: untrusted.example.com" },
    })

    const protectedMetadataResponse = await fetch(`${requestOrigin}/.well-known/oauth-protected-resource/mcp`)
    expect(protectedMetadataResponse.status).toBe(200)
    expect(await protectedMetadataResponse.json()).toMatchObject({
      resource: resourceUrl,
      authorization_servers: [`${publicOrigin}/`],
      scopes_supported: ["mcp:tools"],
    })

    const authorizationMetadataResponse = await fetch(`${requestOrigin}/.well-known/oauth-authorization-server`)
    expect(await authorizationMetadataResponse.json()).toMatchObject({
      issuer: `${publicOrigin}/`,
      authorization_endpoint: `${publicOrigin}/authorize`,
      token_endpoint: `${publicOrigin}/token`,
      registration_endpoint: `${publicOrigin}/register`,
      code_challenge_methods_supported: ["S256"],
    })

    const registrationResponse = await fetch(`${requestOrigin}/register`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        redirect_uris: [redirectUri],
        token_endpoint_auth_method: "none",
        grant_types: ["authorization_code", "refresh_token"],
        response_types: ["code"],
        client_name: "ChatGPT integration test",
        scope: "mcp:tools",
      }),
    })
    expect(registrationResponse.status).toBe(201)
    const client = (await registrationResponse.json()) as { client_id: string }

    const verifier = "test-verifier-that-is-long-enough-for-pkce-0123456789"
    const challenge = createHash("sha256").update(verifier).digest("base64url")
    const authorizeUrl = new URL(`${requestOrigin}/authorize`)
    authorizeUrl.search = new URLSearchParams({
      response_type: "code",
      client_id: client.client_id,
      redirect_uri: redirectUri,
      code_challenge: challenge,
      code_challenge_method: "S256",
      state: "test-state",
      scope: "mcp:tools",
      resource: resourceUrl,
    }).toString()

    const authorizationPageResponse = await fetch(authorizeUrl)
    expect(authorizationPageResponse.status).toBe(200)
    const authorizationPage = await authorizationPageResponse.text()
    expect(authorizationPage).toContain("授权 Microsoft To Do")
    const requestId = authorizationPage.match(/name="request_id" value="([^"]+)"/)?.[1]
    expect(requestId).toBeTruthy()

    const rejectedApprovalResponse = await fetch(`${requestOrigin}/oauth/approve`, {
      method: "POST",
      redirect: "manual",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        request_id: requestId!,
        decision: "approve",
        password: "wrong-password",
      }),
    })
    expect(rejectedApprovalResponse.status).toBe(401)

    const approvalResponse = await fetch(`${requestOrigin}/oauth/approve`, {
      method: "POST",
      redirect: "manual",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        request_id: requestId!,
        decision: "approve",
        password,
      }),
    })
    expect(approvalResponse.status).toBe(302)
    const callback = new URL(approvalResponse.headers.get("location")!)
    expect(callback.origin + callback.pathname).toBe(redirectUri)
    expect(callback.searchParams.get("state")).toBe("test-state")
    const code = callback.searchParams.get("code")
    expect(code).toBeTruthy()

    const tokenResponse = await fetch(`${requestOrigin}/token`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "authorization_code",
        client_id: client.client_id,
        code: code!,
        code_verifier: verifier,
        redirect_uri: redirectUri,
        resource: resourceUrl,
      }),
    })
    expect(tokenResponse.status).toBe(200)
    const tokens = (await tokenResponse.json()) as {
      access_token: string
      refresh_token: string
      scope: string
    }
    expect(tokens.scope).toBe("mcp:tools")
    expect(tokens.access_token).toBeTruthy()
    expect(tokens.refresh_token).toBeTruthy()

    const unauthorizedMcpResponse = await initializeMcp(requestOrigin)
    expect(unauthorizedMcpResponse.status).toBe(401)
    expect(unauthorizedMcpResponse.headers.get("www-authenticate")).toContain(
      `resource_metadata="${publicOrigin}/.well-known/oauth-protected-resource/mcp"`,
    )

    const authorizedMcpResponse = await initializeMcp(requestOrigin, tokens.access_token)
    expect(authorizedMcpResponse.status).toBe(200)
    expect(await authorizedMcpResponse.json()).toMatchObject({
      jsonrpc: "2.0",
      result: {
        serverInfo: { name: "mstodo", version: "1.1.3" },
      },
      id: 1,
    })

    const toolsResponse = await callMcp(requestOrigin, tokens.access_token, {
      jsonrpc: "2.0",
      method: "tools/list",
      params: {},
      id: 2,
    })
    expect(toolsResponse.status).toBe(200)
    const toolsBody = (await toolsResponse.json()) as {
      result: { tools: Array<{ name: string }> }
    }
    expect(toolsBody.result.tools.map((tool) => tool.name)).toEqual(
      expect.arrayContaining(["auth-status", "get-task-lists", "create-task", "update-task"]),
    )

    const refreshResponse = await fetch(`${requestOrigin}/token`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "refresh_token",
        client_id: client.client_id,
        refresh_token: tokens.refresh_token,
        scope: "mcp:tools",
        resource: resourceUrl,
      }),
    })
    expect(refreshResponse.status).toBe(200)
    const refreshedTokens = (await refreshResponse.json()) as {
      access_token: string
      refresh_token: string
    }
    expect(refreshedTokens.access_token).not.toBe(tokens.access_token)
    expect(refreshedTokens.refresh_token).not.toBe(tokens.refresh_token)

    const replayResponse = await fetch(`${requestOrigin}/token`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "refresh_token",
        client_id: client.client_id,
        refresh_token: tokens.refresh_token,
        resource: resourceUrl,
      }),
    })
    expect(replayResponse.status).toBe(400)
    expect(await replayResponse.json()).toMatchObject({ error: "invalid_grant" })
  })
})

function initializeMcp(origin: string, accessToken?: string): Promise<Response> {
  return callMcp(origin, accessToken, {
    jsonrpc: "2.0",
    method: "initialize",
    params: {
      protocolVersion: "2025-06-18",
      capabilities: {},
      clientInfo: { name: "integration-test", version: "1.0.0" },
    },
    id: 1,
  })
}

function callMcp(origin: string, accessToken: string | undefined, body: object): Promise<Response> {
  const headers: Record<string, string> = {
    Accept: "application/json, text/event-stream",
    "Content-Type": "application/json",
  }
  if (accessToken) headers.Authorization = `Bearer ${accessToken}`

  return fetch(`${origin}/mcp`, {
    method: "POST",
    headers,
    body: JSON.stringify(body),
  })
}

function getWithHost(origin: string, path: string, host: string): Promise<{ body: string; status: number }> {
  return new Promise((resolve, reject) => {
    const request = httpRequest(new URL(path, origin), { headers: { Host: host } }, (response) => {
      const chunks: Buffer[] = []
      response.on("data", (chunk: Buffer) => chunks.push(chunk))
      response.on("end", () => {
        resolve({ body: Buffer.concat(chunks).toString("utf8"), status: response.statusCode ?? 0 })
      })
    })
    request.on("error", reject)
    request.end()
  })
}
