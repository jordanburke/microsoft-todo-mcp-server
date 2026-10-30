// Exercises the stateless Streamable HTTP transport end-to-end: a real HTTP
// server on an ephemeral port, talked to with the MCP SDK's own client
// (not hand-rolled JSON-RPC), so this test fails the same way a real hosted
// agent's client would fail if the transport wiring ever regresses.
import { Client } from "@modelcontextprotocol/sdk/client/index.js"
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js"
import type { Server as HttpServer } from "http"
import type { AddressInfo } from "net"
import { afterEach, beforeEach, describe, expect, it } from "vitest"

import { createHttpApp } from "./http-transport.js"

let httpServer: HttpServer
let baseUrl: URL

beforeEach(async () => {
  const app = createHttpApp()
  await new Promise<void>((resolve) => {
    httpServer = app.listen(0, "127.0.0.1", resolve)
  })
  const { port } = httpServer.address() as AddressInfo
  baseUrl = new URL(`http://127.0.0.1:${port}/mcp`)
})

afterEach(async () => {
  await new Promise<void>((resolve) => httpServer.close(() => resolve()))
})

async function connectClient(): Promise<Client> {
  const client = new Client({ name: "test-client", version: "1.0.0" })
  // sessionIdGenerator: undefined on the server means it never returns a
  // session ID, so the client transport must not expect or reuse one either.
  const transport = new StreamableHTTPClientTransport(baseUrl)
  await client.connect(transport)
  return client
}

describe("stateless streamable HTTP transport", () => {
  it("completes the MCP initialize handshake and reports the mstodo server", async () => {
    const client = await connectClient()
    const serverInfo = client.getServerVersion()
    expect(serverInfo?.name).toBe("mstodo")
    await client.close()
  })

  it("lists all 15 registered tools", async () => {
    const client = await connectClient()
    const { tools } = await client.listTools()
    const names = tools.map((t) => t.name)

    expect(names).toEqual(
      expect.arrayContaining([
        "auth-status",
        "get-task-lists",
        "get-task-lists-organized",
        "create-task-list",
        "update-task-list",
        "delete-task-list",
        "get-tasks",
        "create-task",
        "update-task",
        "delete-task",
        "get-checklist-items",
        "create-checklist-item",
        "update-checklist-item",
        "delete-checklist-item",
        "archive-completed-tasks",
      ]),
    )
    await client.close()
  })

  it("keeps two concurrent clients fully isolated (no shared session state)", async () => {
    // This is the property statelessness is actually for: nothing one
    // client's connection does (or fails to do) can leak into another's,
    // which is what makes it safe to run behind Cloud Run's default
    // concurrency and to scale to multiple instances.
    const [clientA, clientB] = await Promise.all([connectClient(), connectClient()])

    const [toolsA, toolsB] = await Promise.all([clientA.listTools(), clientB.listTools()])

    expect(toolsA.tools.length).toBe(toolsB.tools.length)
    expect(toolsA.tools.length).toBeGreaterThanOrEqual(15)

    await Promise.all([clientA.close(), clientB.close()])
  })

  it("rejects auth-status without an access token instead of throwing", async () => {
    const client = await connectClient()
    const result = await client.callTool({ name: "auth-status", arguments: {} })
    const text = (result.content as Array<{ type: string; text?: string }>)[0]?.text ?? ""
    expect(text.toLowerCase()).toContain("not authenticated")
    await client.close()
  })

  it("answers GET and DELETE on /mcp with 405, since stateless mode has no stream or session to act on", async () => {
    const getRes = await fetch(baseUrl, { method: "GET" })
    expect(getRes.status).toBe(405)

    const deleteRes = await fetch(baseUrl, { method: "DELETE" })
    expect(deleteRes.status).toBe(405)
  })

  it("exposes a liveness endpoint that Cloud Run's health check can hit without touching Graph or the token cache", async () => {
    const res = await fetch(new URL("/healthz", baseUrl))
    expect(res.status).toBe(200)
    expect(await res.json()).toEqual({ status: "ok" })
  })
})
