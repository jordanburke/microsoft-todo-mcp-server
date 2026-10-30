// Stateless Streamable HTTP transport for the MCP server.
//
// Why stateless: this transport is meant to run on Cloud Run (see
// deploy/README.md), where any request can land on any instance and multiple
// requests can run concurrently on the same instance. The MCP Streamable HTTP
// spec supports an optional session ID that lets a server keep per-connection
// state (SSE resumability, etc.) between requests. We deliberately opt out of
// that (`sessionIdGenerator: undefined`) and instead give every single
// request its own McpServer + StreamableHTTPServerTransport pair that is
// created, used once, and thrown away. Nothing about a request depends on
// anything left behind by a previous one, so:
//   - two concurrent requests on the same instance never share transport
//     state and can't cross-talk
//   - it doesn't matter which Cloud Run instance (or how many of them)
//     handles a given request
// The only state that legitimately needs to survive across requests and
// instances is the OAuth token cache, which is handled separately by
// token-store.ts, not by anything in this file.
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js"
import express, { type Express, type Request, type Response } from "express"
import type { Server as HttpServer } from "http"

import { createTodoServer } from "./todo-index.js"

function methodNotAllowed(_req: Request, res: Response): void {
  // Stateless mode has no server-initiated stream and no session to resume,
  // so GET (open a stream) and DELETE (end a session) never apply.
  res.status(405).json({
    jsonrpc: "2.0",
    error: {
      code: -32000,
      message: "Method not allowed. This server runs the stateless Streamable HTTP transport (POST only).",
    },
    id: null,
  })
}

export function createHttpApp(): Express {
  const app = express()
  app.use(express.json())

  // Cloud Run (and any load balancer in front of it) needs a cheap liveness
  // check that never touches Microsoft Graph or the token cache.
  app.get("/healthz", (_req: Request, res: Response) => {
    res.status(200).json({ status: "ok" })
  })

  app.post("/mcp", async (req: Request, res: Response) => {
    // A brand-new server + transport per request is the deliberate choice
    // described above -- do not hoist these above the handler.
    const server = createTodoServer()

    try {
      const transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: undefined,
      })

      res.on("close", () => {
        transport.close()
        server.close()
      })

      await server.connect(transport)
      await transport.handleRequest(req, res, req.body)
    } catch (error) {
      console.error("Error handling MCP request:", error)
      if (!res.headersSent) {
        res.status(500).json({
          jsonrpc: "2.0",
          error: { code: -32603, message: "Internal server error" },
          id: null,
        })
      }
    }
  })

  app.get("/mcp", methodNotAllowed)
  app.delete("/mcp", methodNotAllowed)

  return app
}

export function startHttpServer(port: number): Promise<HttpServer> {
  const app = createHttpApp()
  return new Promise((resolve, reject) => {
    // Cloud Run routes traffic to the container on all interfaces; binding to
    // 0.0.0.0 (not the express default of localhost-only in some setups) is
    // required for the health check and real traffic to reach the process.
    const httpServer = app.listen(port, "0.0.0.0", () => resolve(httpServer))
    httpServer.on("error", reject)
  })
}
