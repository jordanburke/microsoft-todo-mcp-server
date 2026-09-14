#!/usr/bin/env node

import { requireBearerAuth } from "@modelcontextprotocol/sdk/server/auth/middleware/bearerAuth.js"
import { getOAuthProtectedResourceMetadataUrl, mcpAuthRouter } from "@modelcontextprotocol/sdk/server/auth/router.js"
import { createMcpExpressApp } from "@modelcontextprotocol/sdk/server/express.js"
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js"
import dotenv from "dotenv"
import express, { type Request, type Response } from "express"

import { MCP_SCOPE, SingleUserOAuthProvider } from "./remote-auth.js"
import { createTodoServer } from "./todo-index.js"

dotenv.config()

export interface HttpServerConfig {
  host: string
  password: string
  port: number
  publicUrl: URL
  trustProxy: number
}

export interface RemoteMcpApplication {
  app: ReturnType<typeof createMcpExpressApp>
  provider: SingleUserOAuthProvider
  resourceUrl: URL
}

export function loadHttpServerConfig(environment: NodeJS.ProcessEnv = process.env): HttpServerConfig {
  const publicUrlValue = environment.MCP_PUBLIC_URL
  const password = environment.MCP_OAUTH_PASSWORD
  if (!publicUrlValue) {
    throw new Error("MCP_PUBLIC_URL is required, for example https://todo.example.com")
  }
  if (!password || password.length < 12) {
    throw new Error("MCP_OAUTH_PASSWORD is required and must contain at least 12 characters")
  }

  const publicUrl = new URL(publicUrlValue)
  const isLoopback = ["localhost", "127.0.0.1", "[::1]"].includes(publicUrl.hostname)
  if (publicUrl.protocol !== "https:" && !(isLoopback && publicUrl.protocol === "http:")) {
    throw new Error("MCP_PUBLIC_URL must use HTTPS unless it points to a loopback address")
  }
  if (publicUrl.pathname !== "/" || publicUrl.search || publicUrl.hash) {
    throw new Error("MCP_PUBLIC_URL must be an origin without a path, query, or fragment")
  }

  const port = parsePositiveInteger(environment.MCP_PORT ?? "3001", "MCP_PORT")
  const trustProxy = parseNonNegativeInteger(environment.MCP_TRUST_PROXY ?? "0", "MCP_TRUST_PROXY")

  return {
    host: environment.MCP_HOST || "127.0.0.1",
    password,
    port,
    publicUrl,
    trustProxy,
  }
}

export function createRemoteMcpApplication(config: HttpServerConfig): RemoteMcpApplication {
  const resourceUrl = new URL("/mcp", config.publicUrl)
  const provider = new SingleUserOAuthProvider(config.password, resourceUrl)
  const allowedHosts = [...new Set([config.publicUrl.hostname, "localhost", "127.0.0.1", "[::1]"])]
  const app = createMcpExpressApp({ host: config.host, allowedHosts })

  if (config.trustProxy > 0) {
    app.set("trust proxy", config.trustProxy)
  }

  app.get("/", (_req, res) => {
    res.json({
      name: "Microsoft To Do MCP",
      mcp_endpoint: resourceUrl.href,
      authorization_server_metadata: new URL("/.well-known/oauth-authorization-server", config.publicUrl).href,
      protected_resource_metadata: getOAuthProtectedResourceMetadataUrl(resourceUrl),
    })
  })

  app.get("/healthz", (_req, res) => {
    res.json({ status: "ok", transport: "streamable-http", oauth: true })
  })

  app.use(
    mcpAuthRouter({
      provider,
      issuerUrl: config.publicUrl,
      resourceServerUrl: resourceUrl,
      scopesSupported: [MCP_SCOPE],
      resourceName: "Microsoft To Do MCP",
      serviceDocumentationUrl: config.publicUrl,
    }),
  )

  app.post("/oauth/approve", express.urlencoded({ extended: false, limit: "8kb" }), (req: Request, res: Response) => {
    const requestId = typeof req.body.request_id === "string" ? req.body.request_id : ""
    const decision = req.body.decision === "approve"
    const password = typeof req.body.password === "string" ? req.body.password : undefined

    if (!requestId) {
      res.status(400).send("Missing authorization request ID")
      return
    }

    try {
      const redirect = provider.completeAuthorization({
        requestId,
        approved: decision,
        password,
      })
      res.redirect(302, redirect.href)
    } catch (error) {
      const message = error instanceof Error ? error.message : "Authorization failed"
      res.status(401).type("html").send(`<!doctype html>
<html lang="zh-CN"><meta charset="utf-8"><title>授权失败</title>
<body><h1>授权失败</h1><p>${escapeHtml(message)}</p><p>请返回上一页重试。</p></body></html>`)
    }
  })

  const bearerAuth = requireBearerAuth({
    verifier: provider,
    requiredScopes: [MCP_SCOPE],
    resourceMetadataUrl: getOAuthProtectedResourceMetadataUrl(resourceUrl),
  })

  app.post("/mcp", bearerAuth, async (req: Request, res: Response) => {
    const server = createTodoServer()
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: undefined,
      enableJsonResponse: true,
    })

    try {
      await server.connect(transport)
      await transport.handleRequest(req, res, req.body)
    } catch (error) {
      console.error("Error handling remote MCP request:", error)
      if (!res.headersSent) {
        res.status(500).json({
          jsonrpc: "2.0",
          error: { code: -32603, message: "Internal server error" },
          id: null,
        })
      }
    } finally {
      await transport.close().catch(() => undefined)
      await server.close().catch(() => undefined)
    }
  })

  const methodNotAllowed = (_req: Request, res: Response) => {
    res
      .set("Allow", "POST")
      .status(405)
      .json({
        jsonrpc: "2.0",
        error: { code: -32000, message: "Method not allowed" },
        id: null,
      })
  }
  app.get("/mcp", bearerAuth, methodNotAllowed)
  app.delete("/mcp", bearerAuth, methodNotAllowed)

  return { app, provider, resourceUrl }
}

export function startHttpServer(config = loadHttpServerConfig()): void {
  const { app, resourceUrl } = createRemoteMcpApplication(config)
  app.listen(config.port, config.host, (error?: Error) => {
    if (error) {
      console.error("Failed to start remote MCP server:", error)
      process.exitCode = 1
      return
    }

    console.error(`Microsoft To Do MCP is listening on http://${config.host}:${config.port}`)
    console.error(`Public MCP URL: ${resourceUrl.href}`)
  })
}

function parsePositiveInteger(value: string, name: string): number {
  const parsed = Number(value)
  if (!Number.isInteger(parsed) || parsed <= 0 || parsed > 65535) {
    throw new Error(`${name} must be an integer between 1 and 65535`)
  }
  return parsed
}

function parseNonNegativeInteger(value: string, name: string): number {
  const parsed = Number(value)
  if (!Number.isInteger(parsed) || parsed < 0) {
    throw new Error(`${name} must be a non-negative integer`)
  }
  return parsed
}

function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, (character) => {
    const entities: Record<string, string> = {
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#039;",
    }
    return entities[character]
  })
}

if (import.meta.url === `file://${process.argv[1]}`) {
  try {
    startHttpServer()
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error)
    console.error(`Unable to start remote MCP server: ${message}`)
    process.exit(1)
  }
}
