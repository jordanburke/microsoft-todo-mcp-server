#!/usr/bin/env node

import { startServer } from "./todo-index.js"

// Log startup info
console.error("Microsoft Todo MCP CLI")
console.error("Token management handled by TokenManager (stored in AppData/microsoft-todo-mcp/)")

// Start the MCP server - token management is handled internally by TokenManager
startServer().catch((error) => {
  const errorMessage = error instanceof Error ? error.message : String(error)
  console.error("Error starting server:", errorMessage)
  process.exit(1)
})
