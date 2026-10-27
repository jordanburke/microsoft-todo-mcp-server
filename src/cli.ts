#!/usr/bin/env node

import { startServer } from "./todo-index.js"

// Token discovery belongs to TokenManager, which reads MSTODO_TOKEN_FILE, then the
// per-user config directory, then tokens.json in the working directory, and logs the
// path it settled on. This entry point only starts the server.
console.error("Microsoft Todo MCP CLI")

startServer().catch((error) => {
  const errorMessage = error instanceof Error ? error.message : String(error)
  console.error("Error starting server:", errorMessage)
  process.exit(1)
})
