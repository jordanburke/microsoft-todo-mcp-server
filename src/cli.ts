#!/usr/bin/env node

import fs from "fs"
import path from "path"
import { fileURLToPath } from "url"

import { startServer } from "./todo-index.js"

// Get the directory path for the current module
const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

// Check for tokens in environment variables
let accessToken = process.env.MS_TODO_ACCESS_TOKEN
let refreshToken = process.env.MS_TODO_REFRESH_TOKEN

// Define token file path
const TOKEN_FILE_PATH = process.env.MSTODO_TOKEN_FILE || path.join(process.cwd(), "tokens.json")

// Log startup info
console.error("Microsoft Todo MCP CLI")
console.error(`Looking for tokens in: ${TOKEN_FILE_PATH}`)

// Check if tokens are missing from environment but available in file
if ((!accessToken || !refreshToken) && fs.existsSync(TOKEN_FILE_PATH)) {
  try {
    console.error("Reading tokens from file...")
    const tokenData = JSON.parse(fs.readFileSync(TOKEN_FILE_PATH, "utf8"))

    // If we found tokens in the file, use them
    if (!accessToken && tokenData.accessToken) {
      accessToken = tokenData.accessToken
      console.error("Using access token from file")
    }

    if (!refreshToken && tokenData.refreshToken) {
      refreshToken = tokenData.refreshToken
      console.error("Using refresh token from file")
    }
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error)
    console.error("Error reading token file:", errorMessage)
  }
}

// Start the MCP server with the available tokens
startServer({
  accessToken,
  refreshToken,
  tokenFilePath: TOKEN_FILE_PATH,
}).catch((error) => {
  const errorMessage = error instanceof Error ? error.message : String(error)
  console.error("Error starting server:", errorMessage)
  process.exit(1)
})
