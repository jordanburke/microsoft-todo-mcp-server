// The local (stdio) server: the To Do tools, with Graph tokens from the token manager.
import dotenv from "dotenv"

import { localCredentials } from "./local-credentials.js"
import { createTodoServer } from "./todo-server.js"

// Load environment variables
dotenv.config({ quiet: true })

// Log the current working directory
console.error("Current working directory:", process.cwd())

// Exported so specs can drive the tools in-process, without a build or a stdio child.
export const server = createTodoServer({ credentials: localCredentials })

// Main function to start the server
export async function startServer(): Promise<void> {
  try {
    // Start the server
    await server.start({ transportType: "stdio" })

    console.error("Server started and listening")
  } catch (error) {
    console.error("Error starting server:", error)
    throw error
  }
}
