#!/usr/bin/env node

// Writes an mcp.json whose server entry points at the token file. It holds no tokens:
// copying the access and refresh tokens into the config spread secrets into another
// app's file, and without the client credentials those copies could not be renewed.
import fs from "fs"
import path from "path"

export type McpConfig = {
  mcpServers: Record<string, { command: string; args: string[]; env: Record<string, string> }>
}

/** The config for a token file at `tokenPath`, which the server reads and renews itself. */
export const mcpConfigFor = (tokenPath: string): McpConfig => ({
  mcpServers: {
    microsoftTodo: {
      command: "npx",
      args: ["--yes", "microsoft-todo-mcp-server"],
      env: { MSTODO_TOKEN_FILE: path.resolve(tokenPath) },
    },
  },
})

/** What a token file lacks for the server to renew it, as readable warnings. */
export const tokenFileWarnings = (tokens: Record<string, unknown>): string[] =>
  [
    !tokens.refreshToken && "has no refreshToken, so the server cannot renew it",
    (!tokens.clientId || !tokens.clientSecret) &&
      "has no clientId/clientSecret, so renewal also needs CLIENT_ID and CLIENT_SECRET in the server's env",
  ].filter((w): w is string => typeof w === "string")

/** Read the token file and write the config; returns the warnings about the token file. */
export function createMcpConfig(tokenPath: string, outputPath: string): string[] {
  const tokens = JSON.parse(fs.readFileSync(tokenPath, "utf8"))
  fs.writeFileSync(outputPath, JSON.stringify(mcpConfigFor(tokenPath), null, 2), "utf8")
  return tokenFileWarnings(tokens)
}

function main(): void {
  const tokenPath = process.argv[2] || path.join(process.cwd(), "tokens.json")
  const outputPath = process.argv[3] || path.join(process.cwd(), "mcp.json")

  console.log(`Reading tokens from: ${tokenPath}`)
  console.log(`Writing config to: ${outputPath}`)

  try {
    const warnings = createMcpConfig(tokenPath, outputPath)
    console.log("MCP configuration file created successfully!")
    console.log(`It points the server at ${path.resolve(tokenPath)}; keep that file where it is.`)
    for (const warning of warnings) console.warn(`Warning: the token file ${warning}.`)
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error)
    console.error("Error creating MCP config:", errorMessage)
    process.exit(1)
  }
}

if (import.meta.url === `file://${process.argv[1]}`) main()
