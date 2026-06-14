#!/usr/bin/env node

import { spawn } from "child_process"
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "fs"
import { homedir } from "os"
import { join } from "path"
import readline from "readline"

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
})

const question = (query: string): Promise<string> => {
  return new Promise((resolve) => rl.question(query, resolve))
}

async function setup() {
  console.log("🚀 Microsoft To Do MCP Server Setup")
  console.log("==================================\n")

  // Check if already configured
  const configDir =
    process.platform === "win32"
      ? join(process.env.APPDATA || join(homedir(), "AppData", "Roaming"), "microsoft-todo-mcp")
      : join(homedir(), ".config", "microsoft-todo-mcp")

  const tokenPath = join(configDir, "tokens.json")

  if (existsSync(tokenPath)) {
    const answer = await question("Tokens already exist. Reconfigure? (y/N): ")
    if (answer.toLowerCase() !== "y") {
      console.log("Setup cancelled.")
      process.exit(0)
    }
  }

  // Check for Azure app credentials
  const hasEnvFile = existsSync(".env")

  if (!hasEnvFile) {
    console.log("\n📋 Azure App Registration Required")
    console.log("You need to create an app registration in Azure Portal first.")
    console.log("\nSteps:")
    console.log("1. Go to https://portal.azure.com")
    console.log("2. Navigate to 'App registrations' and create a new registration")
    console.log("3. Set redirect URI to: http://localhost:3000/callback")
    console.log("4. Add these API permissions: Tasks.Read, Tasks.ReadWrite, User.Read")
    console.log("5. Create a client secret\n")

    const clientId = await question("Enter your CLIENT_ID: ")
    const clientSecret = await question("Enter your CLIENT_SECRET: ")
    const tenantId = (await question("Enter your TENANT_ID (press Enter for 'organizations'): ")) || "organizations"

    // Create .env file
    const envContent = `CLIENT_ID=${clientId}
CLIENT_SECRET=${clientSecret}
TENANT_ID=${tenantId}
REDIRECT_URI=http://localhost:3000/callback
`
    writeFileSync(".env", envContent)
    console.log("✅ Created .env file")
  }

  console.log("\n🔐 Starting authentication flow...")
  console.log("A browser window will open. Please sign in with your Microsoft account.\n")

  // Start the auth server
  const authProcess = spawn("node", ["dist/auth-server.js"], {
    stdio: "inherit",
    shell: true,
  })

  authProcess.on("close", async (code) => {
    if (code === 0) {
      console.log("\n✅ Authentication successful!")

      // Check if tokens were created
      const localTokens = join(process.cwd(), "tokens.json")
      if (existsSync(localTokens)) {
        // Move tokens to proper location and add client credentials
        const tokens = JSON.parse(readFileSync(localTokens, "utf8"))
        const env = readFileSync(".env", "utf8")

        const clientId = env.match(/CLIENT_ID=(.+)/)?.[1]
        const clientSecret = env.match(/CLIENT_SECRET=(.+)/)?.[1]
        const tenantId = env.match(/TENANT_ID=(.+)/)?.[1] || "organizations"

        // Store with credentials for future refreshes
        const enhancedTokens = {
          ...tokens,
          clientId,
          clientSecret,
          tenantId,
        }

        // Create directory if needed
        mkdirSync(configDir, { recursive: true })

        // Save to proper location
        writeFileSync(tokenPath, JSON.stringify(enhancedTokens, null, 2))

        console.log(`\n📁 Tokens saved to: ${tokenPath}`)

        // Update Claude config
        await updateClaudeConfig()

        console.log("\n🎉 Setup complete! Microsoft To Do MCP is ready to use.")
        console.log("Restart Claude Desktop to activate the integration.")
      }
    } else {
      console.error("\n❌ Authentication failed. Please try again.")
    }

    rl.close()
  })
}

async function updateClaudeConfig() {
  const claudeConfigPath =
    process.platform === "win32"
      ? join(process.env.APPDATA || "", "Claude", "claude_desktop_config.json")
      : process.platform === "darwin"
        ? join(homedir(), "Library", "Application Support", "Claude", "claude_desktop_config.json")
        : join(homedir(), ".config", "Claude", "claude_desktop_config.json")

  if (!existsSync(claudeConfigPath)) {
    console.log("\n⚠️  Claude config not found. Add this to your Claude desktop config manually:")
    console.log(
      JSON.stringify(
        {
          "microsoft-todo": {
            command: "npx",
            args: ["microsoft-todo-mcp-server"],
            env: {},
          },
        },
        null,
        2,
      ),
    )
    return
  }

  try {
    const config = JSON.parse(readFileSync(claudeConfigPath, "utf8"))

    // Add or update the microsoft-todo server config
    if (!config.mcpServers) {
      config.mcpServers = {}
    }

    config.mcpServers["microsoft-todo"] = {
      command: "npx",
      args: ["microsoft-todo-mcp-server"],
      env: {}, // No need for tokens in env anymore!
    }

    writeFileSync(claudeConfigPath, JSON.stringify(config, null, 2))
    console.log("\n✅ Updated Claude Desktop configuration")
  } catch (error) {
    console.error("\n⚠️  Could not update Claude config automatically:", error)
  }
}

// Run setup
setup().catch(console.error)
