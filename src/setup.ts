import { join } from "@std/path";

function homedir(): string {
  switch (Deno.build.os) {
    case "linux":
    case "darwin":
      return Deno.env.get("HOME") ?? "/";
    case "windows":
      return (
        (Deno.env.get("USERPROFILE") ??
          (Deno.env.get("HOMEDRIVE") ?? "") +
            (Deno.env.get("HOMEPATH") ?? "")) ||
        "/"
      );
    default:
      return "/";
  }
}

function existsSync(path: string): boolean {
  try {
    Deno.statSync(path);
    return true;
  } catch {
    return false;
  }
}

async function setup() {
  console.log("🚀 Microsoft To Do MCP Server Setup");
  console.log("==================================\n");

  // Check if already configured
  const configDir = Deno.build.os === "windows"
    ? join(
      Deno.env.get("APPDATA") || join(homedir(), "AppData", "Roaming"),
      "microsoft-todo-mcp",
    )
    : join(homedir(), ".config", "microsoft-todo-mcp");

  const tokenPath = join(configDir, "tokens.json");

  if (existsSync(tokenPath)) {
    const answer = prompt("Tokens already exist. Reconfigure? (y/N): ");
    if (answer?.toLowerCase() !== "y") {
      console.log("Setup cancelled.");
      Deno.exit(0);
    }
  }

  // Check for Azure app credentials
  const hasEnvFile = existsSync(".env");

  if (!hasEnvFile) {
    console.log("\n📋 Azure App Registration Required");
    console.log(
      "You need to create an app registration in Azure Portal first.",
    );
    console.log("\nSteps:");
    console.log("1. Go to https://portal.azure.com");
    console.log(
      "2. Navigate to 'App registrations' and create a new registration",
    );
    console.log("3. Set redirect URI to: http://localhost:3000/callback");
    console.log(
      "4. Add these API permissions: Tasks.Read, Tasks.ReadWrite, User.Read",
    );
    console.log("5. Create a client secret\n");

    const clientId = prompt("Enter your CLIENT_ID: ") || "";
    const clientSecret = prompt("Enter your CLIENT_SECRET: ") || "";
    const tenantId = prompt("Enter your TENANT_ID (press Enter for 'organizations'): ") ||
      "organizations";

    // Create .env file
    const envContent = `CLIENT_ID=${clientId}
CLIENT_SECRET=${clientSecret}
TENANT_ID=${tenantId}
REDIRECT_URI=http://localhost:3000/callback
`;
    Deno.writeTextFileSync(".env", envContent);
    console.log("✅ Created .env file");
  }

  console.log("\n🔐 Starting authentication flow...");
  console.log(
    "A browser window will open. Please sign in with your Microsoft account.\n",
  );

  // Start the auth server using Deno
  const command = new Deno.Command("deno", {
    args: [
      "run",
      "--allow-env",
      "--allow-read",
      "--allow-write",
      "--allow-net",
      join(Deno.cwd(), "src", "auth-server.ts"),
    ],
    stdin: "inherit",
    stdout: "inherit",
    stderr: "inherit",
  });

  const process = command.spawn();
  const status = await process.status;

  if (status.success) {
    console.log("\n✅ Authentication successful!");

    // Check if tokens were created
    const localTokens = join(Deno.cwd(), "tokens.json");
    if (existsSync(localTokens)) {
      // Move tokens to proper location and add client credentials
      const tokens = JSON.parse(Deno.readTextFileSync(localTokens));
      let env = "";
      try {
        env = Deno.readTextFileSync(".env");
      } catch {
        console.log("No .env file found");
      }

      const clientIdMatch = env.match(/CLIENT_ID=(.+)/);
      const clientSecretMatch = env.match(/CLIENT_SECRET=(.+)/);
      const tenantIdMatch = env.match(/TENANT_ID=(.+)/);

      const clientId = clientIdMatch?.[1];
      const clientSecret = clientSecretMatch?.[1];
      const tenantId = tenantIdMatch?.[1] || "organizations";

      // Store with credentials for future refreshes
      const enhancedTokens = {
        ...tokens,
        clientId,
        clientSecret,
        tenantId,
      };

      // Create directory if needed
      Deno.mkdirSync(configDir, { recursive: true });

      // Save to proper location
      Deno.writeTextFileSync(
        tokenPath,
        JSON.stringify(enhancedTokens, null, 2),
      );

      console.log(`\n📁 Tokens saved to: ${tokenPath}`);

      // Update Claude config
      updateClaudeConfig();

      console.log("\n🎉 Setup complete! Microsoft To Do MCP is ready to use.");
      console.log("Restart Claude Desktop to activate the integration.");
    }
  } else {
    console.error("\n❌ Authentication failed. Please try again.");
  }
}

function updateClaudeConfig() {
  const os = Deno.build.os;
  const claudeConfigPath = os === "windows"
    ? join(
      Deno.env.get("APPDATA") || "",
      "Claude",
      "claude_desktop_config.json",
    )
    : os === "darwin"
    ? join(
      homedir(),
      "Library",
      "Application Support",
      "Claude",
      "claude_desktop_config.json",
    )
    : join(homedir(), ".config", "Claude", "claude_desktop_config.json");

  if (!existsSync(claudeConfigPath)) {
    console.log(
      "\n⚠️  Claude config not found. Add this to your Claude desktop config manually:",
    );
    console.log(
      JSON.stringify(
        {
          "microsoft-todo": {
            command: "deno",
            args: [
              "run",
              "--allow-env",
              "--allow-read",
              "--allow-write",
              "--allow-net",
              join(Deno.cwd(), "src", "todo-index.ts"),
            ],
            env: {},
          },
        },
        null,
        2,
      ),
    );
    return;
  }

  try {
    const config = JSON.parse(Deno.readTextFileSync(claudeConfigPath));

    // Add or update the microsoft-todo server config
    if (!config.mcpServers) {
      config.mcpServers = {};
    }

    config.mcpServers["microsoft-todo"] = {
      command: "deno",
      args: [
        "run",
        "--allow-env",
        "--allow-read",
        "--allow-write",
        "--allow-net",
        join(Deno.cwd(), "src", "todo-index.ts"),
      ],
      env: {}, // No need for tokens in env anymore!
    };

    Deno.writeTextFileSync(claudeConfigPath, JSON.stringify(config, null, 2));
    console.log("\n✅ Updated Claude Desktop configuration");
  } catch (error) {
    console.error("\n⚠️  Could not update Claude config automatically:", error);
  }
}

// Run setup
setup().catch(console.error);
