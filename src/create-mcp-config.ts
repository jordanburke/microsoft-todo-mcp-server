import { join } from "@std/path";

// Define paths
const tokenPath = Deno.args[0] || join(Deno.cwd(), "tokens.json");
const outputPath = Deno.args[1] || join(Deno.cwd(), "mcp.json");

console.log(`Reading tokens from: ${tokenPath}`);
console.log(`Writing config to: ${outputPath}`);

try {
  // Read the tokens
  const tokenData = JSON.parse(Deno.readTextFileSync(tokenPath));

  // Create the MCP config - only include the actual tokens
  const mcpConfig = {
    mcpServers: {
      microsoftTodo: {
        command: "deno",
        args: [
          "run",
          "--allow-env",
          "--allow-read",
          "--allow-write",
          "--allow-net",
          join(Deno.cwd(), "src", "todo-index.ts"),
        ],
        env: {
          MS_TODO_ACCESS_TOKEN: tokenData.accessToken,
          MS_TODO_REFRESH_TOKEN: tokenData.refreshToken,
        },
      },
    },
  };

  // Write the config
  Deno.writeTextFileSync(outputPath, JSON.stringify(mcpConfig, null, 2));

  console.log("MCP configuration file created successfully!");
  console.log(
    "You can now use the service with Claude or Cursor by referencing this mcp.json file.",
  );
} catch (error: unknown) {
  const errorMessage = error instanceof Error ? error.message : String(error);
  console.error("Error creating MCP config:", errorMessage);
  Deno.exit(1);
}
