import { startServer } from "./todo-index.ts";
import { join } from "@std/path";

// Check for tokens in environment variables
let accessToken = Deno.env.get("MS_TODO_ACCESS_TOKEN") || undefined;
let refreshToken = Deno.env.get("MS_TODO_REFRESH_TOKEN") || undefined;

// Define token file path
const TOKEN_FILE_PATH = Deno.env.get("MSTODO_TOKEN_FILE") || join(Deno.cwd(), "tokens.json");

// Log startup info
console.error("Microsoft Todo MCP CLI");
console.error(`Looking for tokens in: ${TOKEN_FILE_PATH}`);

// Check if tokens are missing from environment but available in file
if (!accessToken || !refreshToken) {
  try {
    const stat = await Deno.stat(TOKEN_FILE_PATH);
    if (stat.isFile) {
      console.error("Reading tokens from file...");
      const tokenData = JSON.parse(Deno.readTextFileSync(TOKEN_FILE_PATH));

      if (!accessToken && tokenData.accessToken) {
        accessToken = tokenData.accessToken;
        console.error("Using access token from file");
      }

      if (!refreshToken && tokenData.refreshToken) {
        refreshToken = tokenData.refreshToken;
        console.error("Using refresh token from file");
      }
    }
  } catch {
    console.error("No token file found at:", TOKEN_FILE_PATH);
  }
}

// Start the MCP server with the available tokens
startServer({
  accessToken,
  refreshToken,
  tokenFilePath: TOKEN_FILE_PATH,
}).catch((error: unknown) => {
  const errorMessage = error instanceof Error ? error.message : String(error);
  console.error("Error starting server:", errorMessage);
  Deno.exit(1);
});
