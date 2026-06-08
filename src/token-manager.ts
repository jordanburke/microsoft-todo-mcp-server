// src/token-manager.ts
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

function readTextFileSync(path: string): string {
  return Deno.readTextFileSync(path);
}

function writeTextFileSync(path: string, data: string): void {
  Deno.writeTextFileSync(path, data);
}

function mkdirSync(path: string, options: { recursive: boolean }): void {
  Deno.mkdirSync(path, { recursive: options.recursive });
}

interface TokenData {
  accessToken: string;
  refreshToken: string;
  expiresAt: number;
}

interface StoredTokenData extends TokenData {
  clientId?: string;
  clientSecret?: string;
  tenantId?: string;
}

export class TokenManager {
  private tokenFilePath: string;
  private currentTokens: StoredTokenData | null = null;

  constructor() {
    // Store tokens in a consistent location across platforms
    const configDir = Deno.build.os === "windows"
      ? join(
        Deno.env.get("APPDATA") || join(homedir(), "AppData", "Roaming"),
        "microsoft-todo-mcp",
      )
      : join(homedir(), ".config", "microsoft-todo-mcp");

    // Create directory if it doesn't exist
    if (!existsSync(configDir)) {
      mkdirSync(configDir, { recursive: true });
    }

    this.tokenFilePath = join(configDir, "tokens.json");
    console.error(`Token file path: ${this.tokenFilePath}`);
  }

  // Try to get tokens from multiple sources
  async getTokens(): Promise<TokenData | null> {
    // 1. Check environment variables first (for backward compatibility)
    const envAccessToken = Deno.env.get("MS_TODO_ACCESS_TOKEN");
    const envRefreshToken = Deno.env.get("MS_TODO_REFRESH_TOKEN");

    if (envAccessToken && envRefreshToken) {
      const envTokens: TokenData = {
        accessToken: envAccessToken,
        refreshToken: envRefreshToken,
        expiresAt: Date.now() + 3600 * 1000, // Assume 1 hour if not specified
      };

      // Check if expired
      if (Date.now() > envTokens.expiresAt) {
        // Try to refresh
        const refreshed = await this.refreshToken(envTokens.refreshToken);
        if (refreshed) {
          return refreshed;
        }
      }
      return envTokens;
    }

    // 2. Check stored token file
    if (existsSync(this.tokenFilePath)) {
      try {
        const data = readTextFileSync(this.tokenFilePath);
        this.currentTokens = JSON.parse(data);

        if (this.currentTokens) {
          // Check if expired
          if (Date.now() > this.currentTokens.expiresAt) {
            // Try to refresh
            const refreshed = await this.refreshToken(
              this.currentTokens.refreshToken,
            );
            if (refreshed) {
              return refreshed;
            }
          }
          return this.currentTokens;
        }
      } catch (error) {
        console.error("Error reading token file:", error);
      }
    }

    // 3. Check legacy token file location
    const legacyPath = join(Deno.cwd(), "tokens.json");
    if (existsSync(legacyPath)) {
      try {
        const data = readTextFileSync(legacyPath);
        const tokens = JSON.parse(data);

        // Migrate to new location
        this.saveTokens(tokens);

        return tokens;
      } catch (error) {
        console.error("Error reading legacy token file:", error);
      }
    }

    return null;
  }

  async refreshToken(refreshToken: string): Promise<TokenData | null> {
    try {
      // Get client credentials from stored tokens or environment
      const clientId = this.currentTokens?.clientId || Deno.env.get("CLIENT_ID");
      const clientSecret = this.currentTokens?.clientSecret || Deno.env.get("CLIENT_SECRET");
      const tenantId = this.currentTokens?.tenantId ||
        Deno.env.get("TENANT_ID") ||
        "organizations";

      if (!clientId || !clientSecret) {
        console.error("Missing client credentials for token refresh");
        return null;
      }

      const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

      const formData = new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        refresh_token: refreshToken,
        grant_type: "refresh_token",
        scope:
          "offline_access Tasks.Read Tasks.ReadWrite Tasks.Read.Shared Tasks.ReadWrite.Shared User.Read",
      });

      const response = await fetch(tokenEndpoint, {
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
        body: formData,
      });

      if (!response.ok) {
        const errorText = await response.text();
        console.error(`Token refresh failed: ${errorText}`);

        // If refresh fails, prompt for re-authentication
        this.promptForReauth();
        return null;
      }

      const data = await response.json();

      const newTokens: StoredTokenData = {
        accessToken: data.access_token,
        refreshToken: data.refresh_token || refreshToken,
        expiresAt: Date.now() + data.expires_in * 1000 - 5 * 60 * 1000, // 5 min buffer
        clientId,
        clientSecret,
        tenantId,
      };

      // Save the refreshed tokens
      this.saveTokens(newTokens);

      // Also update Claude config if possible
      this.updateClaudeConfig(newTokens);

      return newTokens;
    } catch (error) {
      console.error("Error refreshing token:", error);
      this.promptForReauth();
      return null;
    }
  }

  saveTokens(tokens: StoredTokenData): void {
    this.currentTokens = tokens;
    writeTextFileSync(this.tokenFilePath, JSON.stringify(tokens, null, 2));
  }

  // Update Claude config automatically
  updateClaudeConfig(tokens: TokenData): void {
    try {
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
        : join(
          homedir(),
          ".config",
          "Claude",
          "claude_desktop_config.json",
        );

      if (!existsSync(claudeConfigPath)) {
        return;
      }

      const config = JSON.parse(readTextFileSync(claudeConfigPath));

      // Update the microsoft-todo server config
      if (config.mcpServers && config.mcpServers["microsoft-todo"]) {
        config.mcpServers["microsoft-todo"].env = {
          ...config.mcpServers["microsoft-todo"].env,
          MS_TODO_ACCESS_TOKEN: tokens.accessToken,
          MS_TODO_REFRESH_TOKEN: tokens.refreshToken,
        };

        // Write back the updated config
        writeTextFileSync(claudeConfigPath, JSON.stringify(config, null, 2));
        console.error("Updated Claude config with new tokens");
      }
    } catch (error) {
      console.error("Could not update Claude config:", error);
    }
  }

  promptForReauth(): void {
    console.error(`
=================================================================
TOKEN REFRESH FAILED - REAUTHENTICATION REQUIRED

Your Microsoft To Do tokens have expired and could not be refreshed.

To fix this:
1. Open a new terminal
2. Navigate to the microsoft-todo-mcp-server directory
3. Run: deno task auth
4. Complete the authentication in your browser
5. Restart Claude Desktop to use the new tokens

Your tokens are stored in: ${this.tokenFilePath}
=================================================================
    `);
  }

  // Store client credentials with tokens for future refreshes
  storeCredentials(
    clientId: string,
    clientSecret: string,
    tenantId: string,
  ): void {
    if (this.currentTokens) {
      this.currentTokens.clientId = clientId;
      this.currentTokens.clientSecret = clientSecret;
      this.currentTokens.tenantId = tenantId;
      this.saveTokens(this.currentTokens);
    }
  }
}

export const tokenManager = new TokenManager();
