// src/token-manager.ts
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "fs"
import { homedir } from "os"
import { dirname, join, resolve } from "path"

interface TokenData {
  accessToken: string
  refreshToken: string
  expiresAt: number
}

interface StoredTokenData extends TokenData {
  clientId?: string
  clientSecret?: string
  tenantId?: string
}

// A Graph access token is a JWT, so its real expiry is in the `exp` claim. Reading it
// beats guessing: the environment pair carries no expiry of its own, and assuming one
// is what used to keep the refresh branch permanently out of reach.
const expiryFromAccessToken = (accessToken: string): number => {
  const FALLBACK_LIFETIME_MS = 3600 * 1000
  const REFRESH_BUFFER_MS = 5 * 60 * 1000

  try {
    const payload = accessToken.split(".")[1]
    if (!payload) return Date.now() + FALLBACK_LIFETIME_MS

    const { exp } = JSON.parse(Buffer.from(payload, "base64url").toString("utf8"))
    if (typeof exp !== "number") return Date.now() + FALLBACK_LIFETIME_MS

    return exp * 1000 - REFRESH_BUFFER_MS
  } catch {
    // An opaque token is not an error — fall back to the old assumption, but stamp it
    // once (see readEnvTokens) rather than on every call.
    return Date.now() + FALLBACK_LIFETIME_MS
  }
}

export class TokenManager {
  private tokenFilePath: string
  private currentTokens: StoredTokenData | null = null
  private envTokens: TokenData | null = null

  constructor() {
    this.tokenFilePath = TokenManager.resolveTokenFilePath()
    mkdirSync(dirname(this.tokenFilePath), { recursive: true })
    console.error(`Token file path: ${this.tokenFilePath}`)
  }

  private static resolveTokenFilePath(): string {
    // MSTODO_TOKEN_FILE is documented in the README as the way to relocate the token
    // store. It used to be read in cli.ts and handed to startServer, which discarded
    // it, so setting it moved nothing and every tool call failed to authenticate.
    const override = process.env.MSTODO_TOKEN_FILE
    if (override) return resolve(override)

    // Otherwise a consistent location across platforms
    const configDir =
      process.platform === "win32"
        ? join(process.env.APPDATA || join(homedir(), "AppData", "Roaming"), "microsoft-todo-mcp")
        : join(homedir(), ".config", "microsoft-todo-mcp")

    return join(configDir, "tokens.json")
  }

  // The environment pair is a bootstrap credential, not a live one: nothing rewrites
  // those variables when a token is renewed. Stamp its deadline once, when the value
  // is first seen, so the deadline stops moving with the clock.
  private readEnvTokens(): TokenData | null {
    const accessToken = process.env.MS_TODO_ACCESS_TOKEN
    const refreshToken = process.env.MS_TODO_REFRESH_TOKEN
    if (!accessToken || !refreshToken) return null

    if (!this.envTokens || this.envTokens.accessToken !== accessToken) {
      this.envTokens = { accessToken, refreshToken, expiresAt: expiryFromAccessToken(accessToken) }
    }

    return this.envTokens
  }

  // Try to get tokens from multiple sources
  async getTokens(): Promise<TokenData | null> {
    // 0. A token renewed in this process outranks everything below it. The environment
    // pair and the file on disk are both starting points, and the environment pair in
    // particular is never rewritten, so without this the stale value wins every time.
    if (this.currentTokens && Date.now() < this.currentTokens.expiresAt) {
      return this.currentTokens
    }

    // 1. Check environment variables first (for backward compatibility)
    const envTokens = this.readEnvTokens()
    if (envTokens) {
      if (Date.now() < envTokens.expiresAt) {
        return envTokens
      }

      const refreshed = await this.refreshToken(envTokens.refreshToken)
      if (refreshed) {
        return refreshed
      }

      // Hand back the expired token rather than nothing: the caller's 401 says more
      // about what went wrong than a bare "not authenticated".
      return envTokens
    }

    // 2. Check stored token file
    if (existsSync(this.tokenFilePath)) {
      try {
        const data = readFileSync(this.tokenFilePath, "utf8")
        this.currentTokens = JSON.parse(data)

        if (this.currentTokens) {
          // Check if expired
          if (Date.now() > this.currentTokens.expiresAt) {
            // Try to refresh
            const refreshed = await this.refreshToken(this.currentTokens.refreshToken)
            if (refreshed) {
              return refreshed
            }
          }
          return this.currentTokens
        }
      } catch (error) {
        console.error("Error reading token file:", error)
      }
    }

    // 3. Check legacy token file location
    const legacyPath = join(process.cwd(), "tokens.json")
    if (existsSync(legacyPath)) {
      try {
        const data = readFileSync(legacyPath, "utf8")
        const tokens = JSON.parse(data)

        // Migrate to new location
        this.saveTokens(tokens)

        return tokens
      } catch (error) {
        console.error("Error reading legacy token file:", error)
      }
    }

    return null
  }

  async refreshToken(refreshToken: string): Promise<TokenData | null> {
    try {
      // Get client credentials from stored tokens or environment
      const clientId = this.currentTokens?.clientId || process.env.CLIENT_ID
      const clientSecret = this.currentTokens?.clientSecret || process.env.CLIENT_SECRET
      const tenantId = this.currentTokens?.tenantId || process.env.TENANT_ID || "organizations"

      if (!clientId || !clientSecret) {
        console.error("Missing client credentials for token refresh")
        return null
      }

      const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`

      const formData = new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        refresh_token: refreshToken,
        grant_type: "refresh_token",
        scope: "offline_access Tasks.Read Tasks.ReadWrite Tasks.Read.Shared Tasks.ReadWrite.Shared User.Read",
      })

      const response = await fetch(tokenEndpoint, {
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
        body: formData,
      })

      if (!response.ok) {
        const errorText = await response.text()
        console.error(`Token refresh failed: ${errorText}`)

        // If refresh fails, prompt for re-authentication
        this.promptForReauth()
        return null
      }

      const data = await response.json()

      const newTokens: StoredTokenData = {
        accessToken: data.access_token,
        refreshToken: data.refresh_token || refreshToken,
        expiresAt: Date.now() + data.expires_in * 1000 - 5 * 60 * 1000, // 5 min buffer
        clientId,
        clientSecret,
        tenantId,
      }

      // Save the refreshed tokens
      this.saveTokens(newTokens)

      // Also update Claude config if possible
      await this.updateClaudeConfig(newTokens)

      return newTokens
    } catch (error) {
      console.error("Error refreshing token:", error)
      this.promptForReauth()
      return null
    }
  }

  saveTokens(tokens: StoredTokenData): void {
    this.currentTokens = tokens
    writeFileSync(this.tokenFilePath, JSON.stringify(tokens, null, 2), "utf8")
  }

  // Update Claude config automatically
  async updateClaudeConfig(tokens: TokenData): Promise<void> {
    try {
      const claudeConfigPath =
        process.platform === "win32"
          ? join(process.env.APPDATA || "", "Claude", "claude_desktop_config.json")
          : process.platform === "darwin"
            ? join(homedir(), "Library", "Application Support", "Claude", "claude_desktop_config.json")
            : join(homedir(), ".config", "Claude", "claude_desktop_config.json")

      if (!existsSync(claudeConfigPath)) {
        return
      }

      const config = JSON.parse(readFileSync(claudeConfigPath, "utf8"))

      // Update the microsoft-todo server config
      if (config.mcpServers && config.mcpServers["microsoft-todo"]) {
        config.mcpServers["microsoft-todo"].env = {
          ...config.mcpServers["microsoft-todo"].env,
          MS_TODO_ACCESS_TOKEN: tokens.accessToken,
          MS_TODO_REFRESH_TOKEN: tokens.refreshToken,
        }

        // Write back the updated config
        writeFileSync(claudeConfigPath, JSON.stringify(config, null, 2), "utf8")
        console.error("Updated Claude config with new tokens")
      }
    } catch (error) {
      console.error("Could not update Claude config:", error)
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
3. Run: pnpm run auth
4. Complete the authentication in your browser
5. Restart Claude Desktop to use the new tokens

Your tokens are stored in: ${this.tokenFilePath}
=================================================================
    `)
  }

  // Store client credentials with tokens for future refreshes
  async storeCredentials(clientId: string, clientSecret: string, tenantId: string): Promise<void> {
    if (this.currentTokens) {
      this.currentTokens.clientId = clientId
      this.currentTokens.clientSecret = clientSecret
      this.currentTokens.tenantId = tenantId
      this.saveTokens(this.currentTokens)
    }
  }
}

export const tokenManager = new TokenManager()
