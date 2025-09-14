// Authentication server for Microsoft Todo MCP service
import dotenv from "dotenv"
import express, { Request, Response } from "express"
import fs from "fs"
import { join, dirname } from "path"
import {
  ConfidentialClientApplication,
  AccountInfo,
  AuthenticationResult,
  Configuration,
  LogLevel,
} from "@azure/msal-node"
import { fileURLToPath } from "url"

// Initialize environment variables
dotenv.config()
console.log("Environment loaded")
console.log("CLIENT_ID:", process.env.CLIENT_ID ? "Present (hidden)" : "Missing")
console.log("CLIENT_SECRET:", process.env.CLIENT_SECRET ? "Present (hidden)" : "Missing")
console.log(
  "TENANT_ID:",
  process.env.TENANT_ID ? process.env.TENANT_ID : 'Not specified, using "organizations" (multi-tenant)',
)
console.log("REDIRECT_URI:", process.env.REDIRECT_URI || `http://localhost:3000/callback`)

// Get current file directory in ESM
const __filename = fileURLToPath(import.meta.url)
const __dirname = dirname(__filename)

const app = express()
const port = Number(process.env.PORT) || 3000
const TOKEN_FILE_PATH = join(process.cwd(), "tokens.json")

// Determine the tenant ID to use:
// - 'common' for both organization accounts and personal accounts
// - 'organizations' for organization accounts only (multi-tenant)
// - 'consumers' for personal accounts only
// - A specific tenant ID for a single organization
const tenantId = process.env.TENANT_ID || "organizations"

// Display authentication type
if (tenantId === "common") {
  console.log("Authentication type: Both organization and personal accounts (common)")
} else if (tenantId === "organizations") {
  console.log("Authentication type: Organizations only (multi-tenant)")
} else if (tenantId === "consumers") {
  console.log("Authentication type: Personal accounts only")
  console.log(
    "WARNING: Microsoft To Do API has limitations for personal accounts (MailboxNotEnabledForRESTAPI error may occur)",
  )
} else {
  console.log(`Authentication type: Single tenant (${tenantId})`)
}

// MSAL configuration for delegated permissions
const msalConfig: Configuration = {
  auth: {
    clientId: process.env.CLIENT_ID!,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    clientSecret: process.env.CLIENT_SECRET!,
  },
  system: {
    loggerOptions: {
      loggerCallback(loglevel: LogLevel, message: string, containsPii: boolean) {
        console.log(`MSAL Log: ${message}`)
      },
      piiLoggingEnabled: true,
      logLevel: LogLevel.Verbose,
    },
  },
  cache: {
    cachePlugin: {
      beforeCacheAccess: async (cacheContext) => {
        console.log("Cache access requested:", cacheContext)
      },
      afterCacheAccess: async (cacheContext) => {
        console.log("Cache access completed:", cacheContext)
      },
    },
  },
}

console.log("MSAL config created")

// Task-related permission scopes
const scopes = [
  "offline_access", // Put offline_access first to ensure it's not dropped
  "openid", // Add openid scope
  "profile", // Add profile scope
  "Tasks.Read",
  "Tasks.Read.Shared",
  "Tasks.ReadWrite",
  "Tasks.ReadWrite.Shared",
  "User.Read",
]

// Create MSAL application
const cca = new ConfidentialClientApplication(msalConfig)
console.log("MSAL application created")

// Setup a test route to check if server is working
app.get("/test", (req: Request, res: Response) => {
  res.send("Auth server is running correctly")
})

// Helper function to refresh an access token
async function refreshAccessToken(): Promise<{
  success: boolean
  error?: any
  response?: AuthenticationResult | null
  accessToken?: string
  expiresAt?: number
}> {
  try {
    // Get account info from the token cache
    const tokenCache = cca.getTokenCache()
    const accounts = await tokenCache.getAllAccounts()

    if (accounts.length === 0) {
      console.log("No accounts found in the token cache")
      return { success: false, error: "No accounts found in token cache" }
    }

    // Get the first account (we should have only one in this scenario)
    const account = accounts[0]
    console.log("Found account in token cache:", {
      username: account.username,
      localAccountId: account.localAccountId,
      tenantId: account.tenantId,
    })

    // Create a silent request using the account
    const silentRequest = {
      account: account,
      scopes: scopes,
      forceRefresh: true,
    }

    console.log("Attempting to acquire token silently...")
    const response = await cca.acquireTokenSilent(silentRequest)

    console.log("Token refreshed successfully")
    return {
      success: true,
      response: response,
      accessToken: response!.accessToken,
      expiresAt: Date.now() + ((response as any).expiresIn || 3600) * 1000 - 5 * 60 * 1000,
    }
  } catch (error) {
    console.error("Error refreshing token silently:", error)
    return {
      success: false,
      error: error,
    }
  }
}

// Update refresh endpoint to use acquireTokenSilent
app.get("/refresh", async (req: Request, res: Response) => {
  try {
    const result = await refreshAccessToken()

    if (result.success) {
      // Save updated token data
      const tokenData = {
        accessToken: result.accessToken,
        expiresAt: result.expiresAt,
        tokenType: result.response!.tokenType,
        scopes: result.response!.scopes,
      }

      // Save updated token data
      fs.writeFileSync(TOKEN_FILE_PATH, JSON.stringify(tokenData, null, 2), "utf8")

      res.json({
        success: true,
        message: "Token refreshed successfully",
        expiresAt: new Date(result.expiresAt!).toISOString(),
      })
    } else {
      // If silent refresh fails, redirect to login
      console.log("Silent token refresh failed, redirecting to login")
      res.json({
        success: false,
        message: "Token refresh failed, please login again",
        redirectUrl: "/",
      })
    }
  } catch (error: any) {
    console.error("Error in refresh route:", error)
    res.status(500).send(`Error refreshing token: ${error.message}`)
  }
})

// Add a client credentials flow endpoint
app.get("/silentLogin", async (req: Request, res: Response) => {
  try {
    console.log("Silent login endpoint accessed")

    // Client credentials flow requires different scopes format
    // We use resource/.default pattern here
    const clientCredentialRequest = {
      scopes: ["https://graph.microsoft.com/.default"],
      skipCache: true, // Force request to go to the server
    }

    console.log("Attempting client credentials flow with scopes:", clientCredentialRequest.scopes)

    const response = await cca.acquireTokenByClientCredential(clientCredentialRequest)

    if (!response) {
      throw new Error("No response from client credentials flow")
    }

    console.log("Client credentials response received", {
      hasAccessToken: !!response.accessToken,
      tokenType: response.tokenType,
      expiresOn: response.expiresOn,
      scopes: response.scopes,
    })

    // Get token cache after successful client credentials flow
    const tokenCache = cca.getTokenCache()
    const serializedCache = await tokenCache.serialize()
    const cacheJson = JSON.parse(serializedCache)

    console.log("Token cache after client credentials flow:", {
      hasRefreshTokens: !!cacheJson.RefreshTokens,
      hasRefreshToken: !!cacheJson.RefreshToken,
      cacheKeys: Object.keys(cacheJson),
    })

    // Check if we have any tokens that look like refresh tokens
    let refreshTokenFound = false
    for (const key in cacheJson) {
      if (key.toLowerCase().includes("refresh")) {
        refreshTokenFound = true
        console.log(`Found potential refresh token section: ${key}`)
      }
    }

    if (!refreshTokenFound) {
      console.log("No refresh token sections found in cache after client credentials flow")
    }

    // Client credentials flow won't typically have a refresh token
    // since it's an app-only flow with no user context
    res.json({
      success: true,
      message: "Client credentials flow completed",
      accessTokenPresent: !!response.accessToken,
      expiresOn: response.expiresOn,
    })
  } catch (error: any) {
    console.error("Error in silent login:", error)
    res.status(500).send(`Error in silent login: ${error.message}`)
  }
})

// Setup the auth flow
app.get("/", (req: Request, res: Response) => {
  console.log("Root route accessed, generating auth URL...")

  // Display information about potential issues for personal accounts
  let authInfo = `
    <html>
    <head>
      <title>Microsoft To Do MCP Authentication</title>
      <style>
        body { font-family: Arial, sans-serif; margin: 40px; line-height: 1.6; }
        .container { max-width: 800px; margin: 0 auto; }
        .warning { background-color: #fff3cd; border: 1px solid #ffeeba; padding: 15px; border-radius: 4px; margin-bottom: 20px; }
        .primary-button { background-color: #0078d4; color: white; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer; }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>Microsoft To Do MCP Authentication</h1>
  `

  // Add warning for personal accounts
  if (tenantId === "consumers" || tenantId === "common") {
    authInfo += `
        <div class="warning">
          <h3>⚠️ Important Note for Personal Microsoft Accounts</h3>
          <p>The Microsoft Graph API has limitations for personal Microsoft accounts (outlook.com, hotmail.com, live.com, etc.). 
          The To Do API is primarily designed for Microsoft 365 business accounts, not personal accounts.</p>
          <p>If you use a personal Microsoft account, you may encounter a <strong>"MailboxNotEnabledForRESTAPI"</strong> error. 
          This is a Microsoft service limitation, not an issue with this application's code or authentication setup.</p>
        </div>
    `
  }

  authInfo += `
        <p>Click the button below to authenticate with Microsoft and grant access to your To Do tasks.</p>
        <button class="primary-button" onclick="window.location.href='/auth'">Sign in with Microsoft</button>
      </div>
    </body>
    </html>
  `

  res.send(authInfo)
})

// Setup the actual auth route
app.get("/auth", (req: Request, res: Response) => {
  console.log("Auth route accessed, generating auth URL...")
  const authCodeUrlParameters = {
    scopes: scopes,
    redirectUri: process.env.REDIRECT_URI || `http://localhost:${port}/callback`,
    prompt: "consent", // Use only consent to force refresh token
    responseMode: "query" as any,
  }

  console.log("Auth parameters:", {
    scopes: scopes,
    redirectUri: process.env.REDIRECT_URI || `http://localhost:${port}/callback`,
    prompt: "consent",
    responseMode: "query",
  })

  cca
    .getAuthCodeUrl(authCodeUrlParameters)
    .then((response) => {
      console.log("Auth URL generated, redirecting to:", response.substring(0, 80) + "...")
      res.redirect(response)
    })
    .catch((error) => {
      console.error("Error getting auth code URL:", error)
      res.status(500).send(`Error generating authentication URL: ${JSON.stringify(error)}`)
    })
})

// Handle the callback from Microsoft login
app.get("/callback", async (req: Request, res: Response) => {
  console.log("Callback route accessed")
  console.log("Query parameters:", {
    code: req.query.code ? "Present (hidden)" : "Missing",
    state: req.query.state ? "Present" : "Missing",
    error: req.query.error || "None",
    error_description: req.query.error_description || "None",
  })

  const tokenRequest = {
    code: req.query.code as string,
    scopes: scopes,
    redirectUri: process.env.REDIRECT_URI || `http://localhost:${port}/callback`,
  }

  console.log("Token request parameters:", {
    scopes: scopes,
    redirectUri: process.env.REDIRECT_URI || `http://localhost:${port}/callback`,
  })

  try {
    const response = await cca.acquireTokenByCode(tokenRequest)

    // Log full response structure (without sensitive values)
    console.log("Token response structure:", {
      keys: Object.keys(response),
      hasAccessToken: !!response.accessToken,
      hasRefreshToken: !!(response as any).refreshToken,
      hasIdToken: !!response.idToken,
      tokenType: response.tokenType,
      expiresIn: (response as any).expiresIn,
      expiresOn: response.expiresOn,
      scopes: response.scopes,
      account: response.account
        ? {
            username: response.account.username,
            tenantId: response.account.tenantId,
            localAccountId: response.account.localAccountId,
          }
        : null,
    })

    // Get refresh token from token cache
    const tokenCache = cca.getTokenCache()
    const serializedCache = await tokenCache.serialize()
    const cacheJson = JSON.parse(serializedCache)

    // Log the full cache structure for debugging (excluding sensitive values)
    console.log("Full token cache structure keys:", Object.keys(cacheJson))
    if (cacheJson.RefreshToken) {
      console.log("RefreshToken keys in cache:", Object.keys(cacheJson.RefreshToken))
    } else if (cacheJson.RefreshTokens) {
      console.log("RefreshTokens keys in cache:", Object.keys(cacheJson.RefreshTokens))
    }

    // Try different ways to get the refresh token
    let refreshToken: string | null = null

    // Method 1: Check RefreshTokens (plural)
    if (cacheJson.RefreshTokens && Object.keys(cacheJson.RefreshTokens).length > 0) {
      const refreshTokenKeys = Object.keys(cacheJson.RefreshTokens)
      refreshToken = cacheJson.RefreshTokens[refreshTokenKeys[0]].secret
      console.log("Refresh token found using RefreshTokens collection")
    }
    // Method 2: Check RefreshToken (singular)
    else if (cacheJson.RefreshToken && Object.keys(cacheJson.RefreshToken).length > 0) {
      const refreshTokenKeys = Object.keys(cacheJson.RefreshToken)
      refreshToken = cacheJson.RefreshToken[refreshTokenKeys[0]].secret
      console.log("Refresh token found using RefreshToken collection")
    }
    // Method 3: Look for any key with "refresh" in it
    else {
      for (const cacheSection in cacheJson) {
        if (cacheSection.toLowerCase().includes("refresh") && typeof cacheJson[cacheSection] === "object") {
          for (const key in cacheJson[cacheSection]) {
            if (cacheJson[cacheSection][key] && cacheJson[cacheSection][key].secret) {
              refreshToken = cacheJson[cacheSection][key].secret
              console.log(`Refresh token found in ${cacheSection}.${key}`)
              break
            }
          }
          if (refreshToken) break
        }
      }
    }

    if (!refreshToken) {
      console.log("Could not find refresh token in token cache")
    }

    // Calculate token expiration (make sure it's never null)
    const expiresInSeconds = (response as any).expiresIn || 3600
    const expiresAt = Date.now() + expiresInSeconds * 1000 - 5 * 60 * 1000

    console.log("Token expiration details:", {
      expiresInSeconds,
      expiresAt: new Date(expiresAt).toLocaleString(),
      currentTime: new Date().toLocaleString(),
    })

    // Store tokens with client credentials for future refreshes
    const tokenData = {
      accessToken: response.accessToken,
      refreshToken: refreshToken || "",
      expiresAt: expiresAt,
      tokenType: response.tokenType,
      scopes: response.scopes,
      // Add client credentials for automatic refresh
      clientId: process.env.CLIENT_ID,
      clientSecret: process.env.CLIENT_SECRET,
      tenantId: tenantId,
    }

    fs.writeFileSync(TOKEN_FILE_PATH, JSON.stringify(tokenData, null, 2), "utf8")

    console.log("Authentication successful! Token saved to:", TOKEN_FILE_PATH)
    console.log("Refresh token obtained:", refreshToken ? "Yes" : "No")

    // Format token display with safety checks
    const accessTokenDisplay = response.accessToken
      ? `${response.accessToken.substring(0, 15)}...${response.accessToken.substring(response.accessToken.length - 5)}`
      : "Not provided"

    const refreshTokenDisplay = refreshToken
      ? `${refreshToken.substring(0, 10)}...${refreshToken.substring(refreshToken.length - 5)}`
      : "Not provided"

    // Check if the account is a personal account
    const isPersonalAccount =
      response.account &&
      (response.account.username.includes("@outlook.com") ||
        response.account.username.includes("@hotmail.com") ||
        response.account.username.includes("@live.com") ||
        response.account.username.includes("@msn.com"))

    let warningMessage = ""

    if (isPersonalAccount) {
      warningMessage = `
        <div class="warning">
          <h3>⚠️ Important Note for Personal Microsoft Accounts</h3>
          <p>You are signed in with a personal Microsoft account (${response.account!.username}).</p>
          <p>The Microsoft Graph API has limitations for personal Microsoft accounts. The To Do API is primarily designed for Microsoft 365 business accounts, not personal accounts.</p>
          <p>You may encounter a <strong>"MailboxNotEnabledForRESTAPI"</strong> error when trying to access To Do tasks. This is a Microsoft service limitation, not an issue with this application's code or authentication setup.</p>
        </div>
      `
    }

    res.send(`
      <html>
      <head>
        <title>Authentication Successful</title>
        <style>
          body { font-family: Arial, sans-serif; margin: 40px; line-height: 1.6; }
          .container { max-width: 800px; margin: 0 auto; }
          .success { background-color: #d4edda; border: 1px solid #c3e6cb; padding: 15px; border-radius: 4px; margin-bottom: 20px; }
          .warning { background-color: #fff3cd; border: 1px solid #ffeeba; padding: 15px; border-radius: 4px; margin-bottom: 20px; }
          .token-details { background-color: #f8f9fa; padding: 15px; border-radius: 4px; margin-top: 20px; }
          .debug-info { margin-top: 30px; border-top: 1px solid #dee2e6; padding-top: 20px; }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="success">
            <h1>Authentication Successful!</h1>
            <p>You can now close this window and use the Microsoft Todo MCP service.</p>
          </div>
          
          ${warningMessage}
          
          <div class="token-details">
            <h3>Token Details:</h3>
            <ul>
              <li>Access Token: ${accessTokenDisplay}</li>
              <li>Refresh Token: ${refreshTokenDisplay}</li>
              <li>Token Type: ${response.tokenType || "Not provided"}</li>
              <li>Scopes: ${response.scopes ? response.scopes.join(", ") : "Not provided"}</li>
              <li>Expires: ${new Date(expiresAt).toLocaleString()}</li>
            </ul>
          </div>
          
          <div class="debug-info">
            <h3>Debug Information:</h3>
            <pre>${JSON.stringify(
              {
                hasRefreshToken: !!refreshToken,
                tokenType: response.tokenType,
                scopes: response.scopes,
                cacheHasRefreshTokens: cacheJson.RefreshTokens && Object.keys(cacheJson.RefreshTokens).length > 0,
              },
              null,
              2,
            )}</pre>
          </div>
        </div>
      </body>
      </html>
    `)
  } catch (error: any) {
    console.error("Token acquisition error:", {
      errorCode: error.errorCode,
      errorMessage: error.errorMessage,
      subError: error.subError,
      correlationId: error.correlationId,
      stack: error.stack,
    })
    res.status(500).send(`Error acquiring token: ${JSON.stringify(error)}`)
  }
})

// Start the server
app.listen(port, () => {
  console.log(`Auth server running at http://localhost:${port}`)
  console.log("Open your browser and navigate to the URL above to authenticate.")
  console.log(`Or try http://localhost:${port}/test to verify the server is running.`)
})
