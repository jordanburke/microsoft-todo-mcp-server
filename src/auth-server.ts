// Authentication server for Microsoft Todo MCP service
import {
  ConfidentialClientApplication,
  type Configuration,
  type LogLevel,
} from "npm:@azure/msal-node@3.8.1";
import { join } from "@std/path";

// Load .env file manually (in case --env-file flag is not used)
try {
  const envContent = Deno.readTextFileSync(join(Deno.cwd(), ".env"));
  for (const line of envContent.split("\n")) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const eqIndex = trimmed.indexOf("=");
    if (eqIndex === -1) continue;
    const key = trimmed.slice(0, eqIndex).trim();
    const value = trimmed.slice(eqIndex + 1).trim();
    if (!Deno.env.has(key)) {
      Deno.env.set(key, value);
    }
  }
} catch {
  // .env file doesn't exist, that's OK - env vars may be set elsewhere
}

console.log("Environment loaded");
console.log(
  "CLIENT_ID:",
  Deno.env.get("CLIENT_ID") ? "Present (hidden)" : "Missing",
);
console.log(
  "CLIENT_SECRET:",
  Deno.env.get("CLIENT_SECRET") ? "Present (hidden)" : "Missing",
);
console.log(
  "TENANT_ID:",
  Deno.env.get("TENANT_ID") ||
    'Not specified, using "organizations" (multi-tenant)',
);
console.log(
  "REDIRECT_URI:",
  Deno.env.get("REDIRECT_URI") || `http://localhost:3000/callback`,
);

const port = 3000;
const TOKEN_FILE_PATH = join(Deno.cwd(), "tokens.json");

// Determine the tenant ID to use:
const tenantId = Deno.env.get("TENANT_ID") || "organizations";

// Display authentication type
if (tenantId === "common") {
  console.log(
    "Authentication type: Both organization and personal accounts (common)",
  );
} else if (tenantId === "organizations") {
  console.log("Authentication type: Organizations only (multi-tenant)");
} else if (tenantId === "consumers") {
  console.log("Authentication type: Personal accounts only");
  console.log(
    "WARNING: Microsoft To Do API has limitations for personal accounts (MailboxNotEnabledForRESTAPI error may occur)",
  );
} else {
  console.log(`Authentication type: Single tenant (${tenantId})`);
}

// MSAL configuration for delegated permissions
const msalConfig: Configuration = {
  auth: {
    clientId: Deno.env.get("CLIENT_ID")!,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    clientSecret: Deno.env.get("CLIENT_SECRET")!,
  },
  system: {
    loggerOptions: {
      loggerCallback(
        _loglevel: number,
        message: string,
        _containsPii: boolean,
      ) {
        console.log(`MSAL Log: ${message}`);
      },
      piiLoggingEnabled: true,
      logLevel: 0 as LogLevel, // Verbose
    },
  },
  cache: {
    cachePlugin: {
      beforeCacheAccess: async (cacheContext: unknown) => {
        await Promise.resolve();
        console.log("Cache access requested:", cacheContext);
      },
      afterCacheAccess: async (cacheContext: unknown) => {
        await Promise.resolve();
        console.log("Cache access completed:", cacheContext);
      },
    },
  },
};

console.log("MSAL config created");

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
];

// Create MSAL application
const cca = new ConfidentialClientApplication(msalConfig);
console.log("MSAL application created");

// Helper function to refresh an access token
async function refreshAccessToken(): Promise<{
  success: boolean;
  error?: unknown;
  response?: Awaited<ReturnType<typeof cca.acquireTokenSilent>> | null;
  accessToken?: string;
  expiresAt?: number;
}> {
  try {
    // Get account info from the token cache
    const tokenCache = cca.getTokenCache();
    const accounts = await tokenCache.getAllAccounts();

    if (accounts.length === 0) {
      console.log("No accounts found in the token cache");
      return { success: false, error: "No accounts found in token cache" };
    }

    // Get the first account (we should have only one in this scenario)
    const account = accounts[0];
    console.log("Found account in token cache:", {
      username: account.username,
      localAccountId: account.localAccountId,
      tenantId: account.tenantId,
    });

    // Create a silent request using the account
    const silentRequest = {
      account: account,
      scopes: scopes,
      forceRefresh: true,
    };

    console.log("Attempting to acquire token silently...");
    const response = await cca.acquireTokenSilent(silentRequest);

    console.log("Token refreshed successfully");
    return {
      success: true,
      response: response,
      accessToken: response!.accessToken,
      expiresAt: Date.now() +
        ((response as unknown as { expiresIn?: number }).expiresIn || 3600) *
          1000 -
        5 * 60 * 1000,
    };
  } catch (error) {
    console.error("Error refreshing token silently:", error);
    return {
      success: false,
      error: error,
    };
  }
}

// Helper to get query params from URL
function getQueryParams(url: string): URLSearchParams {
  return new URL(url).searchParams;
}

// HTML template helpers
function landingPage(): string {
  const accountTypeWarning = tenantId === "consumers" || tenantId === "common"
    ? `
        <div class="warning">
          <h3>⚠️ Important Note for Personal Microsoft Accounts</h3>
          <p>The Microsoft Graph API has limitations for personal Microsoft accounts (outlook.com, hotmail.com, live.com, etc.).
          The To Do API is primarily designed for Microsoft 365 business accounts, not personal accounts.</p>
          <p>If you use a personal Microsoft account, you may encounter a <strong>"MailboxNotEnabledForRESTAPI"</strong> error.
          This is a Microsoft service limitation, not an issue with this application's code or authentication setup.</p>
        </div>
    `
    : "";

  return `
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
        ${accountTypeWarning}
        <p>Click the button below to authenticate with Microsoft and grant access to your To Do tasks.</p>
        <button class="primary-button" onclick="window.location.href='/auth'">Sign in with Microsoft</button>
      </div>
    </body>
    </html>
  `;
}

function successPage(
  accessToken: string,
  refreshToken: string | null,
  tokenType: string,
  scopesList: string[],
  expiresAt: number,
  isPersonalAccount: boolean,
  cacheJson: Record<string, unknown>,
): string {
  const accessTokenDisplay = accessToken
    ? `${accessToken.substring(0, 15)}...${accessToken.substring(accessToken.length - 5)}`
    : "Not provided";

  const refreshTokenDisplay = refreshToken
    ? `${refreshToken.substring(0, 10)}...${refreshToken.substring(refreshToken.length - 5)}`
    : "Not provided";

  const warningMessage = isPersonalAccount
    ? `
        <div class="warning">
          <h3>⚠️ Important Note for Personal Microsoft Accounts</h3>
          <p>You are signed in with a personal Microsoft account.</p>
          <p>The Microsoft Graph API has limitations for personal Microsoft accounts. The To Do API is primarily designed for Microsoft 365 business accounts, not personal accounts.</p>
          <p>You may encounter a <strong>"MailboxNotEnabledForRESTAPI"</strong> error when trying to access To Do tasks. This is a Microsoft service limitation, not an issue with this application's code or authentication setup.</p>
        </div>
      `
    : "";

  return `
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
            <li>Token Type: ${tokenType || "Not provided"}</li>
            <li>Scopes: ${scopesList ? scopesList.join(", ") : "Not provided"}</li>
            <li>Expires: ${new Date(expiresAt).toLocaleString()}</li>
          </ul>
        </div>
        <div class="debug-info">
          <h3>Debug Information:</h3>
          <pre>${
    JSON.stringify(
      {
        hasRefreshToken: !!refreshToken,
        tokenType,
        scopes: scopesList,
        cacheHasRefreshTokens: cacheJson.RefreshTokens &&
          Object.keys(cacheJson.RefreshTokens as object).length > 0,
      },
      null,
      2,
    )
  }</pre>
        </div>
      </div>
    </body>
    </html>
  `;
}

// Main request handler
async function handleRequest(request: Request): Promise<Response> {
  const url = new URL(request.url);
  const path = url.pathname;

  try {
    switch (path) {
      case "/":
        console.log("Root route accessed");
        return new Response(landingPage(), {
          headers: { "Content-Type": "text/html" },
        });

      case "/test":
        return new Response("Auth server is running correctly");

      case "/auth": {
        console.log("Auth route accessed, generating auth URL...");
        const authCodeUrlParameters = {
          scopes: scopes,
          redirectUri: Deno.env.get("REDIRECT_URI") || `http://localhost:${port}/callback`,
          prompt: "consent",
          responseMode: "query" as const,
        };

        console.log("Auth parameters:", {
          scopes: scopes,
          redirectUri: Deno.env.get("REDIRECT_URI") || `http://localhost:${port}/callback`,
          prompt: "consent",
          responseMode: "query",
        });

        try {
          const response = await cca.getAuthCodeUrl(authCodeUrlParameters);
          console.log(
            "Auth URL generated, redirecting to:",
            response.substring(0, 80) + "...",
          );
          return new Response(null, {
            status: 302,
            headers: { Location: response },
          });
        } catch (error) {
          console.error("Error getting auth code URL:", error);
          return new Response(
            `Error generating authentication URL: ${JSON.stringify(error)}`,
            {
              status: 500,
            },
          );
        }
      }

      case "/callback": {
        const queryParams = getQueryParams(request.url);
        console.log("Callback route accessed");
        console.log("Query parameters:", {
          code: queryParams.get("code") ? "Present (hidden)" : "Missing",
          state: queryParams.get("state") ? "Present" : "Missing",
          error: queryParams.get("error") || "None",
          error_description: queryParams.get("error_description") || "None",
        });

        const code = queryParams.get("code");
        if (!code) {
          return new Response(
            `Authentication error: ${
              queryParams.get("error_description") ||
              "No authorization code received"
            }`,
            { status: 400 },
          );
        }

        const tokenRequest = {
          code,
          scopes: scopes,
          redirectUri: Deno.env.get("REDIRECT_URI") || `http://localhost:${port}/callback`,
        };

        console.log("Token request parameters:", {
          scopes: scopes,
          redirectUri: Deno.env.get("REDIRECT_URI") || `http://localhost:${port}/callback`,
        });

        try {
          const response = await cca.acquireTokenByCode(tokenRequest);

          // Log full response structure (without sensitive values)
          console.log("Token response structure:", {
            keys: Object.keys(response),
            hasAccessToken: !!response.accessToken,
            hasRefreshToken: !!(
              response as unknown as { refreshToken?: string }
            ).refreshToken,
            hasIdToken: !!response.idToken,
            tokenType: response.tokenType,
            expiresIn: (response as unknown as { expiresIn?: number })
              .expiresIn,
            expiresOn: response.expiresOn,
            scopes: response.scopes,
            account: response.account
              ? {
                username: response.account.username,
                tenantId: response.account.tenantId,
                localAccountId: response.account.localAccountId,
              }
              : null,
          });

          // Get refresh token from token cache
          const tokenCache = cca.getTokenCache();
          const serializedCache = await tokenCache.serialize();
          const cacheJson = JSON.parse(serializedCache);

          // Log the full cache structure for debugging
          console.log(
            "Full token cache structure keys:",
            Object.keys(cacheJson),
          );
          if (cacheJson.RefreshToken) {
            console.log(
              "RefreshToken keys in cache:",
              Object.keys(cacheJson.RefreshToken),
            );
          } else if (cacheJson.RefreshTokens) {
            console.log(
              "RefreshTokens keys in cache:",
              Object.keys(cacheJson.RefreshTokens),
            );
          }

          // Try different ways to get the refresh token
          let refreshToken: string | null = null;

          if (
            cacheJson.RefreshTokens &&
            Object.keys(cacheJson.RefreshTokens).length > 0
          ) {
            const refreshTokenKeys = Object.keys(cacheJson.RefreshTokens);
            refreshToken = (
              cacheJson.RefreshTokens as Record<string, { secret: string }>
            )[refreshTokenKeys[0]].secret;
            console.log("Refresh token found using RefreshTokens collection");
          } else if (
            cacheJson.RefreshToken &&
            Object.keys(cacheJson.RefreshToken).length > 0
          ) {
            const refreshTokenKeys = Object.keys(cacheJson.RefreshToken);
            refreshToken = (
              cacheJson.RefreshToken as Record<string, { secret: string }>
            )[refreshTokenKeys[0]].secret;
            console.log("Refresh token found using RefreshToken collection");
          } else {
            for (const cacheSection in cacheJson) {
              if (
                cacheSection.toLowerCase().includes("refresh") &&
                typeof cacheJson[cacheSection] === "object"
              ) {
                const section = cacheJson[cacheSection] as Record<
                  string,
                  { secret: string }
                >;
                for (const key in section) {
                  if (section[key]?.secret) {
                    refreshToken = section[key].secret;
                    console.log(
                      `Refresh token found in ${cacheSection}.${key}`,
                    );
                    break;
                  }
                }
                if (refreshToken) break;
              }
            }
          }

          if (!refreshToken) {
            console.log("Could not find refresh token in token cache");
          }

          // Calculate token expiration
          const expiresInSeconds = (response as unknown as { expiresIn?: number }).expiresIn ||
            3600;
          const expiresAt = Date.now() + expiresInSeconds * 1000 - 5 * 60 * 1000;

          console.log("Token expiration details:", {
            expiresInSeconds,
            expiresAt: new Date(expiresAt).toLocaleString(),
            currentTime: new Date().toLocaleString(),
          });

          // Store tokens with client credentials for future refreshes
          const tokenData = {
            accessToken: response.accessToken,
            refreshToken: refreshToken || "",
            expiresAt: expiresAt,
            tokenType: response.tokenType,
            scopes: response.scopes,
            clientId: Deno.env.get("CLIENT_ID"),
            clientSecret: Deno.env.get("CLIENT_SECRET"),
            tenantId: tenantId,
          };

          Deno.writeTextFileSync(
            TOKEN_FILE_PATH,
            JSON.stringify(tokenData, null, 2),
          );

          console.log(
            "Authentication successful! Token saved to:",
            TOKEN_FILE_PATH,
          );
          console.log("Refresh token obtained:", refreshToken ? "Yes" : "No");

          // Check if the account is a personal account
          const isPersonalAccount = response.account &&
            (response.account.username.includes("@outlook.com") ||
              response.account.username.includes("@hotmail.com") ||
              response.account.username.includes("@live.com") ||
              response.account.username.includes("@msn.com"));

          return new Response(
            successPage(
              response.accessToken || "",
              refreshToken,
              response.tokenType || "",
              response.scopes || [],
              expiresAt,
              !!isPersonalAccount,
              cacheJson as Record<string, unknown>,
            ),
            {
              headers: { "Content-Type": "text/html" },
            },
          );
        } catch (error: unknown) {
          const err = error as Record<string, unknown>;
          console.error("Token acquisition error:", {
            errorCode: err.errorCode,
            errorMessage: err.errorMessage,
            subError: err.subError,
            correlationId: err.correlationId,
          });
          return new Response(
            `Error acquiring token: ${JSON.stringify(error)}`,
            { status: 500 },
          );
        }
      }

      case "/refresh": {
        try {
          const result = await refreshAccessToken();

          if (result.success) {
            const tokenData = {
              accessToken: result.accessToken,
              expiresAt: result.expiresAt,
              tokenType: result.response!.tokenType,
              scopes: result.response!.scopes,
            };

            Deno.writeTextFileSync(
              TOKEN_FILE_PATH,
              JSON.stringify(tokenData, null, 2),
            );

            return Response.json({
              success: true,
              message: "Token refreshed successfully",
              expiresAt: new Date(result.expiresAt!).toISOString(),
            });
          } else {
            console.log("Silent token refresh failed, redirecting to login");
            return Response.json({
              success: false,
              message: "Token refresh failed, please login again",
              redirectUrl: "/",
            });
          }
        } catch (error: unknown) {
          const err = error as { message?: string };
          console.error("Error in refresh route:", error);
          return new Response(
            `Error refreshing token: ${err.message || String(error)}`,
            {
              status: 500,
            },
          );
        }
      }

      case "/silentLogin": {
        try {
          console.log("Silent login endpoint accessed");

          const clientCredentialRequest = {
            scopes: ["https://graph.microsoft.com/.default"],
            skipCache: true,
          };

          console.log(
            "Attempting client credentials flow with scopes:",
            clientCredentialRequest.scopes,
          );

          const response = await cca.acquireTokenByClientCredential(
            clientCredentialRequest,
          );

          if (!response) {
            throw new Error("No response from client credentials flow");
          }

          console.log("Client credentials response received", {
            hasAccessToken: !!response.accessToken,
            tokenType: response.tokenType,
            expiresOn: response.expiresOn,
            scopes: response.scopes,
          });

          const tokenCache = cca.getTokenCache();
          const serializedCache = await tokenCache.serialize();
          const cacheJson = JSON.parse(serializedCache);

          console.log("Token cache after client credentials flow:", {
            hasRefreshTokens: !!cacheJson.RefreshTokens,
            hasRefreshToken: !!cacheJson.RefreshToken,
            cacheKeys: Object.keys(cacheJson),
          });

          let refreshTokenFound = false;
          for (const key in cacheJson) {
            if (key.toLowerCase().includes("refresh")) {
              refreshTokenFound = true;
              console.log(`Found potential refresh token section: ${key}`);
            }
          }

          if (!refreshTokenFound) {
            console.log(
              "No refresh token sections found in cache after client credentials flow",
            );
          }

          return Response.json({
            success: true,
            message: "Client credentials flow completed",
            accessTokenPresent: !!response.accessToken,
            expiresOn: response.expiresOn,
          });
        } catch (error: unknown) {
          const err = error as { message?: string };
          console.error("Error in silent login:", error);
          return new Response(
            `Error in silent login: ${err.message || String(error)}`,
            {
              status: 500,
            },
          );
        }
      }

      default:
        return new Response("Not found", { status: 404 });
    }
  } catch (error: unknown) {
    console.error("Unhandled error:", error);
    return new Response(`Internal server error: ${String(error)}`, {
      status: 500,
    });
  }
}

// Start the server
Deno.serve({ port }, handleRequest);
console.log(`Auth server running at http://localhost:${port}`);
console.log("Open your browser and navigate to the URL above to authenticate.");
console.log(
  "Or try http://localhost:3000/test to verify the server is running.",
);
