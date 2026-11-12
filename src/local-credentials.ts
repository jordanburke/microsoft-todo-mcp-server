// The local server's credentials: the token manager's tokens (env pair, token file or legacy
// ./tokens.json), renewed by the token manager when Graph rejects one. Only the stdio entry
// point imports this, so the HTTP server can never reach the host's tokens.
import { IO, None, Option } from "functype"

import type { CredentialSource, GraphAuth } from "./credentials.js"
import { tokenManager } from "./token-manager.js"

// The access token, or None when there is none to be had (logged, not thrown).
const accessToken = async (): Promise<Option<string>> => {
  console.error("getAccessToken called")
  return (
    await IO.async(() => tokenManager.getTokens())
      .map((tokens) => Option(tokens).map((t) => t.accessToken))
      .run()
  ).fold(
    (error) => {
      console.error("Error getting access token:", error)
      return None<string>()
    },
    (token) => {
      console.error(token.isSome() ? "Successfully retrieved valid token" : "No valid tokens available")
      return token
    },
  )
}

const renew = async (rejected: string): Promise<Option<string>> =>
  Option((await tokenManager.refreshRejectedToken(rejected))?.accessToken)

export const localCredentials: CredentialSource = {
  graphAuth: async () => (await accessToken()).map((token): GraphAuth => ({ accessToken: token, renew })),

  status: async () => {
    const tokens = await tokenManager.getTokens()

    if (!tokens) {
      return "Not authenticated. Please run 'mstodo-setup' (or 'npx -p microsoft-todo-mcp-server mstodo-setup') to authenticate with Microsoft."
    }

    const isExpired = Date.now() > tokens.expiresAt
    const expiryTime = new Date(tokens.expiresAt).toLocaleString()

    return isExpired
      ? `Authentication expired at ${expiryTime}. Will attempt to refresh when you call any API.`
      : `Authenticated. Token expires at ${expiryTime}.`
  },
}
