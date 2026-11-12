// HTTP mode's credentials: the Microsoft token of the caller's own sign-in, from somamcp's
// Azure provider. There is no fallback. A call without a session gets no token, and a token
// Graph rejects is not renewed here: the client refreshes its sign-in, and the provider
// renews Microsoft's token with it. This module never imports the token manager, so an HTTP
// caller can never be served the host's tokens.
import { None, Option } from "functype"
import type { AzureSession } from "somamcp/node"

import type { CredentialSource, GraphAuth } from "./credentials.js"

const noRenewal = async (): Promise<Option<string>> => None<string>()

export const sessionCredentials: CredentialSource<AzureSession> = {
  graphAuth: async ({ session }) =>
    Option(session?.accessToken)
      .filter((token) => token !== "")
      .map((accessToken): GraphAuth => ({ accessToken, renew: noRenewal })),

  status: async ({ session }) => {
    if (!session?.accessToken) return "Not signed in. Connect to this server again to sign in with Microsoft."
    const who = session.upn ? ` as ${session.upn}` : ""
    if (session.expiresAt === undefined) return `Signed in${who}.`
    const expiry = new Date(session.expiresAt * 1000).toLocaleString()
    return Date.now() >= session.expiresAt * 1000
      ? `Signed in${who}. The Microsoft token expired at ${expiry}; your client renews it on its next refresh.`
      : `Signed in${who}. The Microsoft token expires at ${expiry}.`
  },
}
