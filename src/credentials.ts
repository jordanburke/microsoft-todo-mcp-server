// Where a tool call gets the Microsoft token it sends to Graph.
//
// The tools never reach for a token themselves: the server is built with a source, so the
// local (stdio) server uses the token manager while the HTTP server uses each caller's own
// signed-in session and can never fall back to the host's tokens.
import type { Option } from "functype"
import type { Context, SessionAuth } from "somamcp"

/** The token one tool call sends to Graph, and how to replace it after Graph rejects it. */
export type GraphAuth = {
  readonly accessToken: string
  /** A replacement for `rejected`, or None to let Graph's 401 stand. */
  readonly renew: (rejected: string) => Promise<Option<string>>
}

/** Where a tool call gets its Graph token, and what `auth-status` reports. */
export type CredentialSource<S extends SessionAuth = SessionAuth> = {
  readonly graphAuth: (context: Context<S>) => Promise<Option<GraphAuth>>
  readonly status: (context: Context<S>) => Promise<string>
}
