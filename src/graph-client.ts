// Microsoft Graph client, built on functype's Http/IO.
//
// Every request is a lazy IO with a typed error channel (GraphError). Tools run it at the
// MCP boundary with `.run()`, fold the Either, and render a failure with `graphErrorText`.
import { Http, HttpError, type HttpMethod, IO } from "functype"

import type { GraphAuth } from "./credentials.js"

export const MS_GRAPH_BASE = "https://graph.microsoft.com/v1.0"
const USER_AGENT = "microsoft-todo-mcp-server/1.0"
const MAX_PAGES = 100

/** Graph refused the mailbox: Microsoft has not enabled the To Do API for it. */
export type MailboxNotEnabled = { readonly _tag: "MailboxNotEnabled"; readonly status: number; readonly body: string }
/** A paged collection did not end within MAX_PAGES pages. */
export type PageLimit = { readonly _tag: "PageLimit"; readonly pages: number }

export type GraphError = HttpError | MailboxNotEnabled | PageLimit

type Page<T> = { readonly value?: readonly T[]; readonly "@odata.nextLink"?: string }

const http = Http.client({
  baseUrl: MS_GRAPH_BASE,
  defaultHeaders: {
    "User-Agent": USER_AGENT,
    Accept: "application/json",
    "Content-Type": "application/json",
  },
})

const MAILBOX_BANNER = `
=================================================================
ERROR: MailboxNotEnabledForRESTAPI

Microsoft has not enabled the To Do API for this account's mailbox. It is a
Microsoft-side setting, not a problem with your sign-in or this server. To Do
itself still works on the web and in the apps.

Personal accounts (outlook.com, live.com and so on) normally do have API
access; this error means this particular mailbox does not.
=================================================================
`

// One request with one token. DELETE answers 204 with no body, so it is read raw and
// yields undefined; every other call must return JSON, and an empty or non-JSON 2xx body
// is a DecodeError.
const send = <T>(method: HttpMethod, url: string, token: string, body?: unknown): IO<never, HttpError, T> =>
  IO.sync(() => console.error(`Making request to: ${method} ${url}`)).flatMap(() =>
    http
      .request<T>({
        url,
        method,
        headers: { Authorization: `Bearer ${token}` },
        ...(body !== undefined && (method === "POST" || method === "PATCH") ? { body } : {}),
        parseAs: method === "DELETE" ? "raw" : "json",
      })
      .map((response) => (method === "DELETE" ? (undefined as T) : response.data)),
  )

// A 401 proves the token is dead whatever its recorded expiry says: ask the call's
// credentials for a replacement and retry once. No replacement, or the same token back,
// leaves the 401 standing.
const renewAndRetry = <T>(
  error: HttpError,
  method: HttpMethod,
  url: string,
  auth: GraphAuth,
  body?: unknown,
): IO<never, HttpError, T> => {
  if (error._tag !== "HttpStatusError" || error.status !== 401) return IO.fail(error)
  console.error("Got 401, attempting token refresh...")
  return IO.tryPromise({
    try: () => auth.renew(auth.accessToken),
    catch: (cause) => HttpError.networkError(url, method, cause),
  }).flatMap((renewed) =>
    renewed
      .filter((token) => token !== auth.accessToken)
      .fold(
        () => IO.fail<HttpError>(error),
        (token) => send<T>(method, url, token, body),
      ),
  )
}

const classify = (error: HttpError): GraphError => {
  if (error._tag === "HttpStatusError" && error.body.includes("MailboxNotEnabledForRESTAPI")) {
    console.error(MAILBOX_BANNER)
    return { _tag: "MailboxNotEnabled", status: error.status, body: error.body }
  }
  return error
}

/** One Graph request, renewing the token and retrying once on 401. */
export const graphRequest = <T>(
  url: string,
  auth: GraphAuth,
  method: HttpMethod = "GET",
  body?: unknown,
): IO<never, GraphError, T> =>
  send<T>(method, url, auth.accessToken, body)
    .catchAll((error) => renewAndRetry<T>(error, method, url, auth, body))
    .mapError((error) => {
      const classified = classify(error)
      console.error(`Graph request failed: ${method} ${url}: ${graphErrorText(classified)}`)
      return classified
    })
    // Logs the size, not the content: over HTTP the content is a signed-in user's tasks.
    .tap((data) => console.error(`Response received: ${JSON.stringify(data ?? null).length} bytes`))

/**
 * Every item of a paged collection. Graph pages on its own terms and signals more with
 * @odata.nextLink, which a large $top does not suppress. Follows the links to the end, or
 * until `limit` items are collected. A failed page fails the whole call, so a caller never
 * gets a silently truncated list.
 */
export const graphRequestAll = <T>(url: string, auth: GraphAuth, limit?: number): IO<never, GraphError, T[]> => {
  const step = (next: string, items: readonly T[], pages: number): IO<never, GraphError, T[]> =>
    graphRequest<Page<T>>(next, auth).flatMap((page) => {
      const collected = [...items, ...(page.value ?? [])]
      if (limit !== undefined && collected.length >= limit) return IO.succeed(collected.slice(0, limit))
      const link = page["@odata.nextLink"]
      if (!link) return IO.succeed(collected)
      if (pages + 1 >= MAX_PAGES) return IO.fail<GraphError>({ _tag: "PageLimit", pages: pages + 1 })
      return step(link, collected, pages + 1)
    })
  return step(url, [], 0)
}

/** The reason a request failed, as tools show it to the caller. */
export const graphErrorText = (error: unknown): string => {
  const e = error as GraphError
  switch (e?._tag) {
    case "HttpStatusError":
      return `HTTP error! status: ${e.status}, body: ${e.body}`
    case "MailboxNotEnabled":
      return "Microsoft has not enabled the To Do API for this account's mailbox (MailboxNotEnabledForRESTAPI)."
    case "PageLimit":
      return `stopped after ${MAX_PAGES} pages without reaching the end of the collection`
    case "NetworkError":
    case "DecodeError":
      return e.cause instanceof Error ? e.cause.message : String(e.cause)
    default:
      // A defect (a throw inside a map, say) reaches the Left untagged.
      return error instanceof Error ? error.message : String(error)
  }
}

/**
 * Run a request at the tool boundary: the data, or null with the failure logged. For
 * tools whose output does not say why a request failed.
 */
export const runOrNull = async <T>(io: IO<never, GraphError, T>): Promise<T | null> =>
  (await io.run()).fold(
    (error) => {
      console.error("Error making Graph API request:", graphErrorText(error))
      return null
    },
    (data) => data,
  )
