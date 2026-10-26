import { afterEach, describe, expect, it, vi } from "vitest"
import { makeGraphRequestAll } from "../src/todo-index.js"

function jsonResponse(body: unknown): Response {
  return new Response(JSON.stringify(body), {
    status: 200,
    headers: { "Content-Type": "application/json" },
  })
}

afterEach(() => {
  vi.unstubAllGlobals()
})

describe("makeGraphRequestAll", () => {
  it("follows @odata.nextLink and combines all pages", async () => {
    const urls: string[] = []
    vi.stubGlobal("fetch", async (url: string | URL | Request) => {
      urls.push(url.toString())
      if (urls.length === 1) {
        return jsonResponse({
          value: [{ id: "1" }],
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/page2",
        })
      }
      return jsonResponse({ value: [{ id: "2" }] })
    })

    const result = await makeGraphRequestAll("https://graph.microsoft.com/v1.0/page1", "token")

    expect(urls).toHaveLength(2)
    expect(result?.value).toEqual([{ id: "1" }, { id: "2" }])
    expect(result?.["@odata.nextLink"]).toBeUndefined()
  })

  it("returns null when the first request fails", async () => {
    vi.stubGlobal("fetch", async () => new Response("error", { status: 500 }))
    const result = await makeGraphRequestAll("https://graph.microsoft.com/v1.0/tasks", "token")
    expect(result).toBeNull()
  })

  it("stops after maxPages and keeps the pending nextLink", async () => {
    let calls = 0
    vi.stubGlobal(
      "fetch",
      vi.fn(async () => {
        calls++
        return jsonResponse({
          value: [{ id: String(calls) }],
          "@odata.nextLink": `https://graph.microsoft.com/v1.0/page${calls + 1}`,
        })
      }),
    )

    const result = await makeGraphRequestAll("https://graph.microsoft.com/v1.0/page1", "token", 3)

    expect(calls).toBe(3)
    expect(result?.value).toHaveLength(3)
    expect(result?.["@odata.nextLink"]).toBe("https://graph.microsoft.com/v1.0/page4")
  })

  it("handles single-page responses without nextLink", async () => {
    vi.stubGlobal("fetch", async () => jsonResponse({ value: [{ id: "only" }] }))
    const result = await makeGraphRequestAll("https://graph.microsoft.com/v1.0/lists", "token")
    expect(result?.value).toEqual([{ id: "only" }])
  })
})
