import { describe, expect, it } from "vitest"

import { normalizeTaskSelect } from "./todo-index.js"

describe("get-tasks query parameters", () => {
  it("removes fields rejected by Microsoft Graph while preserving supported selections", () => {
    expect(normalizeTaskSelect("id,title,status,bodyLastModifiedDateTime")).toBe("id,status")
  })

  it("omits select when it only contains fields returned by default or unsupported by the endpoint", () => {
    expect(normalizeTaskSelect("title, bodyLastModifiedDateTime")).toBeUndefined()
  })

  it("trims and deduplicates fields", () => {
    expect(normalizeTaskSelect(" id, status, id ")).toBe("id,status")
  })
})
