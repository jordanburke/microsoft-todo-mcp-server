// Checklist-item tools: get, create, update, delete.
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it } from "vitest"

import {
  GRAPH,
  graphError,
  type Harness,
  json,
  noContent,
  resetNetwork,
  restoreNetwork,
  startServer,
} from "./test-support/graph-harness.js"

let h: Harness

beforeAll(async () => {
  h = await startServer()
})
afterAll(() => h.cleanup())
beforeEach(() => resetNetwork(h.net))
afterEach(() => restoreNetwork())

const TASK = "/me/todo/lists/L1/tasks/T1"
const ITEMS = `${TASK}/checklistItems`

describe("get-checklist-items", () => {
  it("reads the task's title, then every page of its items", async () => {
    h.net
      .on("GET", TASK, json({ id: "T1", title: "Move house" }))
      .on(
        "GET",
        ITEMS,
        json({
          value: [{ id: "C1", displayName: "Book van", isChecked: true, createdDateTime: "2026-11-01T14:30:00Z" }],
          "@odata.nextLink": `${GRAPH}${ITEMS}?$skip=1`,
        }),
      )
      .on("GET", `${ITEMS}?$skip=1`, json({ value: [{ id: "C2", displayName: "Pack books", isChecked: false }] }))

    const { text } = await h.callTool("get-checklist-items", { listId: "L1", taskId: "T1" })

    expect(h.net.graphRequests()).toEqual([`GET ${TASK}`, `GET ${ITEMS}`, `GET ${ITEMS}?$skip=1`])
    expect(text).toBe(
      'Checklist items for task "Move house" (ID: T1):\n\n' +
        "✓ Book van (ID: C1)\nCreated: 11/1/2026, 2:30:00 PM\n\n" +
        "○ Pack books (ID: C2)",
    )
  })

  it('calls the task "Unknown Task" when its title cannot be read', async () => {
    h.net.on("GET", TASK, graphError(404, "itemNotFound", "gone")).on("GET", ITEMS, json({ value: [] }))

    expect((await h.callTool("get-checklist-items", { listId: "L1", taskId: "T1" })).text).toBe(
      'No checklist items found for task "Unknown Task" (ID: T1)',
    )
  })

  it("reports a failure to read the items", async () => {
    h.net.on("GET", TASK, json({ id: "T1", title: "x" })).on("GET", ITEMS, graphError(500, "generalException", "boom"))

    expect((await h.callTool("get-checklist-items", { listId: "L1", taskId: "T1" })).text).toBe(
      "Failed to retrieve checklist items for task: T1",
    )
  })
})

describe("create-checklist-item", () => {
  it("POSTs the item, with isChecked only when given", async () => {
    h.net.on("POST", ITEMS, json({ id: "C9", displayName: "Label boxes", isChecked: false }, 201))

    const plain = await h.callTool("create-checklist-item", { listId: "L1", taskId: "T1", displayName: "Label boxes" })
    await h.callTool("create-checklist-item", { listId: "L1", taskId: "T1", displayName: "Done one", isChecked: true })

    expect(h.net.calls.map((c) => c.body)).toEqual([
      { displayName: "Label boxes" },
      { displayName: "Done one", isChecked: true },
    ])
    expect(plain.text).toBe("Checklist item created successfully!\nContent: Label boxes\nID: C9")
  })

  it("reports a failure", async () => {
    h.net.on("POST", ITEMS, graphError(400, "invalidRequest", "bad"))

    expect((await h.callTool("create-checklist-item", { listId: "L1", taskId: "T1", displayName: "x" })).text).toBe(
      "Failed to create checklist item for task: T1",
    )
  })
})

describe("update-checklist-item", () => {
  it("PATCHes the given fields and prints the new state", async () => {
    h.net.on("PATCH", `${ITEMS}/C1`, json({ id: "C1", displayName: "Book big van", isChecked: true }))

    const { text } = await h.callTool("update-checklist-item", {
      listId: "L1",
      taskId: "T1",
      checklistItemId: "C1",
      displayName: "Book big van",
      isChecked: true,
    })

    expect(h.net.calls[0].body).toEqual({ displayName: "Book big van", isChecked: true })
    expect(text).toBe("Checklist item updated successfully!\nContent: Book big van\nStatus: Checked")
  })

  it('prints "Not checked" for an unchecked item', async () => {
    h.net.on("PATCH", `${ITEMS}/C1`, json({ id: "C1", displayName: "x", isChecked: false }))

    expect(
      (
        await h.callTool("update-checklist-item", {
          listId: "L1",
          taskId: "T1",
          checklistItemId: "C1",
          isChecked: false,
        })
      ).text,
    ).toContain("Status: Not checked")
  })

  it("refuses an update with nothing to change, without calling Graph", async () => {
    const { text } = await h.callTool("update-checklist-item", { listId: "L1", taskId: "T1", checklistItemId: "C1" })

    expect(text).toBe("No properties provided for update. Please specify either displayName or isChecked.")
    expect(h.net.calls).toHaveLength(0)
  })

  it("reports a failure", async () => {
    h.net.on("PATCH", `${ITEMS}/C1`, graphError(404, "itemNotFound", "gone"))

    expect(
      (
        await h.callTool("update-checklist-item", {
          listId: "L1",
          taskId: "T1",
          checklistItemId: "C1",
          isChecked: true,
        })
      ).text,
    ).toBe("Failed to update checklist item with ID: C1")
  })
})

describe("delete-checklist-item", () => {
  it("DELETEs the item and reports success on 204", async () => {
    h.net.on("DELETE", `${ITEMS}/C1`, noContent())

    const { text } = await h.callTool("delete-checklist-item", { listId: "L1", taskId: "T1", checklistItemId: "C1" })

    expect(h.net.graphRequests()).toEqual([`DELETE ${ITEMS}/C1`])
    expect(text).toBe("Checklist item with ID: C1 was successfully deleted from task: T1")
  })

  it("KNOWN ISSUE: reports success even when Graph refuses the delete", async () => {
    h.net.on("DELETE", `${ITEMS}/C1`, graphError(404, "itemNotFound", "gone"))

    expect(
      (await h.callTool("delete-checklist-item", { listId: "L1", taskId: "T1", checklistItemId: "C1" })).text,
    ).toBe("Checklist item with ID: C1 was successfully deleted from task: T1")
  })
})
