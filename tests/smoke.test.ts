// Smoke test: spawn the built `dist/todo-index.js`, drive it over stdio with
// JSON-RPC, and confirm (1) the server starts, (2) the tool list exposes the
// new `recurrence` schema on create-task and update-task.
//
// This catches the recurring regression where tsdown emits a CJS-style
// `__require` shim in the published bundle — in that state the server
// crashes before any tool is listed.
//
// Skipped automatically if `dist/` is missing (run `pnpm run build` first).

import { spawn, type ChildProcessWithoutNullStreams } from "node:child_process"
import { existsSync } from "node:fs"
import { resolve } from "node:path"
import { describe, expect, it, beforeAll, afterAll } from "vitest"

const DIST = resolve(__dirname, "..", "dist", "todo-index.js")

interface JsonRpcResponse {
  id?: number
  result?: any
  error?: { code: number; message: string }
}

function send(proc: ChildProcessWithoutNullStreams, msg: any): Promise<JsonRpcResponse> {
  return new Promise((resolveP, rejectP) => {
    const id = msg.id
    const onData = (chunk: Buffer) => {
      buf += chunk.toString()
      let nl
      while ((nl = buf.indexOf("\n")) >= 0) {
        const line = buf.slice(0, nl).trim()
        buf = buf.slice(nl + 1)
        if (!line) continue
        let parsed: any
        try {
          parsed = JSON.parse(line)
        } catch {
          // ignore non-JSON lines (e.g. server stderr might bleed into stdout)
          continue
        }
        if (parsed.id === id) {
          proc.stdout.off("data", onData)
          resolveP(parsed)
          return
        }
      }
    }
    proc.stdout.on("data", onData)
    proc.stdin.write(JSON.stringify(msg) + "\n")
    setTimeout(() => rejectP(new Error("RPC timeout for id " + id)), 10000)
  })
}

let buf = ""
let proc: ChildProcessWithoutNullStreams | null = null

beforeAll(() => {
  if (!existsSync(DIST)) {
    throw new Error(`dist/todo-index.js not found at ${DIST} — run 'pnpm run build' first`)
  }
})

afterAll(async () => {
  if (proc && !proc.killed) {
    proc.kill()
    await new Promise((r) => setTimeout(r, 200))
  }
})

describe("built MCP server", () => {
  it("starts and lists tools including create-task recurrence", async () => {
    proc = spawn("node", [DIST], {
      stdio: ["pipe", "pipe", "pipe"],
      env: { ...process.env, REDIRECT_URI: "http://localhost:3000/callback" },
    }) as ChildProcessWithoutNullStreams
    // Suppress server stderr noise during the test
    proc.stderr.on("data", () => {})

    const init = await send(proc, {
      jsonrpc: "2.0",
      id: 1,
      method: "initialize",
      params: {
        protocolVersion: "2024-11-05",
        capabilities: {},
        clientInfo: { name: "smoke-test", version: "0.0.1" },
      },
    })
    expect(init.error, JSON.stringify(init.error)).toBeUndefined()
    expect(init.result.serverInfo.name).toBe("mstodo")

    // notifications/initialized — fire-and-forget
    proc.stdin.write(JSON.stringify({ jsonrpc: "2.0", method: "notifications/initialized" }) + "\n")

    const tools = await send(proc, { jsonrpc: "2.0", id: 2, method: "tools/list" })
    expect(tools.error).toBeUndefined()
    const names: string[] = tools.result.tools.map((t: any) => t.name)
    expect(names).toContain("create-task")
    expect(names).toContain("update-task")

    const createTask = tools.result.tools.find((t: any) => t.name === "create-task")
    const createProps = createTask.inputSchema.properties
    expect(createProps.recurrence).toBeDefined()
    // The nested pattern.type enum should include the standard MS Graph values
    const patternEnum: string[] = createProps.recurrence.properties.pattern.properties.type.enum
    expect(patternEnum).toEqual(
      expect.arrayContaining([
        "daily",
        "weekly",
        "absoluteMonthly",
        "relativeMonthly",
        "absoluteYearly",
        "relativeYearly",
      ]),
    )

    const updateTask = tools.result.tools.find((t: any) => t.name === "update-task")
    // For update-task, recurrence is anyOf [object, null] — both branches must
    // surface so callers can either set OR clear recurrence.
    const recurrenceSchemas = updateTask.inputSchema.properties.recurrence.anyOf
    const types = recurrenceSchemas.map((s: any) => s.type)
    expect(types).toContain("object")
    expect(types).toContain("null")
  }, 20_000)
})
