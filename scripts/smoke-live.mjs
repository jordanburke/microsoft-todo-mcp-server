#!/usr/bin/env node
// Live smoke test of the MCP server against a real Microsoft To Do account, recorded as a
// snapshot so two builds (e.g. before and after a refactor) can be compared call by call.
//
//   node scripts/smoke-live.mjs --label before            # run, write .smoke/before.json
//   node scripts/smoke-live.mjs --label after --cmd "node dist/cli.js"
//   node scripts/smoke-live.mjs --compare before after    # diff two snapshots
//
// Writes happen only inside two temporary lists ("MCP smoke … (safe to delete)"), which are
// deleted at the end even if a step fails. Existing lists are only read, and what is kept
// from them is shape (counts, which fields print), not titles. Snapshots hold list counts
// and server output, so .smoke/ is gitignored. Tokens come from MSTODO_TOKEN_FILE
// (default ./tokens.json) and are never printed.

import { spawn } from "node:child_process"
import { execSync } from "node:child_process"
import { mkdirSync, readFileSync, writeFileSync } from "node:fs"
import { resolve } from "node:path"

const SNAPSHOT_DIR = resolve(".smoke")
const args = process.argv.slice(2)
const flag = (name) => {
  const i = args.indexOf(name)
  return i >= 0 ? args[i + 1] : undefined
}

// ---------- normalization: strip what differs run to run, keep what differs build to build ----------

const today = new Date()
const pad = (n) => String(n).padStart(2, "0")
const todayIso = `${today.getFullYear()}-${pad(today.getMonth() + 1)}-${pad(today.getDate())}`
const todayUtcIso = today.toISOString().slice(0, 10)
const todayLocale = today.toLocaleDateString()

const normalize = (text) =>
  String(text)
    .replace(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi, "<guid>")
    .replace(/[A-Za-z0-9_-]{24,}={0,2}/g, "<id>")
    .replace(/T\d{2}:\d{2}:\d{2}(\.\d+)?Z?/g, "T<time>")
    .replace(/\b\d{1,2}:\d{2}:\d{2}( ?[AP]M)?/g, "<time>")
    .split(todayIso)
    .join("<today>")
    .split(todayUtcIso)
    .join("<today>")
    .split(todayLocale)
    .join("<today>")

const sortKeys = (value) =>
  Array.isArray(value)
    ? value.map(sortKeys)
    : value && typeof value === "object"
      ? Object.fromEntries(
          Object.keys(value)
            .sort()
            .map((k) => [k, sortKeys(value[k])]),
        )
      : value

// ---------- compare mode ----------

const lineDiff = (a, b) => {
  const x = a.split("\n")
  const y = b.split("\n")
  const lcs = Array.from({ length: x.length + 1 }, () => new Array(y.length + 1).fill(0))
  for (let i = x.length - 1; i >= 0; i--)
    for (let j = y.length - 1; j >= 0; j--)
      lcs[i][j] = x[i] === y[j] ? lcs[i + 1][j + 1] + 1 : Math.max(lcs[i + 1][j], lcs[i][j + 1])
  const out = []
  let i = 0
  let j = 0
  while (i < x.length || j < y.length) {
    if (i < x.length && j < y.length && x[i] === y[j]) {
      i++
      j++
    } else if (j < y.length && (i === x.length || lcs[i][j + 1] >= lcs[i + 1][j])) out.push(`    + ${y[j++]}`)
    else out.push(`    - ${x[i++]}`)
  }
  return out
}

if (args.includes("--compare")) {
  const [a, b] = args.slice(args.indexOf("--compare") + 1)
  const load = (label) => JSON.parse(readFileSync(resolve(SNAPSHOT_DIR, `${label}.json`), "utf8"))
  const A = load(a)
  const B = load(b)
  console.log(`${a}: ${A.meta.git} ${A.meta.version} ${A.meta.date}`)
  console.log(`${b}: ${B.meta.git} ${B.meta.version} ${B.meta.date}\n`)
  const sections = [
    ["initialize", (s) => JSON.stringify(s.initialize, null, 2)],
    ["tools", (s) => JSON.stringify(s.tools, null, 2)],
  ]
  let differences = 0
  for (const [name, render] of sections) {
    const d = lineDiff(render(A), render(B))
    console.log(`${d.length ? "DIFF" : "same"}  ${name}`)
    if (d.length) {
      differences++
      console.log(d.join("\n"))
    }
  }
  const steps = [...new Set([...A.steps.map((s) => s.step), ...B.steps.map((s) => s.step)])]
  for (const step of steps) {
    const sa = A.steps.find((s) => s.step === step)
    const sb = B.steps.find((s) => s.step === step)
    const ra = sa ? JSON.stringify(sa.result, null, 2) : "(missing)"
    const rb = sb ? JSON.stringify(sb.result, null, 2) : "(missing)"
    const d = lineDiff(ra, rb)
    console.log(
      `${d.length ? "DIFF" : "same"}  ${step}${sa?.realData ? "  [real data: may differ if the account changed]" : ""}`,
    )
    if (d.length) {
      differences++
      console.log(d.join("\n"))
    }
  }
  console.log(`\n${differences} section(s) differ.`)
  process.exit(differences ? 1 : 0)
}

// ---------- run mode ----------

const label = flag("--label")
if (!label) {
  console.error("usage: --label <name> [--cmd '<server command>']  |  --compare <a> <b>")
  process.exit(2)
}
const command = (flag("--cmd") ?? "node dist/cli.js").split(" ")
const tokenFile = resolve(process.env.MSTODO_TOKEN_FILE ?? "tokens.json")

const child = spawn(command[0], command.slice(1), {
  env: { ...process.env, MSTODO_TOKEN_FILE: tokenFile },
  stdio: ["pipe", "pipe", "pipe"],
})
let stderr = ""
child.stderr.on("data", (d) => (stderr += d))
let buffer = ""
let nextId = 0
const pending = new Map()
child.stdout.on("data", (d) => {
  buffer += d
  let nl
  while ((nl = buffer.indexOf("\n")) >= 0) {
    const line = buffer.slice(0, nl)
    buffer = buffer.slice(nl + 1)
    if (!line.trim()) continue
    const msg = JSON.parse(line)
    pending.get(msg.id)?.(msg)
  }
})
child.on("exit", (code) => {
  for (const resolveCall of pending.values()) resolveCall({ error: { message: `server exited (${code})` } })
})
const rpc = (method, params) =>
  new Promise((resolveCall) => {
    const id = ++nextId
    pending.set(id, resolveCall)
    child.stdin.write(JSON.stringify({ jsonrpc: "2.0", id, method, params }) + "\n")
    setTimeout(() => resolveCall({ error: { message: `timeout: ${method}` } }), 60_000)
  })

const steps = []
const rawCall = async (tool, toolArgs) => {
  const msg = await rpc("tools/call", { name: tool, arguments: toolArgs })
  if (msg.error) return { error: msg.error.message, text: "" }
  const content = msg.result?.content ?? []
  return {
    isError: Boolean(msg.result?.isError),
    contentTypes: content.map((c) => c.type),
    text: content.map((c) => c.text ?? "").join("\n"),
  }
}
// Record a deterministic step: the full normalized text.
const record = async (step, tool, toolArgs) => {
  const r = await rawCall(tool, toolArgs)
  steps.push({ step, tool, result: { ...r, text: normalize(r.text) } })
  return r
}
// Record a real-data step: only its shape.
const recordShape = (step, tool, shape) => steps.push({ step, tool, realData: true, result: shape })

const idAfter = (text, label = "ID") => text.match(new RegExp(`${label}: (\\S+)`))?.[1]
const countIds = (text) => (text.match(/ID: /g) ?? []).length
const fieldKinds = (text) => {
  const kinds = {}
  for (const line of text.split("\n")) {
    const kind = /^[✓○] ID:/.test(line)
      ? "status"
      : line.match(/^(Due|Importance|Categories|Description|Repeats):/)?.[1]
    if (kind) kinds[kind] = (kinds[kind] ?? 0) + 1
  }
  return sortKeys(kinds)
}

const git = (() => {
  try {
    return execSync("git rev-parse --short HEAD", { encoding: "utf8" }).trim()
  } catch {
    return "unknown"
  }
})()
const version = JSON.parse(readFileSync("package.json", "utf8")).version

const init = await rpc("initialize", {
  protocolVersion: "2025-06-18",
  capabilities: {},
  clientInfo: { name: "smoke-live", version: "1" },
})
if (init.error) {
  console.error("initialize failed:", init.error.message, "\n", stderr.slice(-2000))
  process.exit(1)
}
child.stdin.write(JSON.stringify({ jsonrpc: "2.0", method: "notifications/initialized" }) + "\n")
const toolsMsg = await rpc("tools/list", {})
const tools = sortKeys(
  (toolsMsg.result?.tools ?? []).map((t) => ({ name: t.name, description: t.description, inputSchema: t.inputSchema })),
).sort((a, b) => a.name.localeCompare(b.name))

let listA
let listB
try {
  // --- auth (also lets the server refresh tokens before anything else)
  await record("auth-status", "auth-status", {})

  // --- temporary lists
  const a = await record("create-task-list A", "create-task-list", { displayName: "MCP smoke A (safe to delete)" })
  listA = idAfter(a.text)
  const b = await record("create-task-list B", "create-task-list", { displayName: "MCP smoke B (safe to delete)" })
  listB = idAfter(b.text)
  if (!listA || !listB) throw new Error("could not create the temporary lists")
  await record("update-task-list rename A", "update-task-list", {
    listId: listA,
    displayName: "MCP smoke A renamed (safe to delete)",
  })

  // --- create-task variants
  const plain = await record("create-task plain", "create-task", { listId: listA, title: "smoke plain" })
  const plainId = idAfter(plain.text)
  const full = await record("create-task all fields", "create-task", {
    listId: listA,
    title: "smoke full",
    body: "smoke body text",
    dueDateTime: "2026-12-01T09:00:00Z",
    startDateTime: "2026-11-25T09:00:00Z",
    importance: "high",
    isReminderOn: true,
    reminderDateTime: "2026-11-30T09:00:00Z",
    status: "inProgress",
    categories: ["Blue category"],
  })
  const fullId = idAfter(full.text)
  const monthly = await record("create-task recurrence monthly", "create-task", {
    listId: listA,
    title: "smoke monthly",
    dueDateTime: "2026-11-15T09:00:00Z",
    recurrence: { pattern: { type: "absoluteMonthly", interval: 1, dayOfMonth: 15 }, range: { type: "noEnd" } },
  })
  const monthlyId = idAfter(monthly.text)
  await record("create-task recurrence weekly numbered", "create-task", {
    listId: listA,
    title: "smoke weekly",
    dueDateTime: "2026-11-02T09:00:00Z",
    recurrence: {
      pattern: { type: "weekly", interval: 2, daysOfWeek: ["monday"] },
      range: { type: "numbered", numberOfOccurrences: 4 },
    },
  })
  await record("create-task recurrence refused (missing dayOfMonth)", "create-task", {
    listId: listA,
    title: "never created",
    dueDateTime: "2026-11-15T09:00:00Z",
    recurrence: { pattern: { type: "absoluteMonthly", interval: 1 }, range: { type: "noEnd" } },
  })
  await record("create-task recurrence refused (no due date)", "create-task", {
    listId: listA,
    title: "never created",
    recurrence: { pattern: { type: "daily", interval: 1 }, range: { type: "noEnd" } },
  })

  // --- get-tasks variants on the temporary list
  await record("get-tasks A", "get-tasks", { listId: listA, orderby: "createdDateTime" })
  await record("get-tasks A select=status", "get-tasks", {
    listId: listA,
    select: "status",
    orderby: "createdDateTime",
  })
  await record("get-tasks A select=recurrence,importance", "get-tasks", {
    listId: listA,
    select: "recurrence,importance",
    orderby: "createdDateTime",
  })
  await record("get-tasks A top=2 orderby=createdDateTime", "get-tasks", {
    listId: listA,
    top: 2,
    orderby: "createdDateTime",
  })
  await record("get-tasks A skip=1 orderby=createdDateTime", "get-tasks", {
    listId: listA,
    skip: 1,
    orderby: "createdDateTime",
  })
  await record("get-tasks A filter importance high", "get-tasks", {
    listId: listA,
    filter: "importance eq 'high'",
    orderby: "createdDateTime",
  })
  await record("get-tasks A count", "get-tasks", { listId: listA, count: true, orderby: "createdDateTime" })

  // --- update-task variants
  await record("update-task title+importance", "update-task", {
    listId: listA,
    taskId: plainId,
    title: "smoke plain renamed",
    importance: "low",
  })
  await record("update-task recurrence change (uses existing due)", "update-task", {
    listId: listA,
    taskId: monthlyId,
    recurrence: { pattern: { type: "daily", interval: 3 }, range: { type: "endDate", endDate: "2027-02-01" } },
  })
  await record("update-task recurrence null", "update-task", { listId: listA, taskId: monthlyId, recurrence: null })
  await record("update-task clear due date", "update-task", { listId: listA, taskId: fullId, dueDateTime: "" })
  await record("update-task no properties", "update-task", { listId: listA, taskId: fullId })
  await record("update-task complete", "update-task", { listId: listA, taskId: plainId, status: "completed" })
  await record("get-tasks A after updates", "get-tasks", { listId: listA, orderby: "createdDateTime" })

  // --- checklist items
  const c1 = await record("create-checklist-item 1", "create-checklist-item", {
    listId: listA,
    taskId: fullId,
    displayName: "step one",
  })
  await record("create-checklist-item 2 checked", "create-checklist-item", {
    listId: listA,
    taskId: fullId,
    displayName: "step two",
    isChecked: true,
  })
  const itemId = idAfter(c1.text)
  await record("get-checklist-items", "get-checklist-items", { listId: listA, taskId: fullId })
  await record("update-checklist-item", "update-checklist-item", {
    listId: listA,
    taskId: fullId,
    checklistItemId: itemId,
    displayName: "step one renamed",
    isChecked: true,
  })
  await record("delete-checklist-item", "delete-checklist-item", {
    listId: listA,
    taskId: fullId,
    checklistItemId: itemId,
  })
  await record("get-checklist-items after delete", "get-checklist-items", { listId: listA, taskId: fullId })

  // --- archive completed tasks from A into B
  await record("archive-completed-tasks dry run", "archive-completed-tasks", {
    sourceListId: listA,
    targetListId: listB,
    olderThanDays: 0,
    dryRun: true,
  })
  await record("archive-completed-tasks", "archive-completed-tasks", {
    sourceListId: listA,
    targetListId: listB,
    olderThanDays: 0,
    dryRun: false,
  })
  await record("get-tasks B after archive", "get-tasks", { listId: listB, orderby: "createdDateTime" })

  // --- delete-task and error paths
  await record("delete-task", "delete-task", { listId: listA, taskId: fullId })
  await record("get-tasks A after delete", "get-tasks", { listId: listA, orderby: "createdDateTime" })
  await record("get-tasks unknown list", "get-tasks", { listId: "AAMkAGNvdF9hX3JlYWxfbGlzdF9pZA" })
  await record("update-task unknown task", "update-task", {
    listId: listA,
    taskId: "AAMkAGNvdF9hX3JlYWxfdGFza19pZA",
    title: "x",
  })
  await record("delete-task unknown task", "delete-task", { listId: listA, taskId: "AAMkAGNvdF9hX3JlYWxfdGFza19pZA" })

  // --- real data, shape only
  const lists = await rawCall("get-task-lists", {})
  const realIds = [...lists.text.matchAll(/ID: (\S+)/g)].map((m) => m[1]).filter((id) => id !== listA && id !== listB)
  recordShape("get-task-lists", "get-task-lists", { isError: lists.isError, lists: realIds.length })
  for (const groupBy of ["category", "shared", "type"]) {
    const r = await rawCall("get-task-lists-organized", { groupBy, includeIds: true })
    recordShape(`get-task-lists-organized ${groupBy}`, "get-task-lists-organized", {
      isError: r.isError,
      // The ID reference prints "name: id" lines, so count the real list IDs that appear.
      listIdsShown: realIds.filter((id) => r.text.includes(id)).length,
    })
  }
  const counts = []
  for (const id of realIds) counts.push({ id, n: countIds((await rawCall("get-tasks", { listId: id })).text) })
  counts.sort((x, y) => y.n - x.n)
  recordShape("get-tasks every list (task counts, largest first)", "get-tasks", { counts: counts.map((c) => c.n) })
  const big = counts[0]?.id
  if (big) {
    for (const [name, extra] of [
      ["no select", {}],
      ["select=status", { select: "status" }],
      ["select=dueDateTime,body", { select: "dueDateTime,body" }],
    ]) {
      const r = await rawCall("get-tasks", { listId: big, top: 25, ...extra })
      recordShape(`get-tasks largest list top=25 ${name}`, "get-tasks", {
        isError: r.isError,
        tasks: countIds(r.text),
        fields: fieldKinds(r.text),
      })
    }
    const t3 = await rawCall("get-tasks", { listId: big, top: 3 })
    recordShape("get-tasks largest list top=3", "get-tasks", { tasks: countIds(t3.text) })
  }
  for (const testType of ["odata-select", "odata-expand", "headers", "extensions"]) {
    const r = await rawCall("test-graph-api-exploration", { testType })
    recordShape(`test-graph-api-exploration ${testType}`, "test-graph-api-exploration", {
      isError: r.isError,
      error: r.error,
      returnedText: r.text.length > 0,
    })
  }
} catch (error) {
  steps.push({ step: "harness error", result: { error: String(error) } })
} finally {
  for (const [name, id] of [
    ["A", listA],
    ["B", listB],
  ]) {
    if (!id) continue
    const r = await rawCall("delete-task-list", { listId: id })
    const remaining = await rawCall("get-task-lists", {})
    const cleaned = /success/i.test(r.text) && !remaining.text.includes(id)
    steps.push({ step: `cleanup delete-task-list ${name}`, result: { deleted: cleaned } })
    if (!cleaned) console.error(`WARNING: temporary list ${name} may still exist; delete "MCP smoke ${name}…" by hand`)
  }
  child.kill()
}

mkdirSync(SNAPSHOT_DIR, { recursive: true })
const snapshot = {
  meta: { label, git, version, date: new Date().toISOString(), node: process.version, command: command.join(" ") },
  initialize: sortKeys({
    protocolVersion: init.result.protocolVersion,
    serverInfo: init.result.serverInfo,
    capabilities: init.result.capabilities,
    instructions: init.result.instructions,
  }),
  tools,
  steps,
}
const out = resolve(SNAPSHOT_DIR, `${label}.json`)
writeFileSync(out, JSON.stringify(snapshot, null, 2) + "\n")
const failures = steps.filter((s) => s.result?.error || s.result?.deleted === false)
console.log(`${steps.length} steps, ${tools.length} tools -> ${out}`)
for (const s of steps) {
  const r = s.result
  const status = r.error ? `ERROR ${r.error}` : r.isError ? "isError" : r.deleted === false ? "NOT CLEANED" : "ok"
  const first = typeof r.text === "string" ? r.text.split("\n")[0].slice(0, 90) : JSON.stringify(r).slice(0, 90)
  console.log(`  ${status.padEnd(8)} ${s.step.padEnd(52)} ${first}`)
}
process.exit(failures.length ? 1 : 0)
