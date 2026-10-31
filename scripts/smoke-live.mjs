#!/usr/bin/env node
// Live smoke test of the MCP server against a real Microsoft To Do account, recorded as a
// snapshot so two builds (e.g. before and after a refactor) can be compared call by call.
//
//   node scripts/smoke-live.mjs --label before            # run, write .smoke/before.json
//   node scripts/smoke-live.mjs --label after --cmd "node dist/cli.js"
//   node scripts/smoke-live.mjs --compare before after    # diff two snapshots
//
// Writes happen only inside two temporary lists ("MCP smoke … (safe to delete)"), which are
// deleted at the end even if a step fails or the server dies; leftovers from an earlier
// crashed run are found by name and deleted too. Existing lists are only read, and what is
// kept from them is shape (counts, which fields print), not titles. Snapshots hold list
// counts and server output, so .smoke/ is gitignored. Tokens come from MSTODO_TOKEN_FILE
// (default ./tokens.json) and are never printed.

import { execSync, spawn } from "node:child_process"
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs"
import { resolve } from "node:path"

// One clock and locale for the harness and the server, so a run on another machine or near
// midnight does not differ from the baseline in how dates print.
process.env.TZ = "UTC"
const LOCALE_ENV = { TZ: "UTC", LANG: "en_US.UTF-8", LC_ALL: "en_US.UTF-8" }

const SNAPSHOT_DIR = resolve(".smoke")
const SMOKE_NAMES = [
  "MCP smoke A (safe to delete)",
  "MCP smoke A renamed (safe to delete)",
  "MCP smoke B (safe to delete)",
]
const args = process.argv.slice(2)
const flag = (name) => {
  const i = args.indexOf(name)
  return i >= 0 ? args[i + 1] : undefined
}

// ---------- normalization: strip what differs run to run, keep what differs build to build ----------

const today = new Date()
const todayIso = today.toISOString().slice(0, 10)
const todayUs = today.toLocaleDateString("en-US", { timeZone: "UTC" })

const normalize = (text) =>
  String(text)
    .replace(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi, "<guid>")
    // Graph IDs are long and contain digits; a long all-letter word (an error code such as
    // MailboxNotEnabledForRESTAPI) is left alone so a change to it still shows.
    .replace(/(?=[A-Za-z_-]*\d)[A-Za-z0-9_-]{24,}={0,2}/g, "<id>")
    .replace(/(Token expires at )[^\n]*/g, "$1<when>")
    .replace(/T\d{2}:\d{2}:\d{2}(\.\d+)?Z?/g, "T<time>")
    .replace(/\b\d{1,2}:\d{2}:\d{2}( ?[AP]M)?/g, "<time>")
    .split(todayIso)
    .join("<today>")
    .split(todayUs)
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
  const load = (label) => {
    const path = resolve(SNAPSHOT_DIR, `${label}.json`)
    if (!label || !existsSync(path)) {
      console.error(`no snapshot ${path}`)
      process.exit(2)
    }
    return JSON.parse(readFileSync(path, "utf8"))
  }
  const A = load(a)
  const B = load(b)
  console.log(`${a}: ${A.meta.git} ${A.meta.version} ${A.meta.date}`)
  console.log(`${b}: ${B.meta.git} ${B.meta.version} ${B.meta.date}\n`)
  let behaviour = 0
  let realData = 0
  const show = (name, ra, rb, isRealData) => {
    const d = lineDiff(ra, rb)
    console.log(
      `${d.length ? "DIFF" : "same"}  ${name}${isRealData ? "  [real data: also changes when the account does]" : ""}`,
    )
    if (!d.length) return
    console.log(d.join("\n"))
    if (isRealData) realData++
    else behaviour++
  }
  for (const section of ["initialize", "tools", "protocol"]) {
    show(section, JSON.stringify(A[section] ?? null, null, 2), JSON.stringify(B[section] ?? null, null, 2), false)
  }
  const steps = [...new Set([...A.steps.map((s) => s.step), ...B.steps.map((s) => s.step)])]
  for (const step of steps) {
    const sa = A.steps.find((s) => s.step === step)
    const sb = B.steps.find((s) => s.step === step)
    show(
      step,
      sa ? JSON.stringify(sa.result, null, 2) : "(missing)",
      sb ? JSON.stringify(sb.result, null, 2) : "(missing)",
      Boolean((sa ?? sb).realData),
    )
  }
  console.log(`\n${behaviour} behaviour section(s) differ; ${realData} real-data section(s) differ.`)
  // Real-data drift alone is expected between runs days apart, so only behaviour fails the gate.
  process.exit(behaviour ? 1 : 0)
}

// ---------- run mode ----------

const label = flag("--label")
if (!label) {
  console.error("usage: --label <name> [--cmd '<server command>']  |  --compare <a> <b>")
  process.exit(2)
}
const command = (flag("--cmd") ?? "node dist/cli.js").split(" ")
const tokenFile = resolve(process.env.MSTODO_TOKEN_FILE ?? "tokens.json")
const protocol = { strayStdoutLines: 0, serverRequests: [], serverExitedEarly: false }

// A server over stdio. Never throws: a dead or silent server turns each call into an error
// result, so the run always reaches cleanup.
const connect = () => {
  const child = spawn(command[0], command.slice(1), {
    env: { ...process.env, ...LOCALE_ENV, MSTODO_TOKEN_FILE: tokenFile },
    stdio: ["pipe", "pipe", "pipe"],
  })
  const server = { stderr: "", exited: false, pending: new Map(), nextId: 0 }
  child.stderr.on("data", (d) => (server.stderr += d))
  child.on("error", (e) => (server.stderr += `spawn error: ${e.message}\n`))
  let buffer = ""
  child.stdout.on("data", (d) => {
    buffer += d
    let nl
    while ((nl = buffer.indexOf("\n")) >= 0) {
      const line = buffer.slice(0, nl)
      buffer = buffer.slice(nl + 1)
      if (!line.trim()) continue
      let msg
      try {
        msg = JSON.parse(line)
      } catch {
        // Anything but JSON-RPC on stdout breaks real MCP clients, so it is counted.
        protocol.strayStdoutLines++
        process.stderr.write(`stray stdout: ${line.slice(0, 200)}\n`)
        continue
      }
      // A server-initiated request or notification carries a method; its id is the
      // server's own and must not resolve one of ours.
      if (msg.method) {
        if (msg.id !== undefined) protocol.serverRequests.push(msg.method)
        continue
      }
      server.pending.get(msg.id)?.(msg)
    }
  })
  child.on("exit", (code) => {
    server.exited = true
    for (const done of server.pending.values()) done({ error: { message: `server exited (${code})` } })
    server.pending.clear()
  })
  server.rpc = (method, params) =>
    new Promise((done) => {
      if (server.exited) return done({ error: { message: "server not running" } })
      const id = ++server.nextId
      const timer = setTimeout(() => {
        server.pending.delete(id)
        done({ error: { message: `timeout: ${method}` } })
      }, 60_000)
      server.pending.set(id, (msg) => {
        clearTimeout(timer)
        server.pending.delete(id)
        done(msg)
      })
      child.stdin.write(JSON.stringify({ jsonrpc: "2.0", id, method, params }) + "\n")
    })
  server.notify = (method) => !server.exited && child.stdin.write(JSON.stringify({ jsonrpc: "2.0", method }) + "\n")
  server.kill = () => child.kill()
  return server
}

const handshake = async (server) => {
  const init = await server.rpc("initialize", {
    protocolVersion: "2025-06-18",
    capabilities: {},
    clientInfo: { name: "smoke-live", version: "1" },
  })
  if (!init.error) server.notify("notifications/initialized")
  return init
}

let server = connect()
const steps = []

const callOn = async (srv, tool, toolArgs) => {
  const msg = await srv.rpc("tools/call", { name: tool, arguments: toolArgs })
  if (msg.error) {
    return { error: msg.error.message, errorCode: msg.error.code, errorData: msg.error.data, text: "" }
  }
  const result = msg.result ?? {}
  const content = result.content ?? []
  return {
    isError: Boolean(result.isError),
    resultKeys: Object.keys(result).sort(),
    content: content.map((c) => ({ type: c.type, keys: Object.keys(c).sort() })),
    structuredContent: result.structuredContent,
    text: content.map((c) => c.text ?? "").join("\n"),
  }
}
const rawCall = (tool, toolArgs) => callOn(server, tool, toolArgs)
const normalized = (r) => ({
  ...r,
  error: r.error && normalize(r.error),
  errorData: r.errorData && JSON.parse(normalize(JSON.stringify(r.errorData))),
  structuredContent: r.structuredContent && JSON.parse(normalize(JSON.stringify(r.structuredContent))),
  text: normalize(r.text),
})
// Record a deterministic step: the full normalized result.
const record = async (step, tool, toolArgs) => {
  const r = await rawCall(tool, toolArgs)
  steps.push({ step, tool, result: sortKeys(normalized(r)) })
  return r
}
// Record a real-data step: only its shape.
const recordShape = (step, tool, shape) => steps.push({ step, tool, realData: true, result: sortKeys(shape) })

const idAfter = (text) => text.match(/ID: (\S+)/)?.[1]
// Task lines start with "ID:" (optionally after a status mark); "No tasks found in list
// with ID: …" does not, so an empty list counts as zero.
const countTasks = (text) => (text.match(/^(?:[✓○] )?ID: /gm) ?? []).length
const fieldKinds = (text) => {
  const kinds = {}
  for (const line of text.split("\n")) {
    const kind = /^[✓○] ID:/.test(line)
      ? "status"
      : line.match(/^(Due|Importance|Categories|Description|Repeats):/)?.[1]
    if (kind) kinds[kind] = (kinds[kind] ?? 0) + 1
  }
  return kinds
}
// get-task-lists prints "ID: <id>\nName: <name>…" per list. Null when the listing failed,
// so a failure is never mistaken for "the list is gone".
const listLists = async (srv) => {
  const r = await callOn(srv, "get-task-lists", {})
  if (r.error || r.isError || !r.text.startsWith("Your task lists:")) return null
  return [...r.text.matchAll(/ID: (\S+)\nName: ([^\n]*)/g)].map((m) => ({ id: m[1], name: m[2] }))
}

const git = (() => {
  try {
    return execSync("git rev-parse --short HEAD", { encoding: "utf8" }).trim()
  } catch {
    return "unknown"
  }
})()
const version = JSON.parse(readFileSync("package.json", "utf8")).version

const init = await handshake(server)
if (init.error) {
  console.error("initialize failed:", init.error.message, "\n", server.stderr.slice(-2000))
  server.kill()
  process.exit(1)
}
const toolsMsg = await server.rpc("tools/list", {})
const tools = sortKeys(toolsMsg.result?.tools ?? []).sort((a, b) => a.name.localeCompare(b.name))

let listA
let listB
try {
  // --- auth (also lets the server refresh tokens before anything else)
  await record("auth-status", "auth-status", {})

  // --- temporary lists
  const a = await record("create-task-list A", "create-task-list", { displayName: SMOKE_NAMES[0] })
  listA = idAfter(a.text)
  const b = await record("create-task-list B", "create-task-list", { displayName: SMOKE_NAMES[2] })
  listB = idAfter(b.text)
  if (!listA || !listB) throw new Error("a temporary list was not confirmed; cleanup looks for it by name")
  await record("update-task-list rename A", "update-task-list", { listId: listA, displayName: SMOKE_NAMES[1] })

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

  // --- get-tasks variants on the temporary list (Graph rejects $orderby=title)
  const byCreated = { listId: listA, orderby: "createdDateTime" }
  await record("get-tasks A", "get-tasks", byCreated)
  await record("get-tasks A select=status", "get-tasks", { ...byCreated, select: "status" })
  await record("get-tasks A select=recurrence,importance", "get-tasks", {
    ...byCreated,
    select: "recurrence,importance",
  })
  await record("get-tasks A top=2 orderby=createdDateTime", "get-tasks", { ...byCreated, top: 2 })
  await record("get-tasks A skip=1 orderby=createdDateTime", "get-tasks", { ...byCreated, skip: 1 })
  await record("get-tasks A filter importance high", "get-tasks", { ...byCreated, filter: "importance eq 'high'" })
  await record("get-tasks A count", "get-tasks", { ...byCreated, count: true })

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
  await record("get-tasks A after updates", "get-tasks", byCreated)

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
  await record("get-tasks A after delete", "get-tasks", byCreated)
  await record("get-tasks unknown list", "get-tasks", { listId: "AAMkAGNvdF9hX3JlYWxfbGlzdF9pZA" })
  await record("update-task unknown task", "update-task", {
    listId: listA,
    taskId: "AAMkAGNvdF9hX3JlYWxfdGFza19pZA",
    title: "x",
  })
  await record("delete-task unknown task", "delete-task", { listId: listA, taskId: "AAMkAGNvdF9hX3JlYWxfdGFza19pZA" })

  // --- protocol-level errors: how argument validation and unknown tools are reported
  await record("get-tasks wrong argument type", "get-tasks", { listId: listA, top: "2" })
  await record("get-tasks missing required argument", "get-tasks", {})
  await record("unknown tool", "no-such-tool", {})

  // --- real data, shape only (smoke-named lists, including leftovers, are excluded)
  const lists = (await listLists(server)) ?? []
  const realIds = lists.filter((l) => !SMOKE_NAMES.includes(l.name)).map((l) => l.id)
  recordShape("get-task-lists", "get-task-lists", { listed: lists.length > 0, lists: realIds.length })
  for (const groupBy of ["category", "shared", "type"]) {
    const r = await rawCall("get-task-lists-organized", { groupBy, includeIds: true })
    recordShape(`get-task-lists-organized ${groupBy}`, "get-task-lists-organized", {
      isError: r.isError,
      error: r.error,
      // The ID reference prints "name: id" lines, so count the real list IDs that appear.
      listIdsShown: realIds.filter((id) => r.text.includes(id)).length,
    })
  }
  const counts = []
  for (const id of realIds) counts.push({ id, n: countTasks((await rawCall("get-tasks", { listId: id })).text) })
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
        tasks: countTasks(r.text),
        fields: fieldKinds(r.text),
      })
    }
    const t3 = await rawCall("get-tasks", { listId: big, top: 3 })
    recordShape("get-tasks largest list top=3", "get-tasks", { tasks: countTasks(t3.text) })
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
  // Cleanup needs a live server; start a fresh one if the first has died.
  // A server that died before cleanup fails the run, whatever the snapshot says.
  protocol.serverExitedEarly = server.exited
  let cleaner = server
  if (server.exited) {
    cleaner = connect()
    await handshake(cleaner)
  }
  const before = await listLists(cleaner)
  const targets = new Map()
  for (const id of [listA, listB]) if (id) targets.set(id, "this run")
  for (const l of before ?? []) if (SMOKE_NAMES.includes(l.name) && !targets.has(l.id)) targets.set(l.id, "leftover")
  for (const [id] of targets) await callOn(cleaner, "delete-task-list", { listId: id })
  const after = await listLists(cleaner)
  const leftovers = [...targets].filter(([, origin]) => origin === "leftover").length
  // delete-task-list reports success whatever Graph says, so only a fresh listing that no
  // longer shows the lists counts as proof.
  const verified = before !== null && after !== null
  const remaining = after?.filter((l) => targets.has(l.id) || SMOKE_NAMES.includes(l.name)).length
  const cleanup = { verified, remaining: verified ? remaining : "unknown", leftoversFromEarlierRuns: leftovers }
  steps.push({ step: "cleanup", result: { verified, remaining: cleanup.remaining } })
  if (!verified || remaining) {
    console.error(`WARNING: cleanup ${verified ? `left ${remaining} list(s)` : "could not be verified"}.`)
    console.error(`Delete any list named ${SMOKE_NAMES.map((n) => `"${n}"`).join(", ")} by hand.`)
  }
  if (leftovers) console.error(`Also deleted ${leftovers} smoke list(s) left by an earlier run.`)
  cleaner.kill()
  if (cleaner !== server) server.kill()
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
  protocol: sortKeys(protocol),
  steps,
}
const out = resolve(SNAPSHOT_DIR, `${label}.json`)
writeFileSync(out, JSON.stringify(snapshot, null, 2) + "\n")
const cleanupStep = steps.find((s) => s.step === "cleanup")?.result
const failed =
  protocol.serverExitedEarly ||
  steps.some((s) => s.step === "harness error") ||
  !cleanupStep?.verified ||
  cleanupStep.remaining
console.log(`${steps.length} steps, ${tools.length} tools, ${protocol.strayStdoutLines} stray stdout line(s) -> ${out}`)
for (const s of steps) {
  const r = s.result
  const status = r.error ? "error" : r.isError ? "isError" : "ok"
  const first = typeof r.text === "string" ? r.text.split("\n")[0].slice(0, 90) : JSON.stringify(r).slice(0, 90)
  console.log(`  ${status.padEnd(8)} ${s.step.padEnd(52)} ${first}`)
}
process.exit(failed ? 1 : 0)
