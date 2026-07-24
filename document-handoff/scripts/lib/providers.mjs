import { createReadStream } from 'node:fs'
import { readdir, stat } from 'node:fs/promises'
import { spawnSync } from 'node:child_process'
import { join, basename, dirname, resolve } from 'node:path'
import { homedir } from 'node:os'
import { createInterface } from 'node:readline'
import { fileURLToPath } from 'node:url'

const JSONL_EXT = '.jsonl'

export function detectProvider() {
  if (process.env.CLAUDE_PROJECT_PATH || process.env.CLAUDE_SESSION_ID) return 'claude-code'
  if (process.env.CODEX_SESSION_ID || process.env.CODEX_PROJECT || process.env.CODEX_HOME) return 'codex'
  if (process.env.KIMI_HOME || process.env.KIMI_SESSION_ID) return 'kimi-code'
  return 'unknown'
}

export function listSessionAdapters() {
  return ['claude-code', 'codex', 'kimi-code', 'hermes', 'generic-jsonl']
}

// Strip a Windows extended-path prefix and normalize separators/case for comparison.
export function normPath(value) {
  return resolve(String(value)).replace(/^\\\\\?\\/, '').toLowerCase().replace(/\\/g, '/')
}

function pathWithin(candidate, root) {
  const child = normPath(candidate)
  const parent = normPath(root).replace(/\/$/, '')
  return child === parent || child.startsWith(`${parent}/`)
}

async function walkFiles(root, predicate) {
  const found = []
  const pending = [root]
  while (pending.length) {
    const current = pending.pop()
    let entries
    try { entries = await readdir(current, { withFileTypes: true }) } catch { continue }
    for (const entry of entries) {
      const full = join(current, entry.name)
      if (entry.isDirectory()) pending.push(full)
      else if (entry.isFile() && predicate(full, entry.name)) found.push(full)
    }
  }
  return found
}

async function* jsonRecords(jsonlPath, limit = Number.MAX_SAFE_INTEGER) {
  let input
  try {
    input = createReadStream(jsonlPath, { encoding: 'utf8' })
    const lines = createInterface({ input, crlfDelay: Infinity })
    let count = 0
    for await (const line of lines) {
      if (!line.trim()) continue
      try {
        yield JSON.parse(line)
        count++
      } catch {}
      if (count >= limit) break
    }
  } catch {
    return
  } finally {
    input?.destroy()
  }
}

async function firstJsonRecords(jsonlPath, limit = 50) {
  const records = []
  for await (const record of jsonRecords(jsonlPath, limit)) records.push(record)
  return records
}

function firstString(values) {
  return values.find(value => typeof value === 'string' && value.trim()) || null
}

function recordCwd(record) {
  return firstString([
    record?.cwd,
    record?.workspace,
    record?.workspace_path,
    record?.project_path,
    record?.message?.cwd,
    record?.payload?.cwd,
    record?.payload?.workspace,
    record?.payload?.project_path,
    record?.payload?.metadata?.cwd,
  ])
}

async function jsonlCwd(path) {
  for (const record of await firstJsonRecords(path)) {
    const cwd = recordCwd(record)
    if (cwd) return cwd
  }
  return null
}

function normalizeProviderSpecs(config = {}) {
  const configured = config.providers
  const values = configured === undefined ? ['claude-code', 'codex'] : configured
  if (!Array.isArray(values)) throw new Error('session_capture.providers must be an array')
  return values.map(value => {
    if (typeof value === 'string') return { type: value }
    if (!value || typeof value !== 'object' || typeof value.type !== 'string') {
      throw new Error('each session provider must be a name or an object with type')
    }
    return value
  })
}

function selectedIds(spec) {
  const values = spec.session_ids || spec.include_session_ids || []
  return new Set(values.map(String))
}

function selectedByIdOrProject(id, cwd, sourceRoot, spec) {
  const ids = selectedIds(spec)
  if (ids.size) return ids.has(String(id))
  const roots = [...(spec.cwd_prefixes || []), sourceRoot].filter(Boolean)
  return Boolean(cwd && roots.some(root => pathWithin(cwd, root)))
}

function jsonlSession(provider, path, id, extra = {}) {
  return {
    provider,
    adapter: provider,
    kind: 'jsonl',
    path,
    session_id: String(id),
    session_name: String(extra.session_name || id),
    active: Boolean(extra.active),
    metadata: extra.metadata || {},
  }
}

async function discoverClaude(sourceRoot, spec = {}) {
  const root = spec.root || join(homedir(), '.claude', 'projects')
  const files = await walkFiles(root, (_path, name) => name.endsWith(JSONL_EXT))
  const found = []
  for (const path of files) {
    const id = basename(path, JSONL_EXT)
    const cwd = await jsonlCwd(path)
    if (selectedByIdOrProject(id, cwd, sourceRoot, spec)) {
      found.push(jsonlSession('claude-code', path, id, { metadata: { cwd } }))
    }
  }
  return found
}

async function newestCodexDb(codexHome) {
  let files
  try { files = await readdir(codexHome) } catch { return null }
  const candidates = files.filter(name => /^state(?:_\d+)?\.sqlite$/.test(name))
  if (!candidates.length) return null
  const ranked = await Promise.all(candidates.map(async name => ({
    path: join(codexHome, name),
    mtime: (await stat(join(codexHome, name))).mtimeMs,
  })))
  ranked.sort((left, right) => right.mtime - left.mtime)
  return ranked[0].path
}

async function codexThreadRows(dbPath) {
  if (!dbPath) return []
  let db
  try {
    const { DatabaseSync } = await import('node:sqlite')
    db = new DatabaseSync(dbPath, { readOnly: true })
    db.exec('PRAGMA query_only=ON')
    const columns = db.prepare('PRAGMA table_info(threads)').all().map(row => row.name)
    if (!columns.includes('id') || !columns.includes('cwd')) return []
    const projection = ['id', 'cwd', columns.includes('rollout_path') ? 'rollout_path' : 'NULL AS rollout_path']
    return db.prepare(`SELECT ${projection.join(', ')} FROM threads`).all()
  } catch {
    return []
  } finally {
    try { db?.close() } catch {}
  }
}

function codexIdFromName(path) {
  const match = basename(path).match(/([0-9a-f]{8}-[0-9a-f-]{27})/i)
  return match?.[1] || basename(path, JSONL_EXT)
}

async function discoverCodex(sourceRoot, spec = {}) {
  const home = spec.root || process.env.CODEX_HOME || join(homedir(), '.codex')
  const rows = await codexThreadRows(spec.database || await newestCodexDb(home))
  const selected = new Map()
  for (const row of rows) {
    if (selectedByIdOrProject(row.id, row.cwd, sourceRoot, spec)) selected.set(String(row.id), row)
  }

  const paths = []
  for (const area of ['sessions', 'archived_sessions']) {
    for (const path of await walkFiles(join(home, area), (_path, name) => name.endsWith(JSONL_EXT))) {
      paths.push({ path, active: area === 'sessions' })
    }
  }

  const found = []
  for (const item of paths) {
    const id = codexIdFromName(item.path)
    let row = selected.get(id)
    if (!row && rows.length === 0) {
      const cwd = await jsonlCwd(item.path)
      if (selectedByIdOrProject(id, cwd, sourceRoot, spec)) row = { id, cwd }
    }
    if (row || selectedIds(spec).has(id)) {
      found.push(jsonlSession('codex', item.path, id, {
        active: item.active,
        metadata: { cwd: row?.cwd || null },
      }))
    }
  }

  for (const row of selected.values()) {
    if (!row.rollout_path || found.some(item => item.session_id === String(row.id))) continue
    found.push(jsonlSession('codex', row.rollout_path, row.id, {
      active: pathWithin(row.rollout_path, join(home, 'sessions')),
      metadata: { cwd: row.cwd || null },
    }))
  }
  return found
}

function kimiId(path) {
  const name = basename(path, JSONL_EXT)
  if (name !== 'wire' && name !== 'events' && name !== 'messages') return name
  return basename(dirname(path))
}

async function discoverKimi(sourceRoot, spec = {}) {
  const root = spec.root || process.env.KIMI_HOME || join(homedir(), '.kimi-code', 'sessions')
  const files = await walkFiles(root, (_path, name) => name.endsWith(JSONL_EXT))
  const workspaces = new Set((spec.workspace_ids || []).map(String))
  const found = []
  for (const path of files) {
    const id = kimiId(path)
    const records = await firstJsonRecords(path)
    const cwd = records.map(recordCwd).find(Boolean) || null
    const workspaceId = records.map(record => firstString([
      record?.workspace_id, record?.workspaceId, record?.payload?.workspace_id,
    ])).find(Boolean) || basename(dirname(dirname(path)))
    const explicitWorkspace = workspaces.size && workspaces.has(String(workspaceId))
    if (explicitWorkspace || selectedByIdOrProject(id, cwd, sourceRoot, spec)) {
      found.push(jsonlSession('kimi-code', path, id, {
        active: Boolean(spec.active),
        metadata: { cwd, workspace_id: workspaceId },
      }))
    }
  }
  return found
}

const HERMES_SQLITE_BRIDGE = fileURLToPath(new URL('./hermes_sqlite.py', import.meta.url))

function runHermesSqlite(request) {
  const defaults = process.platform === 'win32' ? ['python', 'py'] : ['python3', 'python']
  const commands = [process.env.PYTHON, ...defaults].filter((value, index, values) =>
    value && values.indexOf(value) === index)
  for (const command of commands) {
    const result = spawnSync(command, [HERMES_SQLITE_BRIDGE], {
      input: JSON.stringify(request), encoding: 'utf8', windowsHide: true,
      maxBuffer: 256 * 1024 * 1024,
    })
    if (result.error?.code === 'ENOENT') continue
    if (result.status !== 0) throw new Error('Hermes SQLite bridge failed')
    try { return JSON.parse(result.stdout) } catch { throw new Error('Hermes SQLite bridge returned invalid JSON') }
  }
  throw new Error('Hermes session capture requires Python 3')
}

async function discoverHermes(sourceRoot, spec = {}) {
  const defaultRoot = process.env.LOCALAPPDATA
    ? join(process.env.LOCALAPPDATA, 'hermes')
    : join(homedir(), '.hermes')
  const path = spec.database || spec.path || join(defaultRoot, 'state.db')
  const sessionTable = spec.session_table || 'sessions'
  const result = runHermesSqlite({
    action: 'discover', database: path, session_table: sessionTable,
    id_column: spec.id_column || null, cwd_column: spec.cwd_column || null,
    session_ids: [...selectedIds(spec)],
  })
  const idColumn = result.id_column
  if (!idColumn) return []
  const cwdColumn = result.cwd_column
  const found = []
  for (const row of result.rows || []) {
    const id = String(row[idColumn])
    const cwd = cwdColumn ? row[cwdColumn] : null
    if (!selectedByIdOrProject(id, cwd, sourceRoot, spec)) continue
    found.push({
      provider: 'hermes', adapter: 'hermes', kind: 'sqlite-session', path,
      session_id: id, session_name: id, active: Boolean(spec.active),
      metadata: { cwd: cwd || null },
      sqlite: { session_table: sessionTable, id_column: idColumn },
    })
  }
  return found
}

function manualSources(config = {}) {
  const values = config.sources || []
  if (!Array.isArray(values)) throw new Error('session_capture.sources must be an array')
  return values.map((source, index) => {
    if (!source || typeof source !== 'object' || !source.path) {
      throw new Error(`session_capture.sources[${index}] requires path`)
    }
    const kind = source.kind || 'jsonl'
    if (!['jsonl', 'supporting'].includes(kind)) {
      throw new Error(`unsupported manual session source kind: ${kind}`)
    }
    const id = String(source.session_id || source.id || basename(source.path, JSONL_EXT))
    return {
      provider: source.provider || 'generic-jsonl',
      adapter: source.adapter || 'generic-jsonl',
      kind,
      path: source.path,
      session_id: id,
      session_name: String(source.session_name || id),
      active: Boolean(source.active),
      metadata: source.metadata || {},
    }
  })
}

const DISCOVERERS = new Map([
  ['claude-code', discoverClaude],
  ['codex', discoverCodex],
  ['kimi-code', discoverKimi],
  ['hermes', discoverHermes],
])

export async function discoverSessions(sourceRoot, config = {}) {
  const found = []
  for (const spec of normalizeProviderSpecs(config)) {
    const discover = DISCOVERERS.get(spec.type)
    if (!discover) throw new Error(`unknown session provider: ${spec.type}`)
    found.push(...await discover(sourceRoot, spec))
  }
  found.push(...manualSources(config))
  const unique = new Map()
  for (const item of found) {
    const key = `${item.provider}\0${normPath(item.path)}\0${item.session_id}\0${item.kind}`
    if (!unique.has(key)) unique.set(key, item)
  }
  return [...unique.values()].sort((left, right) =>
    left.provider.localeCompare(right.provider) || left.session_name.localeCompare(right.session_name))
}

function textFromContent(content) {
  if (typeof content === 'string') return content
  if (!Array.isArray(content)) return ''
  return content
    .filter(item => item && typeof item === 'object' && ['text', 'input_text', 'output_text'].includes(item.type))
    .map(item => item.text || '')
    .filter(Boolean)
    .join('\n')
}

export async function extractCC(jsonlPath) {
  const msgs = []
  for await (const obj of jsonRecords(jsonlPath)) {
    const role = obj.message?.role || obj.role || obj.type || 'unknown'
    const text = textFromContent(obj.message?.content ?? obj.content)
    if (text) msgs.push(`[${role}] ${text}`)
  }
  return msgs
}

export async function extractCodex(jsonlPath) {
  const msgs = []
  for await (const obj of jsonRecords(jsonlPath)) {
    if (obj.type !== 'response_item') continue
    const role = obj.payload?.role || 'unknown'
    const text = textFromContent(obj.payload?.content)
    if (text) msgs.push(`[${role}] ${text}`)
  }
  return msgs
}

export async function extractKimi(jsonlPath) {
  const msgs = []
  for await (const obj of jsonRecords(jsonlPath)) {
    const payload = obj.message || obj.payload || obj
    const role = payload.role || obj.role || obj.type || 'unknown'
    const text = firstString([
      textFromContent(payload.content), payload.text, payload.prompt, payload.response,
    ])
    if (text) msgs.push(`[${role}] ${text}`)
  }
  return msgs
}

export async function exportHermesRows(session) {
  const result = runHermesSqlite({
    action: 'export', database: session.path, session_id: session.session_id,
    session_table: session.sqlite?.session_table || 'sessions',
    id_column: session.sqlite?.id_column || 'id',
  })
  return result.rows || []
}

export async function extractHermes(session) {
  const msgs = []
  for (const { table, row } of await exportHermesRows(session)) {
    const role = firstString([row.role, row.author, row.sender]) || table
    const text = firstString([row.content, row.text, row.message, row.prompt, row.response])
    if (text) msgs.push(`[${role}] ${text}`)
  }
  return msgs
}

export async function extractSession(session) {
  if (session.kind === 'supporting') return []
  if (session.kind === 'sqlite-session') return extractHermes(session)
  if (session.adapter === 'claude-code') return extractCC(session.path)
  if (session.adapter === 'codex') return extractCodex(session.path)
  if (session.adapter === 'kimi-code') return extractKimi(session.path)
  return extractKimi(session.path)
}

export function normalizeExtract(messages, provider, sessionName) {
  return `# Session: ${sessionName} (${provider})\n\n${messages.join('\n\n')}`
}

// Backward-compatible entry points used by existing callers.
export function findCCSessions(sourceRoot) { return discoverClaude(sourceRoot, {}) }
export function findCodexSessions(sourceRoot) { return discoverCodex(sourceRoot, {}) }
