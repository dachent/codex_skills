import { test } from 'node:test'
import assert from 'node:assert/strict'
import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { spawnSync } from 'node:child_process'
import { readState, writeState, STATE_DEFAULT } from '../lib/state.mjs'
import { run as extractSessions } from '../phases/extract-sessions.mjs'
import {
  discoverSessions,
  extractHermes,
  extractKimi,
  listSessionAdapters,
} from '../lib/providers.mjs'
import { captureSessionEvidence, verifySessionEvidence } from '../lib/session-evidence.mjs'

async function tempFolder(label) {
  return mkdtemp(join(tmpdir(), `document-handoff-${label}-`))
}

function runPython(script, args = []) {
  const defaults = process.platform === 'win32' ? ['python', 'py'] : ['python3', 'python']
  const commands = [process.env.PYTHON, ...defaults].filter((value, index, values) =>
    value && values.indexOf(value) === index)
  for (const command of commands) {
    const result = spawnSync(command, ['-c', script, ...args], { encoding: 'utf8', windowsHide: true })
    if (result.error?.code === 'ENOENT') continue
    assert.equal(result.status, 0, result.stderr || 'Python fixture creation failed')
    return
  }
  assert.fail('Python 3 is required for the Hermes adapter test')
}

test('adapter registry covers the supported harnesses and generic JSONL', () => {
  assert.deepEqual(listSessionAdapters(), [
    'claude-code', 'codex', 'kimi-code', 'hermes', 'generic-jsonl',
  ])
})

test('declared sources make the capture harness extensible without code changes', async () => {
  const root = await tempFolder('declared')
  try {
    const path = join(root, 'custom.jsonl')
    await writeFile(path, '{"role":"user","content":"hello"}\n')
    const sessions = await discoverSessions(root, {
      providers: [],
      sources: [{ path, provider: 'custom-agent', session_id: 'custom-1', active: true }],
    })
    assert.equal(sessions.length, 1)
    assert.equal(sessions[0].provider, 'custom-agent')
    assert.equal(sessions[0].session_id, 'custom-1')
  } finally {
    await rm(root, { recursive: true, force: true })
  }
})

test('JSONL capture keeps a stable complete-record prefix and deduplicates blobs', async () => {
  const root = await tempFolder('jsonl')
  try {
    const source = join(root, 'active.jsonl')
    await writeFile(source, '{"n":1}\n{"n":2}\n{"unfinished":')
    const output = join(root, 'output')
    const sessions = ['one', 'two'].map(sessionId => ({
      provider: 'test', adapter: 'generic-jsonl', kind: 'jsonl', path: source,
      session_id: sessionId, session_name: sessionId, active: true, metadata: {},
    }))
    const result = await captureSessionEvidence(sessions, output)
    assert.equal(result.manifest.source_count, 2)
    assert.equal(result.manifest.unique_blob_count, 1)
    assert.equal(result.manifest.sources[0].records, 2)
    assert.equal(result.manifest.sources[0].complete_record_prefix, true)
    assert.equal(result.manifest.sources[0].lifecycle, 'active-prefix')
    const blob = join(result.evidenceRoot, result.manifest.sources[0].blob_path)
    assert.equal(await readFile(blob, 'utf8'), '{"n":1}\n{"n":2}\n')
    assert.deepEqual(await verifySessionEvidence(result.manifestPath), {
      passed: true, failures: [], source_count: 2,
    })
  } finally {
    await rm(root, { recursive: true, force: true })
  }
})

test('supporting files use the stable snapshot path', async () => {
  const root = await tempFolder('supporting')
  try {
    const source = join(root, 'registry.json')
    await writeFile(source, '{"sessions":["a"]}\n')
    const result = await captureSessionEvidence([{
      provider: 'test', adapter: 'generic-jsonl', kind: 'supporting', path: source,
      session_id: 'registry', session_name: 'registry', active: false, metadata: {},
    }], join(root, 'output'))
    assert.equal(result.manifest.sources[0].records, null)
    assert.equal(result.manifest.sources[0].lifecycle, 'snapshot')
    assert.match(result.manifest.sources[0].blob_path, /\.json$/)
  } finally {
    await rm(root, { recursive: true, force: true })
  }
})

test('Kimi wire JSONL is normalized through its adapter', async () => {
  const root = await tempFolder('kimi')
  try {
    const path = join(root, 'wire.jsonl')
    await writeFile(path, '{"payload":{"role":"user","content":"review this"}}\n')
    const messages = await extractKimi(path)
    assert.deepEqual(messages, ['[user] review this'])
  } finally {
    await rm(root, { recursive: true, force: true })
  }
})

test('Hermes adapter exports selected sessions from a read transaction', async () => {
  const root = await tempFolder('hermes')
  try {
    const dbPath = join(root, 'state.db')
    runPython([
      'import sqlite3, sys',
      'db = sqlite3.connect(sys.argv[1])',
      'db.execute("CREATE TABLE sessions (id TEXT PRIMARY KEY, cwd TEXT)")',
      'db.execute("CREATE TABLE messages (id INTEGER PRIMARY KEY, session_id TEXT, role TEXT, content TEXT)")',
      'db.execute("INSERT INTO sessions VALUES (?, ?)", ("keep", sys.argv[2]))',
      'db.execute("INSERT INTO sessions VALUES (?, ?)", ("skip", sys.argv[2]))',
      'db.execute("INSERT INTO messages(session_id, role, content) VALUES (?, ?, ?)", ("keep", "user", "captured"))',
      'db.execute("INSERT INTO messages(session_id, role, content) VALUES (?, ?, ?)", ("skip", "user", "not captured"))',
      'db.commit()',
      'db.close()',
    ].join('\n'), [dbPath, root])

    const sessions = await discoverSessions(root, {
      providers: [{ type: 'hermes', database: dbPath, session_ids: ['keep'] }],
    })
    assert.equal(sessions.length, 1)
    assert.equal(sessions[0].session_id, 'keep')
    assert.deepEqual(await extractHermes(sessions[0]), ['[user] captured'])

    const result = await captureSessionEvidence(sessions, join(root, 'output'))
    const blob = join(result.evidenceRoot, result.manifest.sources[0].blob_path)
    const exported = await readFile(blob, 'utf8')
    assert.match(exported, /captured/)
    assert.doesNotMatch(exported, /not captured/)
    assert.equal(result.manifest.sources[0].lifecycle, 'transaction-snapshot')
  } finally {
    await rm(root, { recursive: true, force: true })
  }
})

test('extract-sessions phase records the evidence manifest in handoff state', async () => {
  const root = await tempFolder('phase')
  try {
    const source = join(root, 'session.jsonl')
    const output = join(root, 'output')
    const statePath = join(output, '.handoff', 'state.json')
    await mkdir(join(output, '.handoff'), { recursive: true })
    await writeFile(source, '{"role":"user","content":"phase test"}\n')
    await writeState(statePath, {
      ...STATE_DEFAULT,
      project: 'phase-test', source_root: root, output_dir: output,
      session_capture: {
        providers: [],
        sources: [{ path: source, provider: 'custom', session_id: 'phase-1' }],
        capture_raw: true,
      },
    })
    await extractSessions(statePath)
    const state = await readState(statePath)
    assert.equal(state.sessions_found.length, 1)
    assert.ok(state.session_evidence_path)
    assert.ok(state.phases_completed.includes('extract-sessions'))
    assert.match(await readFile(join(output, '.handoff', 'tmp', 'phase-1_extract.txt'), 'utf8'), /phase test/)
    const manifest = JSON.parse(await readFile(state.session_evidence_path, 'utf8'))
    assert.equal(manifest.source_count, 1)
  } finally {
    await rm(root, { recursive: true, force: true })
  }
})
