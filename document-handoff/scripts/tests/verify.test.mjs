import { test } from 'node:test'
import assert from 'node:assert/strict'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { writeFile, readFile, mkdir, rm } from 'node:fs/promises'
import { writeState, readState, STATE_DEFAULT } from '../lib/state.mjs'
import { run as verify } from '../phases/verify.mjs'
import { renderMemo } from '../lib/html.mjs'

const REQUIRED_IDS = ['bootstrap','executive-summary','deliverables','current-state',
  'technical-decisions','challenges-blockers','next-steps','context-sources',
  'open-questions','dependencies','environment','testing','architecture','data-flow','changelog']

async function makeVerifyFixture(sections, extraState = {}) {
  const out = join(tmpdir(), `vtest-${process.hrtime.bigint()}`)
  await mkdir(join(out, '.handoff'), { recursive: true })
  const state = { ...STATE_DEFAULT, project: 'vtest', output_dir: out, risk_flags_count: 0, ...extraState }
  const html = await renderMemo(sections, state)
  const memoPath = join(out, 'vtest-memo.html')
  await writeFile(memoPath, html)
  const agentCtxPath = join(out, 'vtest-agent-context.json')
  await writeFile(agentCtxPath, '{}')
  const citationPath = join(out, '.handoff', 'vtest-citation-index.json')
  await writeFile(citationPath, '{}')
  state.memo_path = memoPath
  state.agent_context_path = agentCtxPath
  const statePath = join(out, '.handoff', 'state.json')
  await writeState(statePath, state)
  return { out, statePath }
}

function allSections() {
  return Object.fromEntries(REQUIRED_IDS.map(id => [id, { id, title: id, content: `<p>${id} content</p>`, catalog_ids_referenced: [] }]))
}

const mockAgent = async () => ({ passed: true, issues: [] })

test('verify: passes when all required sections present', async () => {
  const { out, statePath } = await makeVerifyFixture(allSections())
  await verify(statePath, mockAgent)
  const state = await readState(statePath)
  assert.equal(state.verified, true)
  await rm(out, { recursive: true, force: true })
})

test('verify: fails when a required section is missing', async () => {
  const sections = allSections()
  delete sections['bootstrap']
  const { out, statePath } = await makeVerifyFixture(sections)
  await verify(statePath, mockAgent)
  const state = await readState(statePath)
  assert.equal(state.verified, false)
  await rm(out, { recursive: true, force: true })
})

test('verify: fails when privacy-security missing but risk_flags_count > 0', async () => {
  const { out, statePath } = await makeVerifyFixture(allSections(), { risk_flags_count: 3 })
  await verify(statePath, mockAgent)
  const state = await readState(statePath)
  assert.equal(state.verified, false)
  await rm(out, { recursive: true, force: true })
})

test('verify: writes verification.json', async () => {
  const { out, statePath } = await makeVerifyFixture(allSections())
  await verify(statePath, mockAgent)
  const state = await readState(statePath)
  assert.ok(state.verification_path)
  const v = JSON.parse(await readFile(state.verification_path, 'utf8'))
  assert.ok('overall' in v)
  await rm(out, { recursive: true, force: true })
})

test('verify: fails when captured session evidence no longer matches its manifest', async () => {
  const { out, statePath } = await makeVerifyFixture(allSections())
  const evidenceRoot = join(out, '.handoff', 'session-evidence')
  const blobPath = join(evidenceRoot, 'blobs', 'sha256', '00', 'bad.jsonl')
  await mkdir(join(evidenceRoot, 'blobs', 'sha256', '00'), { recursive: true })
  await writeFile(blobPath, 'changed\n')
  const manifestPath = join(evidenceRoot, 'manifest.json')
  await writeFile(manifestPath, JSON.stringify({
    schema_version: 1,
    sources: [{ blob_path: 'blobs/sha256/00/bad.jsonl', bytes: 8, sha256: '0'.repeat(64) }],
  }))
  const state = await readState(statePath)
  state.session_evidence_path = manifestPath
  await writeState(statePath, state)

  await verify(statePath, mockAgent)
  const verifiedState = await readState(statePath)
  const result = JSON.parse(await readFile(verifiedState.verification_path, 'utf8'))
  assert.equal(verifiedState.verified, false)
  assert.equal(result.additional.session_evidence_integrity, false)
  await rm(out, { recursive: true, force: true })
})
