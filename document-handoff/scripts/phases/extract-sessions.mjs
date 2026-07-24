import { mkdir, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { readState, writeState } from '../lib/state.mjs'
import { discoverSessions, extractSession, normalizeExtract } from '../lib/providers.mjs'
import { captureSessionEvidence } from '../lib/session-evidence.mjs'

function safeName(value) {
  return String(value).replace(/[^a-z0-9._-]/gi, '-').slice(0, 180)
}

export async function run(statePath) {
  const state = await readState(statePath)
  const { source_root: sourceRoot, output_dir: outputDir } = state
  const config = state.session_capture || { providers: ['claude-code', 'codex'], capture_raw: true }
  const sessions = await discoverSessions(sourceRoot, config)

  if (sessions.length === 0) {
    console.log('⚠️  No sessions found for this source root.')
    console.log('GATE: Set sessions_validated:true in state.json manually if you want to continue without sessions.')
    return
  }

  console.log('\nSessions and supporting sources found:')
  sessions.forEach((session, index) => {
    console.log(`  ${index + 1}. [${session.provider}] ${session.session_name} (${session.kind})`)
  })
  console.log('\nGATE: Review the sources above. Set sessions_validated:true in state.json to proceed.')

  const tmpDir = join(outputDir, '.handoff', 'tmp')
  await mkdir(tmpDir, { recursive: true })
  const nameCounts = new Map()
  for (const session of sessions.filter(item => item.kind !== 'supporting')) {
    const stem = safeName(session.session_name)
    nameCounts.set(stem, (nameCounts.get(stem) || 0) + 1)
  }
  for (const session of sessions) {
    if (session.kind === 'supporting') continue
    const messages = await extractSession(session)
    const normalized = normalizeExtract(messages, session.provider, session.session_name)
    const stem = safeName(session.session_name)
    const name = nameCounts.get(stem) > 1
      ? `${safeName(session.provider)}-${stem}_extract.txt`
      : `${stem}_extract.txt`
    await writeFile(join(tmpDir, name), normalized)
  }

  if (config.capture_raw !== false) {
    const evidence = await captureSessionEvidence(sessions, outputDir, config)
    state.session_evidence_path = evidence.manifestPath
    console.log(`✓ Captured ${evidence.manifest.source_count} evidence sources (${evidence.manifest.unique_blob_count} unique blobs)`)
  }

  state.sessions_found = sessions
  if (!state.phases_completed.includes('extract-sessions')) state.phases_completed.push('extract-sessions')
  await writeState(statePath, state)
  console.log(`✓ Extracted ${sessions.length} sessions/sources to ${tmpDir}`)
}
