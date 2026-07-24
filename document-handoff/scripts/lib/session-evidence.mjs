import { createHash, randomBytes } from 'node:crypto'
import { createReadStream, createWriteStream } from 'node:fs'
import { mkdir, open, readFile, rename, rm, stat, writeFile } from 'node:fs/promises'
import { basename, dirname, extname, join, relative } from 'node:path'
import { pipeline } from 'node:stream/promises'
import { exportHermesRows } from './providers.mjs'

const HASH_CHUNK = 1024 * 1024
const MAX_ATTEMPTS = 3

function digestBuffer(value) {
  return createHash('sha256').update(value).digest('hex')
}

async function hashPrefix(path, length) {
  const hash = createHash('sha256')
  if (length === 0) return hash.digest('hex')
  const stream = createReadStream(path, { start: 0, end: length - 1 })
  for await (const chunk of stream) hash.update(chunk)
  return hash.digest('hex')
}

async function countNewlines(path) {
  let count = 0
  for await (const chunk of createReadStream(path)) {
    for (let index = 0; index < chunk.length; index++) {
      if (chunk[index] === 0x0a) count++
    }
  }
  return count
}

async function completeJsonlCutoff(path, length) {
  if (length === 0) return 0
  const handle = await open(path, 'r')
  try {
    let cursor = length
    while (cursor > 0) {
      const size = Math.min(HASH_CHUNK, cursor)
      const start = cursor - size
      const buffer = Buffer.allocUnsafe(size)
      const { bytesRead } = await handle.read(buffer, 0, size, start)
      for (let index = bytesRead - 1; index >= 0; index--) {
        if (buffer[index] === 0x0a) return start + index + 1
      }
      cursor = start
    }
    return 0
  } finally {
    await handle.close()
  }
}

async function tempPath(root) {
  const dir = join(root, '.staging')
  await mkdir(dir, { recursive: true })
  return join(dir, `${randomBytes(12).toString('hex')}.tmp`)
}

async function copyPrefix(source, length, destination) {
  if (length === 0) {
    await writeFile(destination, Buffer.alloc(0))
    return
  }
  await pipeline(
    createReadStream(source, { start: 0, end: length - 1 }),
    createWriteStream(destination, { flags: 'wx' }),
  )
}

async function stableJsonlSnapshot(source, evidenceRoot) {
  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    const before = await stat(source)
    const cutoff = await completeJsonlCutoff(source, before.size)
    const firstHash = await hashPrefix(source, cutoff)
    const temp = await tempPath(evidenceRoot)
    try {
      await copyPrefix(source, cutoff, temp)
      const [storedHash, secondHash] = await Promise.all([
        hashPrefix(temp, cutoff),
        hashPrefix(source, cutoff),
      ])
      if (firstHash === storedHash && storedHash === secondHash) {
        return {
          temp, sha256: storedHash, bytes: cutoff,
          records: await countNewlines(temp),
          observed_bytes: before.size,
          complete_record_prefix: cutoff < before.size,
        }
      }
    } catch (error) {
      await rm(temp, { force: true })
      if (attempt === MAX_ATTEMPTS) throw error
      continue
    }
    await rm(temp, { force: true })
  }
  throw new Error(`session source changed during capture: ${basename(source)}`)
}

async function stableFileSnapshot(source, evidenceRoot) {
  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    const before = await stat(source)
    const firstHash = await hashPrefix(source, before.size)
    const temp = await tempPath(evidenceRoot)
    try {
      await copyPrefix(source, before.size, temp)
      const after = await stat(source)
      const [storedHash, secondHash] = await Promise.all([
        hashPrefix(temp, before.size),
        hashPrefix(source, before.size),
      ])
      if (before.size === after.size && firstHash === storedHash && storedHash === secondHash) {
        return { temp, sha256: storedHash, bytes: before.size, records: null }
      }
    } catch (error) {
      await rm(temp, { force: true })
      if (attempt === MAX_ATTEMPTS) throw error
      continue
    }
    await rm(temp, { force: true })
  }
  throw new Error(`supporting source changed during capture: ${basename(source)}`)
}

async function publishBlob(snapshot, evidenceRoot, extension = '') {
  const folder = join(evidenceRoot, 'blobs', 'sha256', snapshot.sha256.slice(0, 2))
  await mkdir(folder, { recursive: true })
  const target = join(folder, `${snapshot.sha256}${extension}`)
  try {
    // Rename works on SMB targets where hard links are unavailable.
    await rename(snapshot.temp, target)
  } catch (error) {
    if (!['EEXIST', 'EPERM'].includes(error.code)) throw error
    const existing = await hashPrefix(target, snapshot.bytes)
    if (existing !== snapshot.sha256 || (await stat(target)).size !== snapshot.bytes) throw error
    await rm(snapshot.temp, { force: true })
  }
  return target
}

function publicSession(session) {
  return {
    provider: session.provider,
    adapter: session.adapter,
    session_id: session.session_id,
    session_name: session.session_name,
    source_kind: session.kind,
    source_path: session.path,
    active: Boolean(session.active),
    metadata: session.metadata || {},
  }
}

async function captureJsonl(session, evidenceRoot) {
  const snapshot = await stableJsonlSnapshot(session.path, evidenceRoot)
  const blob = await publishBlob(snapshot, evidenceRoot, '.jsonl')
  return {
    ...publicSession(session),
    sha256: snapshot.sha256,
    bytes: snapshot.bytes,
    observed_bytes: snapshot.observed_bytes,
    records: snapshot.records,
    complete_record_prefix: snapshot.complete_record_prefix,
    lifecycle: session.active ? 'active-prefix' : 'snapshot',
    blob_path: relative(evidenceRoot, blob).replace(/\\/g, '/'),
  }
}

async function captureSupporting(session, evidenceRoot) {
  const snapshot = await stableFileSnapshot(session.path, evidenceRoot)
  const blob = await publishBlob(snapshot, evidenceRoot, extname(session.path))
  return {
    ...publicSession(session),
    sha256: snapshot.sha256,
    bytes: snapshot.bytes,
    records: null,
    lifecycle: 'snapshot',
    blob_path: relative(evidenceRoot, blob).replace(/\\/g, '/'),
  }
}

async function captureHermes(session, evidenceRoot) {
  const rows = await exportHermesRows(session)
  const data = Buffer.from(rows.map(value => JSON.stringify(value)).join('\n') + (rows.length ? '\n' : ''), 'utf8')
  const snapshot = {
    temp: await tempPath(evidenceRoot),
    sha256: digestBuffer(data),
    bytes: data.length,
    records: rows.length,
  }
  await writeFile(snapshot.temp, data, { flag: 'wx' })
  const blob = await publishBlob(snapshot, evidenceRoot, '.jsonl')
  return {
    ...publicSession(session),
    sha256: snapshot.sha256,
    bytes: snapshot.bytes,
    records: snapshot.records,
    lifecycle: session.active ? 'active-transaction-snapshot' : 'transaction-snapshot',
    blob_path: relative(evidenceRoot, blob).replace(/\\/g, '/'),
  }
}

async function atomicJson(path, value) {
  await mkdir(dirname(path), { recursive: true })
  const temp = `${path}.${randomBytes(8).toString('hex')}.tmp`
  await writeFile(temp, `${JSON.stringify(value, null, 2)}\n`, 'utf8')
  await rename(temp, path)
}

export async function captureSessionEvidence(sessions, outputDir, options = {}) {
  const evidenceRoot = join(outputDir, '.handoff', 'session-evidence')
  await mkdir(evidenceRoot, { recursive: true })
  const sources = []
  for (const session of sessions) {
    if (session.kind === 'sqlite-session') sources.push(await captureHermes(session, evidenceRoot))
    else if (session.kind === 'supporting') sources.push(await captureSupporting(session, evidenceRoot))
    else sources.push(await captureJsonl(session, evidenceRoot))
  }

  const manifest = {
    schema_version: 1,
    created_at: new Date().toISOString(),
    capture_policy: {
      jsonl: 'stable newline-terminated prefix',
      supporting: 'stable two-read snapshot',
      sqlite: 'read transaction with per-session JSONL export',
      content_addressed: true,
      privacy_classification: options.privacy_classification || 'not-requested',
    },
    source_count: sources.length,
    unique_blob_count: new Set(sources.map(source => source.blob_path)).size,
    sources,
  }
  const manifestPath = join(evidenceRoot, 'manifest.json')
  await atomicJson(manifestPath, manifest)
  const checksums = [...new Map(sources.map(source => [source.blob_path, source.sha256])).entries()]
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([path, hash]) => `${hash}  ${path}`)
    .join('\n')
  await writeFile(join(evidenceRoot, 'checksums.sha256'), `${checksums}${checksums ? '\n' : ''}`, 'utf8')
  await rm(join(evidenceRoot, '.staging'), { recursive: true, force: true })
  return { evidenceRoot, manifestPath, manifest }
}

export async function verifySessionEvidence(manifestPath) {
  const manifest = JSON.parse(await readFile(manifestPath, 'utf8'))
  const root = dirname(manifestPath)
  const failures = []
  for (const source of manifest.sources || []) {
    const path = join(root, source.blob_path)
    try {
      const info = await stat(path)
      const hash = await hashPrefix(path, info.size)
      if (info.size !== source.bytes || hash !== source.sha256) failures.push(source.blob_path)
    } catch {
      failures.push(source.blob_path)
    }
  }
  return { passed: failures.length === 0, failures, source_count: manifest.sources?.length || 0 }
}
