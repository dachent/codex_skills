# Multi-harness session capture

`extract-sessions` discovers, normalizes, and optionally preserves raw session evidence. The default remains Claude Code plus Codex. A JSON configuration passed to `init --session-config` can enable Kimi Code, Hermes, or explicitly declared sources.

## Configuration

```json
{
  "providers": [
    "claude-code",
    { "type": "codex", "root": "C:\\Users\\me\\.codex" },
    { "type": "kimi-code", "root": "C:\\Users\\me\\.kimi-code\\sessions" },
    {
      "type": "hermes",
      "database": "C:\\Users\\me\\AppData\\Local\\hermes\\state.db",
      "session_ids": ["session-1"]
    }
  ],
  "sources": [
    {
      "provider": "another-harness",
      "kind": "jsonl",
      "path": "D:\\logs\\session.jsonl",
      "session_id": "session-2",
      "active": true
    },
    {
      "provider": "another-harness",
      "kind": "supporting",
      "path": "D:\\logs\\registry.json",
      "session_id": "registry"
    }
  ],
  "capture_raw": true,
  "privacy_classification": "not-requested"
}
```

Provider entries accept `session_ids` to select known sessions and `cwd_prefixes` to add project roots. Kimi entries also accept `workspace_ids`. Hermes entries may override `session_table`, `id_column`, and `cwd_column` for a compatible schema. Explicit `sources` are the extension point for any JSONL-producing harness without a built-in adapter.

## Evidence contract

Raw evidence is written below `.handoff/session-evidence/`:

- `manifest.json` records provider, session ID, source path, lifecycle, byte and record counts, SHA-256, and blob path.
- `checksums.sha256` provides an independently consumable checksum list.
- `blobs/sha256/` stores content-addressed evidence, deduplicating identical bytes.

Active JSONL files are captured as a newline-terminated prefix. The source prefix is hashed before and after copying, so appends are allowed while mutation of already observed records fails. Supporting files require matching full-file reads. Hermes data is selected inside a read transaction and exported as per-session JSONL; the entire database is not retained by default.

Publishing uses an atomic rename from a same-directory staging file. It does not require hard-link support, which makes the contract suitable for local disks and SMB/Azure Files destinations inherited from the handoff output directory.

The tool does not impose ACL or privacy gates. Those remain destination-policy decisions. Set `privacy_classification` when a workflow has performed a separate classification review.
