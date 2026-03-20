# codex_skills

Windows-specific Codex ports of Office skills sourced from the Anthropic skills repository.

These ports are not just copied prompts. Each one rewrites the default execution path around local Microsoft Office desktop automation on Windows, so the behavior depends on Word, PowerPoint, or Excel being installed and callable through COM.

## Provenance

| Skill | Upstream repo | Source folder | Source branch | Port depth | Why Windows-specific |
| --- | --- | --- | --- | --- | --- |
| `docx-win` | `https://github.com/anthropics/skills` | `skills/docx` | `main` | Light port | Uses local Microsoft Word COM automation and PowerShell wrappers instead of the upstream LibreOffice-first path. |

## Skills

### `docx-win`

`docx-win` is a light port of Anthropic's `skills/docx` skill into a Codex-friendly Windows workflow.

The upstream skill centers on OOXML unpacking/editing plus LibreOffice-based conversion and change-acceptance helpers. This port keeps the document workflow but changes the default execution path to Microsoft Word COM and PowerShell so Word itself handles conversion, tracked changes, comments, fields, pagination, and PDF export.

This skill is Windows-specific because those guarantees depend on a local Microsoft Word desktop install and COM automation rather than a cross-platform LibreOffice pipeline.
