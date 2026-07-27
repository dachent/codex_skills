# Archived packages

This namespace retains historical skill packages that are no longer supported or installable. Archived packages remain in `skills-manifest.json` with `status: archived` for provenance and repository integrity, but generators exclude them from the skill catalog, installation inventory, platform matrix, validation summary, and CI skill matrix.

Do not copy or install packages from this directory. Their internal `SKILL.md`, agent metadata, code, and tests are preserved only so prior design decisions remain reviewable.

## Archived packages

- [`agent-project-orchestrator`](./agent-project-orchestrator): archived because evaluation found no demonstrated value over the original project scaffold for Claude Code and no value over native planning and execution controls for Codex GPT-5.6 Sol. Its proposed durable control plane was not implemented, and its draft parser assumptions did not match observed scaffold use. See the package README and [issue #62](https://github.com/dachent/skills/issues/62).
