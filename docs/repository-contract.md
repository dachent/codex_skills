# Repository integration contract

`skills-manifest.json` is the operational source of truth for active and archived skills, catalog grouping, ownership, platform and agent support, provenance classification, packaging, shared runtimes, validation, and generated mirrors. Non-skill scaffolds are governed by their namespace contract and package README files instead.

## Repository namespaces

- Active skills use one canonical top-level directory and appear in `skills-manifest.json`.
- `scaffolds/<name>/` contains reusable instruction documents, not skills. Every package must include `README.md` and must not contain `SKILL.md` or `agents/openai.yaml`. Scaffolds never appear in the manifest or generated installation inventory.
- `archive/<name>/` contains historical, non-installable skill packages. Every archived package remains registered in the manifest with `status: archived` and a path of exactly `archive/<name>`.
- `.shared/`, `.generated/`, `.codex/`, `.github/`, `.provenance/`, `.upstream/`, `docs/`, and `tools/` contain repository infrastructure rather than independently installable skills.

## Canonical sources

Each active skill has one canonical top-level directory. Archived skill packages live one level under `archive/`. `.claude/skills` may contain only mirrors declared in `generated_mirrors` and produced by `tools/generate_repository_artifacts.py`. Undeclared files there are drift.

## Catalog groups

Catalog groups explain why a skill belongs in the repository, not merely its technical domain. Every active skill belongs to exactly one ordered group declared in `policy.catalog_groups`. The README catalog is generated from those declarations.

## Generated artifacts

Run `python .\tools\generate_repository_artifacts.py` to generate the README catalog, installation inventory, platform/agent matrix, validation summary, declared agent mirrors, and `.generated/agent-mirrors.json`.

Run `python .\tools\generate_repository_artifacts.py --check` to fail on stale marked README sections, stale mirrors, stale hashes, or undeclared files under `.claude/skills`.

Generated README regions carry explicit markers and notices. Prose outside those markers remains hand-maintained.

## Agent mirrors

A mirror is allowed only when an agent needs a material packaging difference. Each declaration identifies a canonical source, a destination under `.claude/skills`, and an explicit transformation. `copy-with-generated-notice` preserves canonical content while inserting a generated-file notice after YAML front matter.

`.generated/agent-mirrors.json` records source and destination SHA-256 values. No mirrors are currently declared; compatible agents should load canonical top-level skills directly.

## Active-skill package

Every supported skill must provide a top-level canonical directory, matching `SKILL.md`, `agents/openai.yaml`, catalog group, source classification, owner, platforms, agents, review date, and validation declarations. A provenance file is required when declared by the manifest.

Archived packages may retain those files for historical review, but they are not active packages and are excluded from all generated installation and support surfaces.

External adaptations should identify an immutable upstream revision. When prior repository history did not preserve one, the source must explicitly declare `provenance_status: revision-unresolved`; this is a visible debt, not a substitute for provenance completion.

## Pull requests and CI

A skill-changing pull request updates the manifest, canonical skill package, provenance and tests as applicable, then regenerates repository artifacts. Scaffold and archive changes update their namespace documentation and applicable manifest records. Repository-integrity CI runs manifest validation, namespace checks, generator check mode, generator tests, metadata validation, provenance checks, and Codex hook validation. Adding, archiving, or removing a skill cannot leave the README stale, and canonical and generated definitions cannot silently diverge.

Office COM validation remains environment-dependent and runs through the controlled self-hosted workflow.
