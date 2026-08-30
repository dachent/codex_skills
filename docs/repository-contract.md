# Repository integration contract

`skills-manifest.json` is the operational authority for active and archived skill inventory, lifecycle, catalog grouping, ownership, platform and agent support, packaging, shared runtimes, validation, and generated mirrors. `.provenance/source-registry.json` is the authority for external source identity, immutable source revisions, license review, alignment metadata, and distribution boundaries. Non-skill scaffolds are governed by their namespace contract and package README files instead.

## Repository namespaces

- Active skills use one canonical top-level directory and appear in `skills-manifest.json`.
- `scaffolds/<name>/` contains reusable instruction documents, not skills. Every package must include `README.md` and must not contain `SKILL.md` or `agents/openai.yaml`. Scaffolds never appear in the manifest or generated installation inventory.
- `archive/<name>/` contains historical, non-installable skill packages. Every archived package remains registered in the manifest with `status: archived` and a path of exactly `archive/<name>`.
- `.shared/`, `.generated/`, `.codex/`, `.github/`, `.provenance/`, `.upstream/`, `docs/`, and `tools/` contain repository infrastructure rather than independently installable skills.

## Canonical sources

Each active skill has one canonical top-level directory. Archived skill packages live one level under `archive/`. `.claude/skills` may contain only mirrors declared in `generated_mirrors` and produced by `tools/generate_repository_artifacts.py`. Undeclared files there are drift.

For externally derived skills, provenance records are keyed by skill name and reference a registered source. The provenance registry owns repository identity, source path, immutable revision, license state, distribution, and alignment metadata. The manifest retains source classification plus materialized source fields for catalog generation and compatibility; `tools/validate_provenance.py` requires those fields to agree with the registry exactly.

## Catalog groups and lifecycle

Catalog groups explain why a skill belongs in the repository, not merely its technical domain. Every active skill belongs to exactly one ordered group declared in `policy.catalog_groups`. The README catalog is generated from those declarations.

`policy.installable_statuses` controls the generated installation inventory. Supported and experimental skills are installable by default. Deprecated skills remain visible in the catalog and platform/agent matrix for compatibility, but they are not offered in the default installation inventory. Archived skills are excluded from active generated surfaces.

## Generated artifacts

Run `python .\tools\generate_repository_artifacts.py` to generate the README catalog, installation inventory, platform/agent matrix, validation summary, declared agent mirrors, and `.generated/agent-mirrors.json`.

Run `python .\tools\generate_repository_artifacts.py --check` to fail on stale marked README sections, stale mirrors, stale hashes, or undeclared files under `.claude/skills`.

Generated README regions carry explicit markers and notices. Prose outside those markers remains hand-maintained.

## Agent mirrors

A mirror is allowed only when an agent needs a material packaging difference. Each declaration identifies a canonical source, a destination under `.claude/skills`, and an explicit transformation. `copy-with-generated-notice` preserves canonical content while inserting a generated-file notice after YAML front matter.

`.generated/agent-mirrors.json` records source and destination SHA-256 values. No mirrors are currently declared; compatible agents should load canonical top-level skills directly.

## Active-skill package

Every supported skill must provide a top-level canonical directory, matching `SKILL.md`, `agents/openai.yaml`, catalog group, source classification, owner, platforms, agents, review date, and validation declarations. A provenance file is required when declared by the manifest.

Every active skill must also have a record in `.provenance/source-registry.json`. External records must reference an immutable registered source revision. A restricted source may remain unresolved as to original ownership or license, but the known imported snapshot itself must still be pinned rather than represented as a fictitious `revision-unresolved` state.

Archived packages may retain package metadata for historical review but are not installable.

## Licensing

The root MIT `LICENSE` governs `repo-owned-original` material. It does not relicense external derivatives or unresolved imports. External distribution boundaries come from `.provenance/source-registry.json` and are explained in `docs/licensing-and-redistribution.md`.

## Pull requests and CI

A skill-changing pull request updates the manifest, canonical skill package, provenance and tests as applicable, then regenerates repository artifacts. Source or licensing changes update the provenance registry first and synchronize the manifest projection. Scaffold and archive changes update their namespace documentation and applicable manifest records.

Repository-integrity CI runs manifest validation, namespace checks, generator check mode, generator tests, metadata validation, provenance cross-checks, and Codex hook validation. Adding, deprecating, archiving, or removing a skill cannot silently leave installation surfaces, provenance, or generated documentation inconsistent.

Office COM validation remains environment-dependent and runs through the controlled self-hosted workflow.
