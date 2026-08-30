# Agent skills, scaffolds, and repository tooling

[![Validate](https://github.com/dachent/skills/actions/workflows/validate.yml/badge.svg?branch=main)](https://github.com/dachent/skills/actions/workflows/validate.yml)

A curated agent-operations repository built for five distinct purposes:

1. **Native application skill extensions** — forks and adaptations that replace generic file transformation with native Microsoft Office automation, including `docx-win`, `pptx-win`, and `xlsx-win`.
2. **Claude and agent workflow ports for Codex** — established workflows such as Grill Me, Handoff, Deep Planning, and Ultraplan adapted to Codex conventions, tools, permissions, and durable artifact contracts.
3. **Repository-owned specialist skills** — original or locally imported implementations, including Code Mapper and Document Handoff, whose current behavior is maintained here.
4. **A Codex design extension pack** — coordinated visual skills that extend Codex design work with shared browser rendering, screenshot evidence, accessibility review, and artifact QA.
5. **Non-installable project scaffolds and historical packages** — reusable instruction documents under `scaffolds/` and retired implementations under `archive/`, kept explicitly separate from the installable skill inventory.

This provenance-and-purpose grouping explains **why content belongs here**. Platform, runtime, agent support, and validation remain separate attributes for installable skills and are generated from [`skills-manifest.json`](./skills-manifest.json).

## Repository model

`skills-manifest.json` is the operational authority for skill inventory, lifecycle, grouping, ownership, platform and agent support, packaging, shared runtimes, validation, and generated mirrors. [`.provenance/source-registry.json`](./.provenance/source-registry.json) is the authority for external source identity, immutable source revisions, license review, alignment metadata, and distribution boundaries. Source facts duplicated into the manifest are a CI-validated projection of that registry, not an independent authority. Each active skill has one canonical top-level directory. [`scaffolds/`](./scaffolds) contains documented non-skills; [`archive/`](./archive) contains non-installable historical packages. See [`docs/repository-contract.md`](./docs/repository-contract.md).

## Skill catalog

<!-- BEGIN GENERATED: skill-catalog -->
<!-- Generated section: skill catalog. Source: skills-manifest.json. Generator: tools/generate_repository_artifacts.py. Do not edit by hand. -->

### Native application skill extensions

Forked or adapted skills that replace generic file handling with native desktop application automation.

| Skill | Purpose | Provenance |
| --- | --- | --- |
| [`docx-win`](./docx-win) | native microsoft word automation for windows .docx workflows. use when chatgpt or codex is running on windows with microsoft word installed and needs to create, edit, review, convert, or verify word documents through word com automation. trigger for .docx or .doc requests, professional word deliverables, no-template document polish, tracked changes, comments, find and replace, table of contents, headers and footers, page numbering, layout-sensitive edits, or exporting a word document to pdf for review. prefer this skill over libreoffice-based document workflows when word is available. | heavy adaptation; `anthropics/skills` |
| [`pptx-win`](./pptx-win) | windows powerpoint automation and no-template deck design support for .pptx files in codex app or other local windows environments with microsoft 365 installed. use when chatgpt needs to open, inspect, edit, create, render, export, or qa powerpoint presentations on windows, especially new decks without a template, existing deck edits, speaker notes, slide images, pdf export, placeholder replacement, visual QA, or smoke testing via native powerpoint com automation. prefer this skill over libreoffice-based flows when powerpoint desktop is available. | heavy adaptation; `anthropics/skills` |
| [`xlsx-win`](./xlsx-win) | native microsoft excel automation for windows .xlsx/.xlsm/.xls/.csv/.tsv workflows. use when chatgpt, codex, or claude code is running on windows with microsoft 365 excel installed and needs to refresh, recalculate, validate, or route a workbook through excel com automation. trigger for workbook connections, cached values, pivottables, calculation correctness, refreshing an existing power query connection, data model-aware routing, chart-ready data, no-template spreadsheet deliverables, or excel environment self-test. does not support authoring or editing power query m code, or macro execution -- see "known gaps" in SKILL.md. do not use for cloud execution, non-windows environments, google sheets api workflows, or machines without excel desktop installed. | heavy adaptation; `anthropics/skills` |

### Claude and agent workflow ports for Codex

External agent workflows ported or substantially adapted for Codex conventions, tools, and artifact contracts.

| Skill | Purpose | Provenance |
| --- | --- | --- |
| [`adversarial-plan-review-codex`](./adversarial-plan-review-codex) | Use when a plan needs hostile review before execution, especially high-risk coding, business deliverables, migrations, no-git changes, weak validation, stale assumptions, rollback gaps, or plans that must be safe for another agent to execute. | medium adaptation; `dachent/cdc05151d047708c290bd4da0aaeed96` |
| [`deep-planning-codex`](./deep-planning-codex) | Deprecated for GPT-5.6/Sol: use native Codex Plan Mode, with focused verification or adversarial review skills when needed. | heavy adaptation; `dachent/cdc05151d047708c290bd4da0aaeed96` |
| [`grill-me-codex`](./grill-me-codex) | Use when the user says grill me, wants to stress-test a plan or design, needs a rigorous interview before committing to a decision, or asks for adversarial product, architecture, or implementation questions. | medium adaptation; `mattpocock/skills` |
| [`grill-with-docs-codex`](./grill-with-docs-codex) | Use when the user wants to stress-test a plan against project terminology, domain language, CONTEXT.md, ADRs, existing docs, or code-backed architectural decisions. | medium adaptation; `mattpocock/skills` |
| [`handoff-codex`](./handoff-codex) | Use when the user asks for a handoff, session summary, context packet, continuation note, or wants another agent or future session to pick up the current work. | medium adaptation; `mattpocock/skills` |
| [`repo-map-codex`](./repo-map-codex) | Deprecated for GPT-5.6/Sol: native Plan Mode performs read-first repository grounding and evidence collection. | medium adaptation; `dachent/cdc05151d047708c290bd4da0aaeed96` |
| [`ultraplan-codex`](./ultraplan-codex) | Deprecated for GPT-5.6/Sol: use native Codex Plan Mode for grounded, decision-complete implementation plans. | heavy adaptation; `6missedcalls/ultraplan` |
| [`verification-plan-codex`](./verification-plan-codex) | Use when a plan needs proof criteria before execution, especially coding changes, business deliverables, mixed business-coding projects, acceptance criteria, rollback triggers, manual checks, or final validation design. | medium adaptation; `dachent/cdc05151d047708c290bd4da0aaeed96` |

### Repository-owned specialist skills

Original or locally imported workflows whose current implementation is maintained in this repository.

| Skill | Purpose | Provenance |
| --- | --- | --- |
| [`code-intelligence`](./code-intelligence) | Claude Code-only explicit provider router for assessing installed Graphify, code-mapper, or selective CodeQL routes. Not for Codex GPT-5.6 Sol, whose native inspection, Plan Mode, explorer work, and subagents cover this function without an added analyzer. | repo owned original |
| [`code-mapper-skill`](./code-mapper-skill) | Generate deterministic Python import, reference, artifact, contract, catalog, OpenLineage, and explicitly authorized local CodeQL maps. Use for explicit blast-radius, caller, contract, lineage, or local value/taint evidence requests against an approved local worktree. | repo owned original |
| [`document-handoff`](./document-handoff) | Create a comprehensive cross-harness project archive and handoff package with content-addressed Codex, Claude Code, Kimi Code, Hermes, or declared session evidence. Use for milestone archives and cold-start continuation; distinct from a lightweight session handoff. | local source import |

### Codex design extension pack

A coordinated set of visual design skills built around shared browser rendering, evidence, and QA infrastructure.

| Skill | Purpose | Provenance |
| --- | --- | --- |
| [`canvas-design-codex`](./canvas-design-codex) | Use when Codex needs to build, revise, debug, or verify canvas, SVG, WebGL, Three.js, custom charting, diagram, game, generative visual, or pixel-based browser artwork where screenshot, pixel, animation, or image bounds evidence matters. | repo owned original |
| [`frontend-design-codex`](./frontend-design-codex) | Use when Codex is building, revising, or reviewing a frontend UI, visual system, web app screen, responsive layout, component surface, dashboard view, or no-template product interface that needs semantic design tokens, visual polish, accessibility, screenshots, or browser QA. | repo owned original |
| [`web-artifacts-builder-codex`](./web-artifacts-builder-codex) | Use when Codex needs to create, run, package, revise, or verify a standalone web artifact with an explicit reviewer runtime, relative-asset checks, screenshots, console and request capture, rerun instructions, and evidence packaging. | repo owned original |
<!-- END GENERATED: skill-catalog -->

## Installation inventory

<!-- BEGIN GENERATED: installation-inventory -->
<!-- Generated section: installation inventory. Source: skills-manifest.json. Generator: tools/generate_repository_artifacts.py. Do not edit by hand. -->

Only supported or experimental entries below are installable. Deprecated skills remain cataloged for compatibility but are excluded from this inventory; `scaffolds/` and `archive/` are also excluded. Copy only the skills required by the target agent, plus listed shared components.

| Skill | Canonical directory | Shared components |
| --- | --- | --- |
| `adversarial-plan-review-codex` | [`adversarial-plan-review-codex`](./adversarial-plan-review-codex) | — |
| `canvas-design-codex` | [`canvas-design-codex`](./canvas-design-codex) | `.shared/visual-runtime` |
| `code-intelligence` | [`code-intelligence`](./code-intelligence) | — |
| `code-mapper-skill` | [`code-mapper-skill`](./code-mapper-skill) | — |
| `document-handoff` | [`document-handoff`](./document-handoff) | — |
| `docx-win` | [`docx-win`](./docx-win) | `.shared/office-com` |
| `frontend-design-codex` | [`frontend-design-codex`](./frontend-design-codex) | `.shared/visual-runtime` |
| `grill-me-codex` | [`grill-me-codex`](./grill-me-codex) | — |
| `grill-with-docs-codex` | [`grill-with-docs-codex`](./grill-with-docs-codex) | — |
| `handoff-codex` | [`handoff-codex`](./handoff-codex) | — |
| `pptx-win` | [`pptx-win`](./pptx-win) | `.shared/office-com` |
| `verification-plan-codex` | [`verification-plan-codex`](./verification-plan-codex) | — |
| `web-artifacts-builder-codex` | [`web-artifacts-builder-codex`](./web-artifacts-builder-codex) | `.shared/visual-runtime` |
| `xlsx-win` | [`xlsx-win`](./xlsx-win) | `.shared/office-com` |
<!-- END GENERATED: installation-inventory -->

## Platform and agent matrix

<!-- BEGIN GENERATED: platform-agent-matrix -->
<!-- Generated section: platform and agent matrix. Source: skills-manifest.json. Generator: tools/generate_repository_artifacts.py. Do not edit by hand. -->

| Skill | Platforms | Agents | Status |
| --- | --- | --- | --- |
| [`adversarial-plan-review-codex`](./adversarial-plan-review-codex) | `cross-platform` | `codex` | `supported` |
| [`canvas-design-codex`](./canvas-design-codex) | `cross-platform` | `codex` | `supported` |
| [`code-intelligence`](./code-intelligence) | `cross-platform` | `claude-code` | `supported` |
| [`code-mapper-skill`](./code-mapper-skill) | `cross-platform` | `codex`, `claude-code` | `supported` |
| [`deep-planning-codex`](./deep-planning-codex) | `cross-platform` | `codex` | `deprecated` |
| [`document-handoff`](./document-handoff) | `cross-platform` | `codex`, `claude-code` | `supported` |
| [`docx-win`](./docx-win) | `windows` | `codex`, `claude-code` | `supported` |
| [`frontend-design-codex`](./frontend-design-codex) | `cross-platform` | `codex` | `supported` |
| [`grill-me-codex`](./grill-me-codex) | `cross-platform` | `codex` | `supported` |
| [`grill-with-docs-codex`](./grill-with-docs-codex) | `cross-platform` | `codex` | `supported` |
| [`handoff-codex`](./handoff-codex) | `cross-platform` | `codex` | `supported` |
| [`pptx-win`](./pptx-win) | `windows` | `codex`, `claude-code` | `supported` |
| [`repo-map-codex`](./repo-map-codex) | `cross-platform` | `codex` | `deprecated` |
| [`ultraplan-codex`](./ultraplan-codex) | `cross-platform` | `codex` | `deprecated` |
| [`verification-plan-codex`](./verification-plan-codex) | `cross-platform` | `codex` | `supported` |
| [`web-artifacts-builder-codex`](./web-artifacts-builder-codex) | `cross-platform` | `codex` | `supported` |
| [`xlsx-win`](./xlsx-win) | `windows` | `codex`, `claude-code` | `supported` |
<!-- END GENERATED: platform-agent-matrix -->

## Prerequisites

Common repository tooling:

- Python 3.12 for validation and bundled helpers;
- Node.js 20 and npm for the visual runtime and Document Handoff tests;
- PowerShell 7 for repository and Office automation validation;
- Codex or Claude Code configured to load local skills.

Office skills require Windows, an interactive signed-in desktop session, and the corresponding Microsoft 365 desktop application. Code Mapper uses `grimp` and `jedi`; Git URL targets also require Git. Planning skills require permission to write their declared planning artifacts.

## Installation

### Codex

1. Clone the repository.
2. Select only skills listed in the generated installation inventory; never install from `scaffolds/` or `archive/`.
3. Copy the selected canonical top-level skill directories into the Codex skills directory, commonly `%USERPROFILE%\.codex\skills\`.
4. Copy shared components shown in the generated installation inventory.
5. Keep directory names unchanged.
6. Run the relevant validation before relying on the skill on a new machine.

### Claude Code

Load compatible canonical top-level skill directories from the generated installation inventory directly. Do not install `scaffolds/` or `archive/`. `.claude/skills` is reserved for generated, manifest-declared mirrors only; it is not a second source tree.

## Validation summary

<!-- BEGIN GENERATED: validation-summary -->
<!-- Generated section: validation summary. Source: skills-manifest.json. Generator: tools/generate_repository_artifacts.py. Do not edit by hand. -->

| Skill | Hosted commands | Environment-dependent commands |
| --- | ---: | ---: |
| `adversarial-plan-review-codex` | 1 | 0 |
| `canvas-design-codex` | 1 | 0 |
| `code-intelligence` | 1 | 0 |
| `code-mapper-skill` | 1 | 0 |
| `deep-planning-codex` | 1 | 0 |
| `document-handoff` | 1 | 0 |
| `docx-win` | 1 | 1 |
| `frontend-design-codex` | 1 | 0 |
| `grill-me-codex` | 0 | 0 |
| `grill-with-docs-codex` | 0 | 0 |
| `handoff-codex` | 0 | 0 |
| `pptx-win` | 1 | 1 |
| `repo-map-codex` | 1 | 0 |
| `ultraplan-codex` | 1 | 0 |
| `verification-plan-codex` | 1 | 0 |
| `web-artifacts-builder-codex` | 1 | 0 |
| `xlsx-win` | 1 | 1 |
<!-- END GENERATED: validation-summary -->

## Validation and CI

Regenerate repository documentation and declared mirrors:

```powershell
python .\tools\generate_repository_artifacts.py
```

Fail without modifying files when generated output is stale:

```powershell
python .\tools\generate_repository_artifacts.py --check
```

Hosted validation exposes:

- **Repository integrity** — manifest, generated documentation and mirrors, metadata, provenance, hooks, and agent definitions;
- **Static and contract checks** — PowerShell analysis, Python compilation, Office boundaries, and visual-runtime contracts;
- **Behavioral tests** — Document Handoff and deep-planning validation;
- **Required** — the stable branch-protection aggregate.

Code Mapper has a dedicated CodeQL workflow. True Office runtime validation remains separate because GitHub-hosted runners do not provide reliable Microsoft Office COM automation; `.github/workflows/office-smoke.yml` uses a controlled self-hosted Windows runner.

The shared visual runtime uses a **render-inspect-lint-revise** loop and is checked by `tools/test_visual_runtime_contract.py`. Codex visual skills are checked by `tools/test_visual_skills_contract.py`. `.codex/hooks.json` contains warning-only reminders, validated by `tools/validate_codex_hooks.py`; hooks do not instantiate Office COM.

## Repository layout

- [`skills-manifest.json`](./skills-manifest.json): operational inventory and generation source;
- [`docs/repository-contract.md`](./docs/repository-contract.md): packaging, generation, and maintenance rules;
- [`.generated/`](./.generated): generated registries, including mirror hashes;
- [`.shared/office-com/`](./.shared/office-com): Office preflight and COM boundary runtime;
- [`.shared/visual-runtime/`](./.shared/visual-runtime): browser screenshot, lint, PDF, and image-evidence tooling;
- [`.codex/agents`](./.codex/agents): specialized review and packaging agents;
- [`scaffolds/`](./scaffolds): reusable instruction scaffolds that are explicitly not skills;
- [`archive/`](./archive): non-installable historical skill packages;
- each remaining manifest-declared top-level skill directory: an active canonical skill definition and implementation;
- [`.github/workflows/`](./.github/workflows): hosted validation, drift checks, and Office smoke workflows;
- [`tools/`](./tools): generators, validators, and contract tests.

## Provenance and licensing

`skills-manifest.json` owns package and lifecycle facts; `.provenance/source-registry.json` owns source, immutable revision, license-review, alignment, and distribution facts. Repository-integrity CI requires the manifest's materialized source fields to agree with the provenance registry, so source typos or stale revisions cannot silently pass as valid metadata.

The root [`LICENSE`](./LICENSE) is MIT and applies to material classified as `repo-owned-original`. It does **not** relicense external derivatives or unresolved local imports. Matt Pocock and UltraPlan derivatives retain their recorded MIT terms; Anthropic-derived Office skills retain their applicable upstream license boundaries; `deep_planning.txt` derivatives and `document-handoff` remain restricted while their original-source licensing is unresolved. See [`docs/licensing-and-redistribution.md`](./docs/licensing-and-redistribution.md) for the authoritative policy.

Pinned revisions are reviewed baselines, not a claim that every adaptation automatically tracks upstream `main`. Scheduled drift checks identify upstream changes; pins move only after an explicit alignment review.

## Contributing

Update `skills-manifest.json` whenever skill inventory, grouping, support, runtime, packaging, or validation changes. Update `.provenance/source-registry.json` when source identity, pinned revision, license review, alignment, or distribution changes; then synchronize the manifest's source projection and regenerate repository artifacts. Document scaffolds under `scaffolds/` without skill metadata; move retired packages under `archive/` and mark them `archived` in the manifest. CI fails on inventory drift, manifest/provenance disagreement, stale generated README sections or mirrors, invalid metadata, or namespace violations.
