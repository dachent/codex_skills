from __future__ import annotations

import json
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MANIFEST = ROOT / "skills-manifest.json"
REGISTRY = ROOT / ".provenance" / "source-registry.json"
README = ROOT / "README.md"
GENERATOR = ROOT / "tools" / "generate_repository_artifacts.py"
PROVENANCE_VALIDATOR = ROOT / "tools" / "validate_provenance.py"
LICENSING_DOC = ROOT / "docs" / "licensing-and-redistribution.md"
CONTRACT_DOC = ROOT / "docs" / "repository-contract.md"
SELF = ROOT / "tools" / "normalize_authorities_20260830.py"
WORKFLOW = ROOT / ".github" / "workflows" / "normalize-authorities-20260830.yml"


def write(path: Path, text: str) -> None:
    path.write_text(text.rstrip() + "\n", encoding="utf-8")


def replace_section(text: str, start_heading: str, end_heading: str, replacement: str) -> str:
    start = text.index(start_heading)
    end = text.index(end_heading, start)
    return text[:start] + replacement.rstrip() + "\n\n" + text[end:]


def normalize_manifest(manifest: dict, registry: dict) -> None:
    policy = manifest["policy"]
    policy["installable_statuses"] = ["supported", "experimental"]
    policy["provenance_registry"] = ".provenance/source-registry.json"

    records = registry["skills"]
    sources = registry["sources"]
    for skill in manifest["skills"]:
        record = records.get(skill["name"])
        if record is None:
            continue
        classification = record["classification"]
        source_id = record.get("source")
        if source_id is None:
            skill["source"] = {"classification": classification}
            continue

        source = sources[source_id]
        projected = {"classification": classification}
        if source.get("repository"):
            projected["repository"] = source["repository"]
        projected["path"] = record["source_path"]
        if classification == "local-source-import":
            projected["initial_commit"] = source["revision"]
            if source.get("license_review") == "restricted-pending-review":
                projected["provenance_status"] = "unresolved-original-source"
        else:
            projected["revision"] = source["revision"]
        skill["source"] = projected


def normalize_registry(registry: dict) -> None:
    registry["reviewed_on"] = "2026-08-30"
    registry["repository_license"] = {
        "status": "mixed-license",
        "policy_document": "docs/licensing-and-redistribution.md",
        "root_license": "LICENSE",
        "default_for_repo_owned_originals": "MIT",
        "root_license_scope": "The root MIT license applies to repository-owned original material only; external derivatives and unresolved imports remain governed by their recorded source and distribution terms."
    }
    for record in registry["skills"].values():
        if record.get("classification") == "repo-owned-original":
            record["distribution"] = "MIT"


def update_generator() -> None:
    text = GENERATOR.read_text(encoding="utf-8")
    old = "def active(m): return [s for s in m['skills'] if s.get('status')!='archived']\ndef notice(name):"
    new = "def active(m): return [s for s in m['skills'] if s.get('status')!='archived']\ndef installable(m):\n statuses=set(m.get('policy',{}).get('installable_statuses',['supported','experimental']))\n return [s for s in active(m) if s.get('status') in statuses]\ndef notice(name):"
    if old not in text:
        raise RuntimeError("generator active() anchor not found")
    text = text.replace(old, new, 1)
    text = text.replace("comps={s['name']:[] for s in active(m)}", "comps={s['name']:[] for s in installable(m)}", 1)
    text = text.replace(
        "out=[notice('installation inventory'),'','Only the entries below are installable. Their top-level skill directories are canonical; `scaffolds/` and `archive/` are excluded. Copy only the skills required by the target agent, plus listed shared components.'",
        "out=[notice('installation inventory'),'','Only supported or experimental entries below are installable. Deprecated skills remain cataloged for compatibility but are excluded from this inventory; `scaffolds/` and `archive/` are also excluded. Copy only the skills required by the target agent, plus listed shared components.'",
        1,
    )
    text = text.replace("for s in sorted(active(m),key=lambda x:x['name']): out.append(f\"| `{s['name']}` | [`{s['path']}`](./{s['path']}) | {fmt(comps.get(s['name'],[]))} |\")", "for s in sorted(installable(m),key=lambda x:x['name']): out.append(f\"| `{s['name']}` | [`{s['path']}`](./{s['path']}) | {fmt(comps.get(s['name'],[]))} |\")", 1)
    write(GENERATOR, text)


def update_readme_prose() -> None:
    text = README.read_text(encoding="utf-8")
    old_model = "`skills-manifest.json` is the operational source of truth for active and archived skill packages, catalog grouping, ownership, support, provenance classification, packaging, shared runtimes, validation, and generated agent mirrors. Each active skill has one canonical top-level directory. [`scaffolds/`](./scaffolds) contains documented non-skills; [`archive/`](./archive) contains non-installable historical packages. See [`docs/repository-contract.md`](./docs/repository-contract.md)."
    new_model = "`skills-manifest.json` is the operational authority for skill inventory, lifecycle, grouping, ownership, platform and agent support, packaging, shared runtimes, validation, and generated mirrors. [`.provenance/source-registry.json`](./.provenance/source-registry.json) is the authority for external source identity, immutable source revisions, license review, alignment metadata, and distribution boundaries. Source facts duplicated into the manifest are a CI-validated projection of that registry, not an independent authority. Each active skill has one canonical top-level directory. [`scaffolds/`](./scaffolds) contains documented non-skills; [`archive/`](./archive) contains non-installable historical packages. See [`docs/repository-contract.md`](./docs/repository-contract.md)."
    if old_model not in text:
        raise RuntimeError("README repository-model anchor not found")
    text = text.replace(old_model, new_model, 1)

    provenance = """## Provenance and licensing

`skills-manifest.json` owns package and lifecycle facts; `.provenance/source-registry.json` owns source, immutable revision, license-review, alignment, and distribution facts. Repository-integrity CI requires the manifest's materialized source fields to agree with the provenance registry, so source typos or stale revisions cannot silently pass as valid metadata.

The root [`LICENSE`](./LICENSE) is MIT and applies to material classified as `repo-owned-original`. It does **not** relicense external derivatives or unresolved local imports. Matt Pocock and UltraPlan derivatives retain their recorded MIT terms; Anthropic-derived Office skills retain their applicable upstream license boundaries; `deep_planning.txt` derivatives and `document-handoff` remain restricted while their original-source licensing is unresolved. See [`docs/licensing-and-redistribution.md`](./docs/licensing-and-redistribution.md) for the authoritative policy.

Pinned revisions are reviewed baselines, not a claim that every adaptation automatically tracks upstream `main`. Scheduled drift checks identify upstream changes; pins move only after an explicit alignment review."""
    text = replace_section(text, "## Provenance and licensing", "## Contributing", provenance)

    old_contrib = "Update `skills-manifest.json` whenever skill inventory, grouping, support, source, runtime, packaging, or validation changes. Document scaffolds under `scaffolds/` without skill metadata; move retired packages under `archive/` and mark them `archived` in the manifest. Then run the generator and check mode. CI fails when an active or archived skill is unregistered, a scaffold masquerades as a skill, generated README sections are stale, a declared mirror diverges, or undeclared files appear under `.claude/skills`."
    new_contrib = "Update `skills-manifest.json` whenever skill inventory, grouping, support, runtime, packaging, or validation changes. Update `.provenance/source-registry.json` when source identity, pinned revision, license review, alignment, or distribution changes; then synchronize the manifest's source projection and regenerate repository artifacts. Document scaffolds under `scaffolds/` without skill metadata; move retired packages under `archive/` and mark them `archived` in the manifest. CI fails on inventory drift, manifest/provenance disagreement, stale generated README sections or mirrors, invalid metadata, or namespace violations."
    if old_contrib not in text:
        raise RuntimeError("README contributing anchor not found")
    text = text.replace(old_contrib, new_contrib, 1)
    write(README, text)


def update_docs() -> None:
    licensing = """# Licensing and redistribution

This repository is a **mixed-license collection**. The root [`LICENSE`](../LICENSE) is MIT, but it does not override licenses, notices, or unresolved rights attached to imported or derivative material.

## Root MIT scope

Material classified as `repo-owned-original` in `.provenance/source-registry.json` is distributed under the root MIT license. The root license does not relicense third-party source material, externally derived packages, or imports whose original source or license remains unresolved.

## External derivatives and imports

The authoritative mapping is `.provenance/source-registry.json`.

- Matt Pocock skill derivatives and the UltraPlan derivative are distributed under their recorded upstream MIT licenses; reviewed license evidence is retained under `.upstream/licenses/`.
- Anthropic-derived Office skills remain subject to the license and notice boundaries applicable to each recorded upstream skill snapshot. This repository does not infer a repository-wide Anthropic license.
- `deep_planning.txt` derivatives and `document-handoff` have unresolved original-source licensing. They are marked `restricted` and must not be redistributed outside this repository until their source owner and license are documented.

## Authority and precedence

`skills-manifest.json` is authoritative for package lifecycle, support, packaging, and validation. `.provenance/source-registry.json` is authoritative for source identity, immutable source revision, license review, alignment metadata, and distribution. Any source facts copied into the manifest are a validated projection of the provenance registry. If a generated catalog conflicts with the registry, the registry controls and generation must be repaired.

A downstream user must evaluate each externally derived skill under its recorded terms. Repository-generated catalogs summarize provenance but do not replace license files, notices, or source-specific obligations.

## Review process

Every active skill has a provenance record. External sources record an immutable revision, source path, retrieval date, license review, port depth, intentional divergence, owner, and last alignment review. `tools/validate_provenance.py` enforces registry coverage, manifest/registry agreement, and root-license scope. Scheduled drift checks compare registered GitHub sources against their pinned revisions and identify local skills requiring alignment review.
"""
    write(LICENSING_DOC, licensing)

    contract = """# Repository integration contract

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

Run `python .\\tools\\generate_repository_artifacts.py` to generate the README catalog, installation inventory, platform/agent matrix, validation summary, declared agent mirrors, and `.generated/agent-mirrors.json`.

Run `python .\\tools\\generate_repository_artifacts.py --check` to fail on stale marked README sections, stale mirrors, stale hashes, or undeclared files under `.claude/skills`.

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
"""
    write(CONTRACT_DOC, contract)


def update_provenance_validator() -> None:
    validator = r'''from __future__ import annotations

import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
REGISTRY = ROOT / ".provenance" / "source-registry.json"
MANIFEST = ROOT / "skills-manifest.json"
SHA_RE = re.compile(r"^[0-9a-f]{40}$")
EXTERNAL = {"local-source-import", "light-adaptation", "medium-adaptation", "heavy-adaptation", "derived-work"}
VALID_REVIEWS = {"reviewed", "reviewed-with-boundaries", "restricted-pending-review"}
VALID_DISTRIBUTIONS = {"MIT", "source-license-boundary", "restricted", "repository-policy"}


def load(path: Path, failures: list[str]) -> dict:
    if not path.is_file():
        failures.append(f"missing file: {path.relative_to(ROOT)}")
        return {}
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        failures.append(f"invalid JSON in {path.relative_to(ROOT)}: {exc}")
        return {}
    if not isinstance(value, dict):
        failures.append(f"{path.relative_to(ROOT)} must contain an object")
        return {}
    return value


def required_text(obj: dict, key: str, context: str, failures: list[str]) -> str | None:
    value = obj.get(key)
    if not isinstance(value, str) or not value.strip():
        failures.append(f"{context}: missing non-empty '{key}'")
        return None
    return value


def main() -> int:
    failures: list[str] = []
    registry = load(REGISTRY, failures)
    manifest = load(MANIFEST, failures)
    if not registry or not manifest:
        for item in failures:
            print(f"ERROR: {item}", file=sys.stderr)
        return 1

    if registry.get("schema_version") != 1:
        failures.append("source registry schema_version must be 1")

    license_policy = registry.get("repository_license")
    default_original_license = None
    if not isinstance(license_policy, dict):
        failures.append("repository_license must be an object")
    else:
        required_text(license_policy, "status", "repository_license", failures)
        policy_document = required_text(license_policy, "policy_document", "repository_license", failures)
        root_license = required_text(license_policy, "root_license", "repository_license", failures)
        default_original_license = required_text(license_policy, "default_for_repo_owned_originals", "repository_license", failures)
        required_text(license_policy, "root_license_scope", "repository_license", failures)
        if policy_document and not (ROOT / policy_document).is_file():
            failures.append(f"repository license policy does not exist: {policy_document}")
        if root_license:
            root_license_path = ROOT / root_license
            if not root_license_path.is_file():
                failures.append(f"root license does not exist: {root_license}")
            elif default_original_license == "MIT" and not root_license_path.read_text(encoding="utf-8").startswith("MIT License"):
                failures.append("repository_license declares MIT but root LICENSE is not MIT")

    sources = registry.get("sources")
    if not isinstance(sources, dict) or not sources:
        failures.append("sources must be a non-empty object")
        sources = {}
    for source_id, source in sorted(sources.items()):
        context = f"sources.{source_id}"
        if not isinstance(source, dict):
            failures.append(f"{context}: must be an object")
            continue
        kind = required_text(source, "kind", context, failures)
        revision = required_text(source, "revision", context, failures)
        required_text(source, "retrieved_on", context, failures)
        required_text(source, "license", context, failures)
        review = required_text(source, "license_review", context, failures)
        if revision and not SHA_RE.fullmatch(revision):
            failures.append(f"{context}.revision must be a 40-character lowercase SHA")
        if review and review not in VALID_REVIEWS:
            failures.append(f"{context}.license_review is unsupported: {review}")
        if kind == "github":
            required_text(source, "repository", context, failures)
            required_text(source, "default_branch", context, failures)
        if review == "reviewed":
            evidence = required_text(source, "license_evidence", context, failures)
            if evidence and not (ROOT / evidence).is_file():
                failures.append(f"{context}: missing license evidence: {evidence}")

    records = registry.get("skills")
    if not isinstance(records, dict) or not records:
        failures.append("skills must be a non-empty object")
        records = {}

    manifest_skills = {item.get("name"): item for item in manifest.get("skills", []) if isinstance(item, dict) and isinstance(item.get("name"), str)}
    active_names = {name for name, item in manifest_skills.items() if item.get("status") != "archived"}
    missing = sorted(active_names - set(records))
    extra = sorted(set(records) - active_names)
    if missing:
        failures.append("active skills missing provenance records: " + ", ".join(missing))
    if extra:
        failures.append("provenance records not present as active manifest skills: " + ", ".join(extra))

    for name in sorted(active_names & set(records)):
        item = manifest_skills[name]
        record = records[name]
        context = f"skills.{name}"
        if not isinstance(record, dict):
            failures.append(f"{context}: must be an object")
            continue
        manifest_source = item.get("source")
        if not isinstance(manifest_source, dict):
            failures.append(f"{name}: manifest source must be an object")
            manifest_source = {}

        classification = required_text(record, "classification", context, failures)
        required_text(record, "port_depth", context, failures)
        required_text(record, "intentional_divergence", context, failures)
        owner = required_text(record, "owner", context, failures)
        required_text(record, "last_alignment_review", context, failures)
        distribution = required_text(record, "distribution", context, failures)
        if distribution and distribution not in VALID_DISTRIBUTIONS:
            failures.append(f"{context}.distribution is unsupported: {distribution}")
        if classification and manifest_source.get("classification") != classification:
            failures.append(f"{name}: manifest source classification disagrees with provenance registry")
        if owner and item.get("owner") != owner:
            failures.append(f"{name}: manifest owner disagrees with provenance registry")

        source_id = record.get("source")
        source_path = record.get("source_path")
        if classification in EXTERNAL:
            if not isinstance(source_id, str) or source_id not in sources:
                failures.append(f"{context}: external derivative must reference a registered source")
                continue
            if not isinstance(source_path, str) or not source_path.strip():
                failures.append(f"{context}: external derivative must record source_path")
                continue
            source = sources[source_id]
            revision = source.get("revision")
            repository = source.get("repository")
            if repository is not None and manifest_source.get("repository") != repository:
                failures.append(f"{name}: manifest source repository disagrees with provenance registry")
            if manifest_source.get("path") != source_path:
                failures.append(f"{name}: manifest source path disagrees with provenance registry")
            if classification == "local-source-import":
                if manifest_source.get("initial_commit") != revision:
                    failures.append(f"{name}: manifest initial_commit disagrees with provenance registry revision")
            elif manifest_source.get("revision") != revision:
                failures.append(f"{name}: manifest source revision disagrees with provenance registry")

            review = source.get("license_review")
            if review == "restricted-pending-review" and distribution != "restricted":
                failures.append(f"{context}: unresolved source licensing requires restricted distribution")
            if source.get("license") == "MIT" and distribution != "MIT":
                failures.append(f"{context}: MIT source must retain MIT distribution")
            if review == "reviewed-with-boundaries" and distribution != "source-license-boundary":
                failures.append(f"{context}: bounded upstream licensing requires source-license-boundary distribution")
        else:
            if source_id is not None or source_path is not None:
                failures.append(f"{context}: repo-owned original must not reference an external source")
            if any(key in manifest_source for key in ("repository", "path", "revision", "initial_commit")):
                failures.append(f"{name}: repo-owned manifest source must not contain external source fields")
            if default_original_license and distribution != default_original_license:
                failures.append(f"{context}: repo-owned original distribution must match root license {default_original_license}")

    if failures:
        for item in failures:
            print(f"ERROR: {item}", file=sys.stderr)
        return 1
    print(f"Provenance validation passed for {len(records)} active skills and {len(sources)} sources.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
'''
    write(PROVENANCE_VALIDATOR, validator)


def main() -> int:
    manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
    registry = json.loads(REGISTRY.read_text(encoding="utf-8"))
    normalize_registry(registry)
    normalize_manifest(manifest, registry)
    write(REGISTRY, json.dumps(registry, indent=2, ensure_ascii=False))
    write(MANIFEST, json.dumps(manifest, indent=2, ensure_ascii=False))

    update_generator()
    update_provenance_validator()
    update_docs()
    update_readme_prose()

    subprocess.run(["python", str(GENERATOR)], cwd=ROOT, check=True)

    # Remove one-shot migration plumbing so the PR contains only durable repository changes.
    if SELF.exists():
        SELF.unlink()
    if WORKFLOW.exists():
        WORKFLOW.unlink()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
