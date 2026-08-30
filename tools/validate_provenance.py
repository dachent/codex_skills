from __future__ import annotations

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
