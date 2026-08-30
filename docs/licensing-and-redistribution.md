# Licensing and redistribution

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
