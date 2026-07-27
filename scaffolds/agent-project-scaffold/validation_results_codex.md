# Scaffold validation report

**Date:** 26 July 2026  
**Candidate:** `publish/claude_code_deep_planning.md`  
**Candidate SHA-256:** `B6113AAFEAD3856EFB1485039C4F801A8DFD30B41B5C0715C16B99C6B1921AC2`  
**Publication verdict:** **PASS**

## Outcome

The proposed scaffold was compressed from 1,184 to 811 words while retaining all 12 phases, approval gates, authoritative-state rules, dependency readiness, bounded retries, delegation contracts, one-writer controls, exact sprint routing, evidence rules, rollback triggers, and selective review. Four unsafe or stale clauses were removed or corrected: generic model advice, automatic local fallback when authority cannot be bounded, unapproved durable-store creation, and substitution of a tracker when the authoritative tracker is inaccessible.

## Verification matrix

| Criterion | Proof | Expected | Result | Evidence | Failure or recovery trigger |
|---|---|---|---|---|---|
| Structure and specificity | `tests/structural/Test-Scaffold.ps1` | 750–850 words; exact phase order; required controls present; rejected clauses absent | **PASS** — 811 words, 30/30 checks | `evidence/structural-output.json` | Any failed assertion blocks publication |
| State-machine safety | `tests/structural/Test-StateTransitions.ps1` | Correct routing for completion, readiness, blockers, retries, authority, writers, tracker access, and lessons storage | **PASS** — 9/9 scenarios | `evidence/state-transition-output.json` | Any incorrect route blocks publication |
| Runtime portability | `tests/structural/Test-Portability.ps1` | Codex, Claude, Hermes, and Kimi discoverable; runtime-neutral fallback semantics present | **PASS** | `evidence/portability-output.json` | Missing command or semantic contract blocks publication |
| GitHub Markdown | `tests/rendering/Test-GfmRendering.ps1` | Scaffold headings/lists/code and memo headings/links/tables render through GitHub GFM | **PASS** | `evidence/rendering-output.json`; rendered HTML under `tests/rendering` | Rendering defect requires correction and retest |
| Codex behavior | Ephemeral Codex 0.145.0 / GPT-5.6 Sol fixture | Execute only SC-001; one declared check; one tracker; record native substitution; stop before SC-002 | **PASS** — 16/16 postconditions | `evidence/codex-smoke-verification.json`; `evidence/codex-smoke-output.txt` | Any scope escape, repeated check, protected-file mutation, or shadow state blocks publication |
| Adversarial review | Mixed technical/business review | No BLOCKING or IMPORTANT finding | **PASS** | `evidence/adversarial-review.md` | Any unresolved BLOCKING or IMPORTANT finding blocks publication |

## Smoke-test attempt history

Attempt 1 was a harness failure: `--ignore-user-config` prevented the CLI from using its authenticated configuration, and both transports returned HTTP 401 before model execution. No fixture file changed. The retry removed that flag after `codex login status` confirmed ChatGPT authentication, which was a material environment change permitted by the scaffold's retry rule.

Attempt 2 completed SC-001, created the exact 20-byte LF-terminated deliverable, ran the declared check once with exit 0 and `SMOKE_CHECK_PASS`, updated only the authoritative tracker and handoff, marked SC-002 ready but unauthorized, and did not create `second-item.txt`.

## Source and recovery evidence

- Current pre-update gist revision: `ddeb80cea25ff158f9264a8d7abe4016b9c12e36`.
- Current scaffold snapshot: `claude_code_deep_planning.txt`, SHA-256 `EC9E91D9552809BF2BF3E16F1262E126B09285087090EABCCC269A84DA5B1672`.
- Evaluation memo source and publish copy both have SHA-256 `5CDBEC3D47510779984D275E774CA055CDC18B79016292CF59297F90D2573FE6`.
- Full hashes are recorded in `evidence/content-hashes.json`.
- The remote update is permitted only if the live revision and content still match the frozen snapshot. A post-update mismatch triggers a compensating restoration of the `.txt` scaffold and removal of the added memo.

## Limitations

- Codex received the only live behavioral certification.
- Claude, Hermes, and Kimi portability is documentation- and capability-grounded but not live-certified.
- Model behavior can drift; the retained tests provide a reproducible regression package rather than a permanent guarantee.

## Publication verification

**PASS.** The initial guarded update advanced the gist from `ddeb80cea25ff158f9264a8d7abe4016b9c12e36` to `93db3febb8eefb4b65e049bbb36a9ae70fc14fec`. A subsequent approved publication renamed the memo and added this validation report; its final revision is recorded in `evidence/gist-after-followup.json` to avoid a self-referential report update.

- Remote files are exactly `claude_code_deep_planning.md`, `zzz_scaffold_evaluationplan.md`, and `zzz_validation_results_codex.md`; the former `.txt` and memo filenames are absent.
- Retrieved scaffold SHA-256 is `B6113AAFEAD3856EFB1485039C4F801A8DFD30B41B5C0715C16B99C6B1921AC2`, exactly matching the staged candidate.
- Retrieved memo SHA-256 is `5CDBEC3D47510779984D275E774CA055CDC18B79016292CF59297F90D2573FE6`, exactly matching the moved source memo.
- Owner `dachent`, public visibility, and description `Deep planning general prompt template` are unchanged.
- Both remote files report `text/markdown`, neither is truncated, and every post-update check passed.
- Rollback was not invoked.

The initial and follow-up API responses and independently re-read remote states are retained under `evidence`.
