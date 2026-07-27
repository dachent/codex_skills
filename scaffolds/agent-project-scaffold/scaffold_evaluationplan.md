# Decision memo: project scaffold architecture

**Date:** 26 July 2026  
**Question:** Is the operating scaffold a graph/loop; which external concepts improve it; and what should remain portable across Claude Code/Opus 5, Codex/GPT-5.6 Sol, Hermes/GLM-5.2, and Kimi Code CLI/K3?

## Decision

Keep the [gist scaffold](https://gist.github.com/dachent/cdc05151d047708c290bd4da0aaeed96) as the **sole canonical process** and edit it directly. Do not wrap or invoke it from a skill, copy it into `references/scaffold.md`, or add a database, event log, parser, or shadow tracker. Jira or Markdown remains authoritative state; the sprint controller is the loop; the runtime and its tools, permissions, isolation, and hooks are the harness. If deterministic skill discovery later becomes worth its maintenance cost, migrate the entire canon into one flat `SKILL.md` and retire the gist—never maintain both.

This choice rests on operator evidence: the gist has supported two-month projects and 4–6-hour sprints. That is meaningful but anecdotal, not a controlled performance proof. The proposed additions are justified only where they close observed failure classes: false readiness, unbounded retry, conflicting writes, context loss, verifier gaming, stale state, and silent continuation beyond authority.

## What the system actually is

The precise model is a **state machine whose controller loops over authoritative work state; that state is graph-shaped when recorded relationships constrain readiness**:

| Term | Finding | Consequence |
|---|---|---|
| **Graph** | A backlog is graph-shaped only where recorded relationships constrain readiness. Jira can encode this; Markdown is only a list unless it records dependencies. | Select dependency-ready items; use native Jira/Markdown links, not a graph database. |
| **Loop** | A sprint is one iteration. The loop is `select ready → execute → verify → update current item → select next/stop`. Scheduling merely triggers it. | Preserve the end-of-sprint transition; no loop service is needed. |
| **Harness** | The gist is policy/controller instructions. The CLI, filesystem, sandbox, tools, permissions, workers, worktrees, connectors, and hooks provide enforcement. | Preflight capabilities; never infer isolation from prose. |
| **Gate** | Deterministic checks are evidence, not truth. Failure should cause bounded correction or stop—not “loop until pass.” | Declare acceptance evidence before action; add fresh-context review only when risk warrants it. |

## Evidence that changes the design

The social threads below are design claims, not empirical authority; their implementations and failure modes matter more.

| Source | Retain | Reject |
|---|---|---|
| [Loop → Graph → Harness thread](https://x.com/i/status/2080621294979023358) and [implementation](https://github.com/Archive228/loop-graph-harness) | Explicit boundaries; compact task/result contracts; bounded worker context; budgeted fan-out; integration ownership. | “Clean” as empty context; parent retaining only returns; mandatory final checker; “cannot exit until verifier passes.” These can omit constraints/evidence, over-verify, or retry indefinitely. |
| [Loop-engineering archive](https://twitter-thread.com/t/2064374643729773029) and [original thread](https://x.com/0xCodez/status/2064374643729773029) | Durable external state; manual-first operation; hard limits; permission, collision, and unattended-loop security controls. | Weekly recurrence and 50% acceptance thresholds; scheduling as the loop definition; automated tests as universal truth; mandatory maker/checker split; a new state file beside Jira/Markdown. |
| Rejected [agent-project-orchestrator](https://github.com/dachent/skills/tree/main/agent-project-orchestrator) and [issue #62](https://github.com/dachent/skills/issues/62) | Its failure is the strongest negative test for this decision. | Seven SQLite tables, an event log, and `mirror-import` assumed a fixed Markdown grammar absent from the gist and real artifacts. The design also admitted transaction/event gaps and silent natural-key drops. It created duplicate authority without demonstrated execution value. |

It does **not** prove gists are intrinsically superior; it proves that an additional architectural layer must earn its synchronization and enforcement costs.

## Archive228 repository assessment

[Archive228 currently lists nine repositories](https://github.com/Archive228?tab=repositories). Passing fixtures show that a demo reproduces its intended example—not production isolation, authorization, concurrency safety, or contract completeness. The limitations below come from code and test review.

| Repository | Disposition |
|---|---|
| [file-memory-vs-vector](https://github.com/Archive228/file-memory-vs-vector) | **Take narrowly:** stable-key current truth plus supersession history, using existing Jira/Markdown. **No new store:** 11 fixtures compare an engineered top-1 bag-of-words baseline with files; real vector systems can update, filter, and rerank, while the file code lacks atomicity, locking, compare-and-swap, and path validation. |
| [graph-vs-vector-provenance](https://github.com/Archive228/graph-vs-vector-provenance) | **Conditional:** native links `dependency → artifact → evidence → commit` for genuinely multi-hop provenance. **No graph DB:** 10 fixtures use a hand-authored toy graph, naive entity linking/BFS, no confidence/version/time model, and a top-1 vector strawman. |
| [loop-graph-harness](https://github.com/Archive228/loop-graph-harness) | **Take concepts:** bounded workers, compact contracts, budgeted fan-out, merge owner, gate calibration. **Reject runtime:** despite 18 passing fixtures, context isolation is modeled in-process, fan-out is sequential, HITL auto-approves by default, return size is unenforced, cost is charged after spawning, and merge verification is weak. |
| [verify-gate-loop](https://github.com/Archive228/verify-gate-loop) | **Use for deterministic local mutations:** `propose → precheck → one-writer commit`; rejection writes nothing; preserve a rollback point. Its 11 fixtures do not solve authorization, concurrency, external effects, or automatic recovery. |
| [verifier-gate](https://github.com/Archive228/verifier-gate) | **Take:** predeclared machine checks, bounded feedback, fail closed. **Do not reuse harness:** nine fixtures execute generated code twice in-process without time/resource/network isolation; finite canned cases can be gamed; only the latest failure survives; source is checked rather than integrated state. |
| [adversarial-contract-gate](https://github.com/Archive228/adversarial-contract-gate) | **High-risk use only:** a fresh reviewer receives authoritative criteria, negative/boundary cases, and the merged artifact—not maker rationale. Its nine fixtures mask hard-coded “negotiation,” mutable lists inside a frozen dataclass, circular ground truth, and unsandboxed in-process evaluation. |
| [loopkit](https://github.com/Archive228/loopkit) | **Cherry-pick only:** keep procedures small; periodically strip and retest scaffolding after model changes. **Reject package:** `run.sh` is unbounded, ignores verifier failure, trusts maker-editable `STATUS`, does not enforce `BLOCKED`, exposes mutable MCP/npm surfaces, duplicates pre-compaction state, masks an empty sync check, has documentation/count drift, and is expressly Claude-shaped. |
| [fable-sentinel](https://github.com/Archive228/fable-sentinel) | **Meta-pattern only:** an optional runtime/model-drift canary with an accepted baseline and fail-closed credential/API errors. Direct use is Fable-specific; missing credentials can report a dry-run, API failures are omitted, first run self-baselines, diffs do not fail, prices are hard-coded, and prompts are logged. Ten explicit tests pass, but the declared `npm test` fails under the reviewed environment. |
| [memory](https://github.com/Archive228/memory) | **Reject:** sprint close already supplies the shift note. This adds vendor-specific, model-written duplicate state; review found fail-open consolidation, raw transcript-tail persistence with secret risk, README drift, and failing expected-file checks under `make test`. |

## Applicability across the four targets

The semantics are designed to port; conformance and enforcement remain runtime- and surface-specific. Shared [Agent Skills syntax](https://agentskills.io/specification) does not equal shared invocation, permission, hook, isolation, or concurrency behavior.

| Target | Application and limit |
|---|---|
| **Claude Code / Opus 5** | [Skills](https://code.claude.com/docs/en/skills), [subagents](https://code.claude.com/docs/en/sub-agents), and [worktrees](https://code.claude.com/docs/en/worktrees) can implement the policy. Worktrees isolate checkouts, not services or external effects. Follow [Opus 5 guidance](https://platform.claude.com/docs/en/build-with-claude/prompt-engineering/prompting-claude-opus-5): run each declared gate once per unchanged candidate and rerun only after a relevant repair or input/environment change; do not add a generic verifier subagent. Add fresh review only for sizeable independent or risk-justified work. |
| **Codex / GPT-5.6 Sol** | [Skills](https://learn.chatgpt.com/docs/build-skills), [subagents](https://learn.chatgpt.com/docs/agent-configuration/subagents), [worktrees](https://learn.chatgpt.com/docs/environments/git-worktrees), and [scheduled tasks](https://learn.chatgpt.com/docs/automations) support the pattern, but availability and guarantees vary by surface. Phase 8a must discover them. Parallelism is best for independent/read-heavy work; write-heavy work needs explicit isolation and integration ownership. |
| **Hermes / GLM-5.2** | Hermes supports [skills](https://hermes-agent.nousresearch.com/docs/user-guide/features/skills), [context files](https://hermes-agent.nousresearch.com/docs/user-guide/features/context-files), and [delegation](https://hermes-agent.nousresearch.com/docs/user-guide/features/delegation). Subagents receive isolated conversations and terminal sessions, but those guarantees do not establish disjoint write surfaces. Persist tracker state before [context compression](https://hermes-agent.nousresearch.com/docs/developer-guide/context-compression-and-caching/). Treat tests as evidence because [GLM-5.2](https://z.ai/blog/glm-5.2) explicitly addresses reward-hacking pressure. |
| **Kimi Code CLI / K3** | Target current [Kimi Code](https://github.com/MoonshotAI/kimi-code), its [skills](https://www.kimi.com/code/docs/en/kimi-code-cli/customization/skills.html), and [subagents](https://www.kimi.com/code/docs/en/kimi-code-cli/customization/agents.html); the legacy [Kimi CLI](https://github.com/MoonshotAI/kimi-cli) is winding down. Subagents have isolated context, but filesystem/worktree isolation is not documented: default to one writer without an explicitly isolated checkout. [K3](https://www.kimi.com/blog/kimi-k3) is the model, not the orchestration layer. |

Across all four: write Phase outcomes first; use a named skill only when installed; record a substitution only when it materially changes execution or evidence. If Jira is authoritative but inaccessible, block or use a human-mediated update—never create shadow Markdown.

## Recommended edits to the gist

| Location | Minimum edit | Value |
|---|---|---|
| **Operating rules / Phase 0** | Define evidence-based resume/audit entry; one authoritative tracker; outcome-first fallback for unavailable commands; and reapproval when target, acceptance, architecture, authority, external effects, or irreversible risk changes. Remove generic “strongest model justified” advice. | Prevents restart drift, duplicate truth, fabricated commands, silent scope creep, and runtime-aging prose. |
| **Phase 4** | State the expected observation before each probe. | Prevents retrofitting interpretations to results. |
| **Phase 6** | Define the required architecture/dependency/risk outcome; use `/ultraplan` only if installed. | Keeps the outcome independent of one command ecosystem. |
| **Phase 7** | Write into existing Markdown/Jira. Require only a stable reference, outcome, material dependencies, approved scope, and acceptance evidence; add write/stop boundaries only where risk warrants. Use tracker-native fields/comments. | Makes readiness computable without imposing a universal schema or second tracker. |
| **Phase 8a** | Verify readiness, dependencies, non-destructive external reachability, permissions, checks, and authority; label baseline failures and forbid regressions; test whether context, writes, and side effects can be bounded; set observable attempt/time/cost/context limits. | Prevents blind execution, false attribution, unsafe delegation, and runaway autonomy. |
| **Phase 8b—action** | Make the smallest coherent, reversible change; run declared checks; record exact evidence. Retry only after a materially different corrective action, repaired failure, or changed input/environment. Gate changes require reapproval; never “loop until pass.” | Prevents retry theatre, gate weakening, and unapproved workarounds. |
| **Phase 8b—delegation** | Send a bounded task packet and require a compact result/evidence reference. Use one writer unless write **and external-effect** surfaces are demonstrably disjoint; assign merge ownership and verify integrated state. If authority cannot be bounded, reduce autonomy or stop—local execution is not automatically safer. | Prevents context loss, conflicts, unsafe fan-out, and individually passing but broken merges. |
| **Sprint close** | Update the current item with outcome, evidence, decisions, Dead Ends, blockers, and new dependencies; route `complete → 8c`, `no ready item → blocked`, `limit/block/reapproval → handoff and stop`, otherwise `ready → next item within authority/limits`. | Supplies the actual return edge and prevents stale Jira or unauthorized continuation. |
| **Phase 8c** | Validate each candidate final integrated state against Phase 0 criteria; do not add duplicate generic checker passes. Run each declared gate once per unchanged candidate and rerun only after a relevant repair or input/environment change. Treat tests as evidence; add a fresh-context reviewer only for impact, ambiguity, security, or gameable checks. Put reusable lessons in an already approved store; otherwise keep them in the handoff. | Preserves final acceptance without ritualized over-verification or another stale state surface. |
| **Compression pass** | Define tracker, reapproval, fallback, and stop semantics once; remove their repetitions and vague clauses such as “any other material finding.” | Reduces prompt tax and contradictory restatements. |

Leave Phases 1–3 and 5, the human approval cadence, and the existing Dead Ends mechanism intact.

The update is **architecturally surgical, not textually minor**: the original is 534 words; the current draft is 1,184 (+650; 2.22×). That prompt tax is too high for a final form. Compress toward roughly 750–850 words while preserving seven indispensable controls: single authority; readiness/dependencies; bounded retry/stop; one writer/integration owner; exact sprint-close transition; checks as evidence; selective review. Task/result packets, expected probe observations, and runtime fallback are useful but should each be stated once.

## Validation boundary

Review covered the exact gist, both supplied Markdown sources, source and tests for all nine repositories, the rejected orchestrator record, and official runtime documentation. An adversarial swarm separately attacked source support, architecture, gist preservation, and compression. A disposable Codex/Markdown sprint completed one item, checked it once, updated backlog/handoff, and selected—but did not start—the next item outside authorized scope. Jira, Claude, Hermes, and Kimi/K3 were not live-certified because the necessary project/runtime configurations were unavailable; portability is therefore documentation-grounded and capability-gated.
