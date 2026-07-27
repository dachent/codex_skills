# Deep planning project scaffold

{PROBLEM DESCRIPTION}

## Operating rules

- Resume at the earliest phase whose required output or approval lacks current, trustworthy evidence. For an audit, enter wherever needed to test the requested claims; prior work is evidence, not permission to skip review.
- Use one authoritative execution tracker: the existing Jira project or Markdown backlog, when available, otherwise the Phase 7 Markdown plan. Record dependencies and readiness there; never create a shadow tracker. If the authoritative tracker is inaccessible, block or use a human-mediated update rather than substituting another tracker.
- Use each named skill when installed. Otherwise produce the phase outcome with native capabilities and record any substitution that materially changes execution or evidence. A handoff records the phase output, decisions and approval, current state, and next phase.
- Keep target, scope, criteria, architecture, authority, external effects, and irreversible risk fixed during an approved execution attempt. Any material change returns to the relevant planning phase and approval gate.
- After Phases 0, 1, 2b, 3, 4, 5, 6, and 7, stop for explicit user PROCEED; after approval, write a handoff before continuing. Phase 8a separately stops for PROCEED, which authorizes Phase 8b within the recorded boundaries.

## Phase 0

Use `/mattpocock:grill-with-docs` to sharpen the user's draft success criteria, failure criteria, and out-of-scope boundaries rather than inventing them. Estimate complexity and context needs. Consult an existing approved cross-project lessons or Dead Ends store and incorporate relevant entries before grilling.

## Phase 1

Catalog relevant files, logs, code, inputs, outputs, and dependencies as `[what | attempt | outcome: succeeded/failed/partial/unreached | known failure mechanism]`.

## Phase 2a

Classify the catalog by relevance; fully review high-relevance materials and summarize lower-relevance materials with rationale.

## Phase 2b

Use `/superpowers:systematic-debugging` on the failure corpus to produce a failure autopsy covering root causes, wrong decision points, partial state, and actual accomplishments. Create a Dead Ends Registry of approaches not to pursue.

## Phase 3

Use `/mattpocock:grill-with-docs` against the synthesis and failure corpus to expose missing assumptions, risks, and edge cases. Add newly eliminated paths to Dead Ends.

## Phase 4

Run targeted, low-cost, high-impact probes. State the expected observation before each probe, eliminate failed paths, and update Dead Ends.

## Phase 5

Use `/superpowers:brainstorming` to design the memo, workfolder, catalog, and handoff approach. Reject proposals conflicting with Dead Ends. After presenting approaches and writing the design, stop; Phase 6, not brainstorming, owns the execution plan.

## Phase 6

Resolve the selected path's architecture, sequence, dependencies, risks, and open questions. Use `/ultraplan` when installed.

## Phase 7

Use `/superpowers:writing-plans` to record the final execution plan in the authoritative tracker. Give every item a stable reference, outcome, material dependencies, approved scope, and acceptance evidence; add allowed writes and stop conditions when relevant. Include validation checkpoints and rollback triggers.

## Phase 8a

Reconcile tracker state, directories, and tools. Verify item readiness, dependencies, permissions, checks, and non-destructive reachability of required systems; record material pre-existing failures. Determine whether worker context, authority, writes, and external effects can be bounded. Set observable attempt, time, cost, or context limits for approved autonomy.

## Phase 8b

Select a ready, authorized item and execute the approved plan using `/superpowers:subagent-driven-development` or `/superpowers:dispatching-parallel-agents` only for safely independent work; otherwise use `/superpowers:executing-plans` or native execution. Make the smallest coherent, preferably reversible change. Run declared checks and record exact evidence.

Retry only after a new diagnosis or a relevant artifact, candidate, input, environment, or gate change, and only within approved limits. Never loop until pass or weaken criteria without reapproval. On a blocker, exhausted limit, or required reapproval, hand off current state and stop.

Delegate only when context, authority, writes, and external effects are bounded. Send `[item | outcome | scope | instructions/references/Dead Ends | allowed writes | acceptance evidence | limits]`; require `[status | artifact/diff | checks/evidence | uncertainty | blockers/dependencies | next action]`. Use one writer unless write and external-effect surfaces are demonstrably independent. Assign integration ownership and verify merged state. If boundaries cannot be enforced, reduce autonomy or stop.

After each sprint, update the current item with outcome, evidence, decisions, Dead Ends, blockers, and dependencies. Then route: plan complete -> Phase 8c; ready and authorized work remains -> next item; no item ready -> blocked; limit, blocker, or reapproval -> handoff and stop.

## Phase 8c

Use `/superpowers:verification-before-completion` to test the integrated result against Phase 0 criteria and boundaries. Run each declared gate once per unchanged candidate; rerun only after a relevant repair or input/environment change. Treat deterministic checks as evidence, not infallible proof. Add a fresh reviewer only when impact, ambiguity, security, or gameable checks justify it, and provide criteria, constraints, integrated artifacts, and evidence rather than maker rationale.

Produce the final handoff with deliverables, final tracker state, Dead Ends, and lessons delta. Append reusable lessons only to an already approved durable store; otherwise retain them in the handoff. Finish with `/mattpocock:handoff` when installed.
