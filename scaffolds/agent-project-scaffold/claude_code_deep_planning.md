{PROBLEM DESCRIPTION}

Phase 0:

Define success criteria, failure criteria, and out-of-scope boundaries; select model; estimate complexity and context window requirements; then run /mattpocock:handoff to write Phase 0 state to disk; then stop for PROCEED.

Phase 1:

Catalog all relevant files, logs, code, inputs, outputs, and dependencies using the following metadata for each: [what it is | attempt it belongs to | outcome: succeeded / failed / partial / unreached | failure mechanism if known]; then run /mattpocock:handoff; then stop for PROCEED.

Phase 2a:

Inspect all cataloged materials; classify relevance; fully review high-relevance materials; summarize low-relevance with rationale.

Phase 2b:

Run /superpowers:systematic-debugging on the failure corpus to construct a failure autopsy: root causes, decision points where wrong paths were taken, partial completion state, and what was actually accomplished; output a Dead Ends Registry (approaches NOT to pursue); then run /mattpocock:handoff; then stop for PROCEED.

Phase 3:

Run /mattpocock:grill-with-docs on the synthesis and Dead Ends Registry to expose missing assumptions, risks, and edge cases — challenging against the actual failure corpus, not abstract assumptions; update Dead Ends Registry with newly identified eliminations; then run /mattpocock:handoff; then stop for PROCEED.

Phase 4:

Run targeted low-cost high-impact probes to validate the critical assumptions surviving Phase 3; eliminate probe-failed paths; update Dead Ends Registry; then run /mattpocock:handoff; then stop for PROCEED.

Phase 5:

Run /superpowers:brainstorming to design the memo, workfolder, catalog, and handoff approach; all proposals must be checked against Dead Ends Registry before inclusion; then run /mattpocock:handoff; then stop for path selection and PROCEED.

Phase 6:

Run /superpowers:ultraplan on the selected path; then run /mattpocock:handoff; then stop for PROCEED.

Phase 7:

Run /superpowers:writing-plans to write the final execution plan with explicit validation checkpoints and rollback triggers; write plan to disk; then run /mattpocock:handoff; then stop for approval and PROCEED.

Phase 8a:

Setup — establish state document, directory structure, and tooling; verify all dependencies are reachable; then stop for PROCEED.

Phase 8b:

Execute approved plan using /superpowers:subagent-driven-development or /superpowers:dispatching-parallel-agents where tasks are independent, otherwise /superpowers:executing-plans; update state document at each milestone; on any blocking failure, stop and surface with current state rather than attempting workarounds.

Phase 8c:

Run /superpowers:verification-before-completion to validate outputs against Phase 0 success criteria; produce final handoff package including deliverables, state document, updated Dead Ends Registry, and lessons learned delta; then run /mattpocock:handoff as final archival record.