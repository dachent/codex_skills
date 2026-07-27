# Project scaffolds

This namespace contains reusable agent operating scaffolds. A scaffold is a document that a user supplies as task or project instructions; it is not an installable skill.

Scaffold packages must:

- include a package `README.md` that explains use, scope, and provenance;
- contain no `SKILL.md` or `agents/openai.yaml`;
- remain outside `skills-manifest.json` and the generated installation inventory;
- keep validation and evaluation evidence beside the scaffold when that evidence is part of the published package.

Do not copy this directory into a Codex, Claude Code, or other agent skill directory. Open the selected package and use its scaffold according to its README.

## Available scaffolds

- [`agent-project-scaffold`](./agent-project-scaffold): a gated project-planning and execution scaffold migrated from Gist history, with its evaluation plan and Codex validation report.
