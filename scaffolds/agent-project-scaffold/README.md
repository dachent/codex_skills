# Agent project scaffold

> [!IMPORTANT]
> This package is a **scaffold, not a skill**. Do not install it or copy it into an agent's skills directory. It intentionally has no `SKILL.md` or `agents/openai.yaml`.

Use [`claude_code_deep_planning.md`](./claude_code_deep_planning.md) as project-level instructions when a task benefits from explicit discovery, planning, approval gates, bounded implementation, evidence, and sprint closure. Supply the problem description, follow the scaffold's authority boundaries, and stop at its required approval gates.

## Published files

- [`claude_code_deep_planning.md`](./claude_code_deep_planning.md): the current 811-word scaffold.
- [`scaffold_evaluationplan.md`](./scaffold_evaluationplan.md): the architecture decision and evaluation memo.
- [`validation_results_codex.md`](./validation_results_codex.md): the Codex validation report.

The three files above are byte-exact copies of the latest source Gist files after removing only the requested `zzz_` filename prefixes. This README is repository-only documentation.

## Validation scope

The published evaluation passed structural, Markdown rendering, adversarial state-transition, portability, and isolated Codex smoke checks. Codex received a live behavioral smoke test. Claude Code, Hermes, and Kimi received static portability assessment only; the validation report states that limitation.

## Source history

Source Gist: [`dachent/cdc05151d047708c290bd4da0aaeed96`](https://gist.github.com/dachent/cdc05151d047708c290bd4da0aaeed96), description `Deep planning general prompt template`.

The repository keeps the scaffold at one stable Markdown path from its first imported revision. Seven ordered commits replay the complete Gist revision sequence; five contain distinct scaffold contents, one records a filename-only change, and one records companion-file changes. Each commit includes the source revision, timestamp, original filename, and SHA-256 as commit trailers.

| Source revision | Committed at (UTC) | Original scaffold filename | Scaffold SHA-256 |
| --- | --- | --- | --- |
| `75978d8fd61ad9262d182bb7f29b09742c3e9d84` | 2026-06-10 17:17:26 | `deep_planning.txt` | `7A254F0898148379769773DA025029000DC068A1FCF6643F4CE5BBC255C80EB2` |
| `aed37033cb04897aefd9281f93f9fff82f9a98e8` | 2026-06-10 19:33:37 | `deep_planning.txt` | `D2B74200254818A5373E9DADF2254C755BDCB967B27FC87D96B418C751356B34` |
| `e9579a6184a2277a946c7632114e5a664ebddbd9` | 2026-06-10 20:15:36 | `deep_planning.txt` | `0CFCDC62B3911FDC95C6C28CEE86E1A53C53176EFC61B2AA9FFE6D45777C4865` |
| `6ea4c02e5aa60c9991e1e4d1c50089c01cd6ec83` | 2026-06-13 20:19:36 | `claude_code_deep_planning.txt` | `0CFCDC62B3911FDC95C6C28CEE86E1A53C53176EFC61B2AA9FFE6D45777C4865` |
| `ddeb80cea25ff158f9264a8d7abe4016b9c12e36` | 2026-07-16 20:56:06 | `claude_code_deep_planning.txt` | `EC9E91D9552809BF2BF3E16F1262E126B09285087090EABCCC269A84DA5B1672` |
| `93db3febb8eefb4b65e049bbb36a9ae70fc14fec` | 2026-07-26 19:44:46 | `claude_code_deep_planning.md` | `B6113AAFEAD3856EFB1485039C4F801A8DFD30B41B5C0715C16B99C6B1921AC2` |
| `ef2adb8dcb702eb39c0888cb3e455c7cc40c977d` | 2026-07-26 19:57:56 | `claude_code_deep_planning.md` | `B6113AAFEAD3856EFB1485039C4F801A8DFD30B41B5C0715C16B99C6B1921AC2` |

The final companion-file hashes are:

- `scaffold_evaluationplan.md`: `5CDBEC3D47510779984D275E774CA055CDC18B79016292CF59297F90D2573FE6`
- `validation_results_codex.md`: `66B0BBB9F1F9960286D8E12486F6310A312CBC16B86DCD1F256C52B02753C97D`
