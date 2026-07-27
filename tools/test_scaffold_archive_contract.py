from __future__ import annotations

import hashlib
import json
import re
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
PACKAGE = ROOT / "scaffolds" / "agent-project-scaffold"
EXPECTED_HASHES = {
    "claude_code_deep_planning.md": "B6113AAFEAD3856EFB1485039C4F801A8DFD30B41B5C0715C16B99C6B1921AC2",
    "scaffold_evaluationplan.md": "5CDBEC3D47510779984D275E774CA055CDC18B79016292CF59297F90D2573FE6",
    "validation_results_codex.md": "66B0BBB9F1F9960286D8E12486F6310A312CBC16B86DCD1F256C52B02753C97D",
}
GIST_REVISIONS = {
    "75978d8fd61ad9262d182bb7f29b09742c3e9d84",
    "aed37033cb04897aefd9281f93f9fff82f9a98e8",
    "e9579a6184a2277a946c7632114e5a664ebddbd9",
    "6ea4c02e5aa60c9991e1e4d1c50089c01cd6ec83",
    "ddeb80cea25ff158f9264a8d7abe4016b9c12e36",
    "93db3febb8eefb4b65e049bbb36a9ae70fc14fec",
    "ef2adb8dcb702eb39c0888cb3e455c7cc40c977d",
}


class ScaffoldArchiveContractTest(unittest.TestCase):
    def test_published_files_are_byte_exact(self):
        self.assertEqual(
            {path.name for path in PACKAGE.iterdir() if path.is_file()},
            {"README.md", *EXPECTED_HASHES},
        )
        for name, expected in EXPECTED_HASHES.items():
            actual = hashlib.sha256((PACKAGE / name).read_bytes()).hexdigest().upper()
            self.assertEqual(actual, expected, name)

    def test_scaffold_word_count_and_non_skill_boundary(self):
        text = (PACKAGE / "claude_code_deep_planning.md").read_text(encoding="utf-8")
        self.assertGreaterEqual(len(text.split()), 750)
        self.assertLessEqual(len(text.split()), 850)
        self.assertFalse(any(PACKAGE.rglob("SKILL.md")))
        self.assertFalse(any(path.parent.name == "agents" for path in PACKAGE.rglob("openai.yaml")))
        readme = (PACKAGE / "README.md").read_text(encoding="utf-8")
        self.assertIn("not a skill", readme.lower())
        self.assertIn("do not install", readme.lower())

    def test_readmes_have_resolvable_local_links(self):
        readmes = [
            ROOT / "scaffolds" / "README.md",
            PACKAGE / "README.md",
            ROOT / "archive" / "README.md",
        ]
        for readme in readmes:
            text = readme.read_text(encoding="utf-8")
            for target in re.findall(r"\[[^]]+\]\(([^)]+)\)", text):
                if "://" in target or target.startswith("#"):
                    continue
                self.assertTrue((readme.parent / target).resolve().exists(), f"{readme}: {target}")

    def test_all_gist_revisions_are_documented(self):
        readme = (PACKAGE / "README.md").read_text(encoding="utf-8")
        for revision in GIST_REVISIONS:
            self.assertIn(revision, readme)
        self.assertNotIn("zzz_", "\n".join(path.name for path in PACKAGE.iterdir()))

    def test_orchestrator_is_archived_and_excluded_from_catalog(self):
        manifest = json.loads((ROOT / "skills-manifest.json").read_text(encoding="utf-8"))
        item = next(skill for skill in manifest["skills"] if skill["name"] == "agent-project-orchestrator")
        self.assertEqual(item["status"], "archived")
        self.assertEqual(item["path"], "archive/agent-project-orchestrator")
        self.assertFalse(item["validation"]["ci_enabled"])
        self.assertEqual(item["validation"]["hosted_commands"], [])

        readme = (ROOT / "README.md").read_text(encoding="utf-8")
        for section in (
            "skill-catalog",
            "installation-inventory",
            "platform-agent-matrix",
            "validation-summary",
        ):
            start = f"<!-- BEGIN GENERATED: {section} -->"
            end = f"<!-- END GENERATED: {section} -->"
            body = readme.split(start, 1)[1].split(end, 1)[0]
            self.assertNotIn("agent-project-orchestrator", body)


if __name__ == "__main__":
    unittest.main(verbosity=2)
