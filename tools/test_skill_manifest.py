from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from validate_skill_manifest import validate_manifest


class ManifestValidatorTest(unittest.TestCase):
    def make_repo(self, root: Path) -> Path:
        skill = root / "sample-skill"
        (skill / "agents").mkdir(parents=True)
        (skill / "SKILL.md").write_text(
            "---\nname: sample-skill\ndescription: Sample.\n---\n\n# Sample\n",
            encoding="utf-8",
        )
        (skill / "agents" / "openai.yaml").write_text(
            "interface:\n  display_name: Sample\n  short_description: Sample skill\n  default_prompt: Use $sample-skill.\n",
            encoding="utf-8",
        )
        (root / ".shared" / "runtime").mkdir(parents=True)
        manifest = {
            "schema_version": 1,
            "repository": "example/repo",
            "policy": {
                "supported_statuses": ["supported", "archived"],
                "source_classifications": ["repo-owned-original"],
                "required_packaging_for_supported": ["skill_file", "agent_metadata"],
                "catalog_groups": [{"key": "test", "title": "Test", "description": "Test skills."}],
            },
            "generated_mirrors": [],
            "shared_components": [
                {"name": "runtime", "path": ".shared/runtime", "consumers": ["sample-skill"]}
            ],
            "skills": [
                {
                    "name": "sample-skill",
                    "path": "sample-skill",
                    "family": "test",
                    "catalog_group": "test",
                    "status": "supported",
                    "description": "Sample.",
                    "platforms": ["linux"],
                    "agents": ["codex"],
                    "runtimes": {"type": "prompt"},
                    "packaging": {
                        "skill_file": "sample-skill/SKILL.md",
                        "agent_metadata": "sample-skill/agents/openai.yaml",
                    },
                    "source": {"classification": "repo-owned-original"},
                    "validation": {"hosted_commands": [], "environment_dependent_commands": []},
                    "owner": "@owner",
                    "last_reviewed": "2026-07-11",
                }
            ],
        }
        path = root / "skills-manifest.json"
        path.write_text(json.dumps(manifest), encoding="utf-8")
        return path

    def test_valid_repository(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            manifest = self.make_repo(root)
            self.assertEqual(validate_manifest(root, manifest), [])

    def test_unregistered_skill_fails(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            manifest = self.make_repo(root)
            extra = root / "unregistered"
            extra.mkdir()
            (extra / "SKILL.md").write_text(
                "---\nname: unregistered\ndescription: X\n---\n", encoding="utf-8"
            )
            errors = validate_manifest(root, manifest)
            self.assertTrue(any("unregistered top-level skill directory" in item for item in errors))

    def test_missing_metadata_fails(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            manifest = self.make_repo(root)
            (root / "sample-skill" / "agents" / "openai.yaml").unlink()
            errors = validate_manifest(root, manifest)
            self.assertTrue(any("missing agent metadata" in item for item in errors))

    def test_archived_skill_uses_registered_archive_path(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            manifest_path = self.make_repo(root)
            archived = root / "archive" / "retired-skill"
            (archived / "agents").mkdir(parents=True)
            (root / "archive" / "README.md").write_text("# Archive\n", encoding="utf-8")
            (archived / "README.md").write_text("# Retired\n", encoding="utf-8")
            (archived / "SKILL.md").write_text(
                "---\nname: retired-skill\ndescription: Archived.\n---\n", encoding="utf-8"
            )
            (archived / "agents" / "openai.yaml").write_text(
                "interface:\n  display_name: Retired\n  short_description: Archived skill\n"
                "  default_prompt: Do not use.\n",
                encoding="utf-8",
            )
            data = json.loads(manifest_path.read_text(encoding="utf-8"))
            data["skills"].append(
                {
                    "name": "retired-skill",
                    "path": "archive/retired-skill",
                    "family": "test",
                    "catalog_group": "test",
                    "status": "archived",
                    "description": "Archived.",
                    "platforms": ["linux"],
                    "agents": ["codex"],
                    "runtimes": {"type": "prompt"},
                    "packaging": {
                        "skill_file": "archive/retired-skill/SKILL.md",
                        "agent_metadata": "archive/retired-skill/agents/openai.yaml",
                    },
                    "source": {"classification": "repo-owned-original"},
                    "validation": {
                        "hosted_commands": [],
                        "environment_dependent_commands": [],
                    },
                    "owner": "@owner",
                    "last_reviewed": "2026-07-27",
                }
            )
            manifest_path.write_text(json.dumps(data), encoding="utf-8")
            self.assertEqual(validate_manifest(root, manifest_path), [])

    def test_active_skill_cannot_use_nested_path(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            manifest_path = self.make_repo(root)
            data = json.loads(manifest_path.read_text(encoding="utf-8"))
            data["skills"][0]["path"] = "nested/sample-skill"
            manifest_path.write_text(json.dumps(data), encoding="utf-8")
            errors = validate_manifest(root, manifest_path)
            self.assertTrue(any("active skill path must be one top-level directory" in item for item in errors))

    def test_unregistered_archived_skill_fails(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            manifest_path = self.make_repo(root)
            archived = root / "archive" / "orphan"
            archived.mkdir(parents=True)
            (root / "archive" / "README.md").write_text("# Archive\n", encoding="utf-8")
            (archived / "SKILL.md").write_text(
                "---\nname: orphan\ndescription: Orphan.\n---\n", encoding="utf-8"
            )
            errors = validate_manifest(root, manifest_path)
            self.assertTrue(any("unregistered archived skill directory" in item for item in errors))

    def test_documented_scaffold_namespace_is_not_a_skill(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            manifest_path = self.make_repo(root)
            package = root / "scaffolds" / "sample-scaffold"
            package.mkdir(parents=True)
            (root / "scaffolds" / "README.md").write_text("# Scaffolds\n", encoding="utf-8")
            (package / "README.md").write_text("# Sample\n", encoding="utf-8")
            (package / "prompt.md").write_text("Example scaffold.\n", encoding="utf-8")
            self.assertEqual(validate_manifest(root, manifest_path), [])

    def test_scaffold_cannot_masquerade_as_installable_skill(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            manifest_path = self.make_repo(root)
            package = root / "scaffolds" / "sample-scaffold"
            package.mkdir(parents=True)
            (root / "scaffolds" / "README.md").write_text("# Scaffolds\n", encoding="utf-8")
            (package / "README.md").write_text("# Sample\n", encoding="utf-8")
            (package / "SKILL.md").write_text(
                "---\nname: sample-scaffold\ndescription: Wrong.\n---\n", encoding="utf-8"
            )
            errors = validate_manifest(root, manifest_path)
            self.assertTrue(any("scaffold packages must not contain SKILL.md" in item for item in errors))


if __name__ == "__main__":
    unittest.main(verbosity=2)
