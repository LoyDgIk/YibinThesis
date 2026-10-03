from __future__ import annotations

import json
from pathlib import Path
import shutil
import sys
import subprocess
import tempfile
import unittest
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "lib"))

import yibinthesis_cli as cli  # noqa: E402
from yibinthesis_cli import build as cli_build  # noqa: E402


class CliTests(unittest.TestCase):
    def test_new_creates_external_utf8_project(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            project = Path(temporary) / "我的论文"
            result = cli.main(
                [
                    "new",
                    str(project),
                    "--type",
                    "thesis",
                    "--discipline",
                    "science",
                    "--title",
                    "中文题目",
                    "--author",
                    "张三",
                    "--grade",
                    "2023",
                ]
            )
            self.assertEqual(result, 0)
            main = project / "latex" / "main.tex"
            metadata = project / "latex" / "metadata.tex"
            self.assertTrue(main.is_file())
            self.assertIn(r"\documentclass[thesis,science]{yibinthesis}", main.read_text(encoding="utf-8"))
            self.assertIn("中文题目", metadata.read_text(encoding="utf-8"))
            config = json.loads((project / "yibinthesis.project.json").read_text(encoding="utf-8"))
            self.assertEqual(config["main"], "latex/main.tex")
            self.assertEqual(config["schemaVersion"], 1)

    def test_new_rejects_path_inside_tool_repository(self) -> None:
        target = cli.resource_root() / "tests" / "should-not-be-created"
        self.assertEqual(cli.main(["new", str(target)]), 2)
        self.assertFalse(target.exists())

    def test_dry_run_does_not_write(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            project = Path(temporary) / "dry-run"
            self.assertEqual(cli.main(["new", str(project), "--dry-run"]), 0)
            self.assertFalse(project.exists())

    def test_build_defaults_to_all_and_uses_absolute_config(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            project = Path(temporary) / "project"
            project.mkdir()
            (project / "yibinthesis.project.json").write_text(
                '{"schemaVersion": 1, "main": "latex/main.tex"}\n',
                encoding="utf-8",
            )
            with patch.object(cli_build, "_powershell", return_value="pwsh"), patch.object(
                cli_build.subprocess, "run", return_value=type("Result", (), {"returncode": 0})()
            ) as run:
                self.assertEqual(cli.main(["build", "--project", str(project)]), 0)
            command = run.call_args.args[0]
            self.assertIn("all", command)
            self.assertIn(str(project / "yibinthesis.project.json"), command)
            self.assertEqual(run.call_args.kwargs["cwd"], project.resolve())

    def test_init_alias_uses_same_scaffold(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            project = Path(temporary) / "proposal"
            self.assertEqual(cli.main(["init", str(project), "--type", "proposal"]), 0)
            main = (project / "latex" / "main.tex").read_text(encoding="utf-8")
            metadata = (project / "latex" / "metadata.tex").read_text(encoding="utf-8")
            self.assertIn(r"\documentclass[proposal,humanities]{yibinthesis}", main)
            self.assertIn("template-year           = 2022", metadata)

    def test_build_expands_metadata_in_deliverable_name(self) -> None:
        powershell = shutil.which("pwsh") or shutil.which("powershell")
        if powershell is None:
            self.skipTest("PowerShell is required for build.ps1 integration coverage")
        with tempfile.TemporaryDirectory() as temporary:
            project = Path(temporary) / "project"
            self.assertEqual(
                cli.main(["new", str(project), "--title", "Title", "--author", "Author"]),
                0,
            )
            config_path = project / "yibinthesis.project.json"
            config = json.loads(config_path.read_text(encoding="utf-8"))
            config["deliverables"]["pdf"] = "build/deliverables/{{title}}.txt"
            config_path.write_text(
                json.dumps(config, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
            )
            result = subprocess.run(
                [
                    powershell,
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(ROOT / "build.ps1"),
                    "clean",
                    "-Config",
                    str(config_path),
                ],
                cwd=project,
                check=False,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            output = result.stdout + result.stderr
            self.assertNotEqual(result.returncode, 0)
            self.assertIn("Title.txt", output)


if __name__ == "__main__":
    unittest.main()
