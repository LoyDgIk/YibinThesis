from __future__ import annotations

import json
from pathlib import Path
import re
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

    def test_new_review_creates_bilingual_abstracts_and_untitled_introduction(self) -> None:
        for discipline in ("humanities", "science"):
            with self.subTest(discipline=discipline), tempfile.TemporaryDirectory() as temporary:
                project = Path(temporary) / "文献综述"
                self.assertEqual(
                    cli.main(["new", str(project), "--type", "literature-review", "--discipline", discipline]),
                    0,
                )
                latex = project / "latex"
                main = (latex / "main.tex").read_text(encoding="utf-8")
                self.assertIn(rf"\documentclass[literature-review,{discipline}]{{yibinthesis}}", main)
                self.assertEqual(
                    re.findall(r"\\input\{([^}]+)\}", main),
                    ["metadata", "chapters/00-abstract-cn", "chapters/01-abstract-en", "chapters/body"],
                )
                self.assertLess(main.index(r"\makeyibinliteraturereviewcover"), main.index(r"\input{chapters/00-abstract-cn}"))
                self.assertLess(main.index(r"\input{chapters/body}"), main.index(r"\printyibinbibliography"))
                self.assertNotIn(r"\tableofcontents", main)
                for suffix, environment, keywords in (
                    ("00-abstract-cn", "cnabstract", "cnkeywords"),
                    ("01-abstract-en", "enabstract", "enkeywords"),
                ):
                    abstract = (latex / "chapters" / f"{suffix}.tex").read_text(encoding="utf-8")
                    self.assertIn(rf"\begin{{{environment}}}", abstract)
                    self.assertIn(rf"\end{{{environment}}}", abstract)
                    self.assertIn(rf"\{keywords}{{", abstract)
                body = (latex / "chapters" / "body.tex").read_text(encoding="utf-8")
                first_paragraph = next(line for line in body.splitlines() if line and not line.startswith("%"))
                self.assertFalse(first_paragraph.startswith("\\"))
                self.assertNotRegex(body, r"\\(?:chapter|section|yibinunnumberedchapter)\*?\{引言\}")
                self.assertNotIn(r"\yibinclosingchapter", body)
                chapters = re.findall(r"\\chapter\{([^}]+)\}", body)
                self.assertEqual(chapters[-1], "结论与展望")
                self.assertGreaterEqual(len(chapters), 3)
                self.assertNotIn(r"\clearpage", body)
                metadata = (latex / "metadata.tex").read_text(encoding="utf-8")
                for field in ("second-advisor", "second-advisor-title", "date"):
                    self.assertRegex(metadata, rf"(?m)^\s*{field}\s*=\s*\{{\}}")
                self.assertIn("（填写所在地）", metadata)
                self.assertIn("（填写邮编）", metadata)
                self.assertNotRegex((latex / "references.bib").read_text(encoding="utf-8"), r"(?m)^\s*@")

    def test_new_thesis_introduction_has_the_five_template_sections(self) -> None:
        for discipline in ("humanities", "science"):
            with self.subTest(discipline=discipline), tempfile.TemporaryDirectory() as temporary:
                project = Path(temporary) / "论文"
                self.assertEqual(cli.main(["new", str(project), "--discipline", discipline]), 0)
                introduction = (project / "latex" / "chapters" / "10-introduction.tex").read_text(encoding="utf-8")
                self.assertTrue(introduction.startswith(r"\yibinopeningchapter"))
                self.assertEqual(
                    re.findall(r"\\section\{([^}]+)\}", introduction),
                    ["研究的背景", "国内外研究现状", "选题的目的", "课题研究的意义", "研究的内容及方法"],
                )

    def test_second_advisor_option_keeps_external_keys_for_other_document_types(self) -> None:
        for document_type in ("thesis", "proposal"):
            with self.subTest(document_type=document_type), tempfile.TemporaryDirectory() as temporary:
                project = Path(temporary) / document_type
                self.assertEqual(
                    cli.main(["new", str(project), "--type", document_type, "--second-advisor", "测试导师", "--second-advisor-title", "测试职称"]),
                    0,
                )
                metadata = (project / "latex" / "metadata.tex").read_text(encoding="utf-8")
                self.assertRegex(metadata, r"(?m)^\s*external-advisor\s*=\s*\{测试导师\}")
                self.assertRegex(metadata, r"(?m)^\s*external-advisor-title\s*=\s*\{测试职称\}")
                self.assertNotRegex(metadata, r"(?m)^\s*second-advisor(?:-title)?\s*=")

    def test_review_metadata_roundtrips_through_build_config_with_both_advisor_options(self) -> None:
        powershell = shutil.which("pwsh") or shutil.which("powershell")
        if powershell is None:
            self.skipTest("PowerShell is required for build.ps1 integration coverage")
        for discipline, advisor_option in (("humanities", "second-advisor"), ("science", "external-advisor")):
            with self.subTest(discipline=discipline, advisor_option=advisor_option), tempfile.TemporaryDirectory() as temporary:
                project = Path(temporary) / "综述项目"
                self.assertEqual(
                    cli.main([
                        "new", str(project), "--type", "literature-review", "--discipline", discipline,
                        "--college", "测试学部&中心", "--affiliation-city", "测试市", "--affiliation-postcode", "064400",
                        f"--{advisor_option}", "测试导师_乙", f"--{advisor_option}-title", "测试职称",
                    ]),
                    0,
                )
                metadata = (project / "latex" / "metadata.tex").read_text(encoding="utf-8")
                self.assertNotRegex(metadata, r"(?m)^\s*external-advisor(?:-title)?\s*=")
                self.assertIn(r"{测试学部\&中心}", metadata)
                config_path = project / "yibinthesis.project.json"
                config = json.loads(config_path.read_text(encoding="utf-8"))
                config["deliverables"]["pdf"] = (
                    "build/deliverables/{{college}}-{{affiliation-city}}-{{affiliation-postcode}}-"
                    "{{second-advisor}}-{{second-advisor-title}}.txt"
                )
                config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
                # An invalid extension exposes the expanded filename before any toolchain or build runs.
                result = subprocess.run(
                    [powershell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ROOT / "build.ps1"), "clean", "-Config", str(config_path)],
                    cwd=project, check=False, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30,
                )
                self.assertNotEqual(result.returncode, 0)
                self.assertIn("测试学部-中心-测试市-064400-测试导师_乙-测试职称.txt", result.stdout + result.stderr)

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

    def test_new_proposal_has_semantic_headings_in_two_fields(self) -> None:
        for discipline in ("humanities", "science"):
            with self.subTest(discipline=discipline), tempfile.TemporaryDirectory() as temporary:
                project = Path(temporary) / "开题报告"
                self.assertEqual(
                    cli.main(["new", str(project), "--type", "proposal", "--discipline", discipline]),
                    0,
                )
                main = (project / "latex" / "main.tex").read_text(encoding="utf-8")
                parts = re.split(r"\\yibinproposalfield\{([^}]+)\}", main)
                fields = dict(zip(parts[1::2], parts[2::2]))
                self.assertEqual(
                    list(fields),
                    ["significance", "research-status", "research-content", "research-approach", "schedule", "references", "advisor-opinion"],
                )
                for name in ("significance", "research-status"):
                    self.assertRegex(fields[name], r"^\{\s*\\section\{[^}]+\}")
                self.assertIn(r"\subsection{", fields["research-status"])
                for name in ("research-content", "research-approach", "schedule", "references", "advisor-opinion"):
                    self.assertNotRegex(fields[name], r"\\(?:section|subsection|subsubsection|paragraph)\{")
                self.assertNotIn("（1）", main)
                self.assertNotIn("(1)", main)

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
