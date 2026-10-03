"""Creation of independent LaTeX consumer projects."""

from __future__ import annotations

import json
from pathlib import Path
import re
from typing import Mapping

from .runtime import (
    CLI_VERSION,
    GENERATED_MARKER_NAME,
    PROJECT_CONFIG_NAME,
    CliError,
    is_inside,
    read_utf8,
    resource_root,
    schema_path,
    write_utf8,
)

DOCUMENT_TYPES = ("thesis", "proposal", "literature-review")
DISCIPLINES = ("humanities", "science")


def escape_tex(value: str) -> str:
    replacements = {
        "\\": r"\textbackslash{}", "{": r"\{", "}": r"\}", "%": r"\%",
        "&": r"\&", "#": r"\#", "_": r"\_", "^": r"\textasciicircum{}",
        "~": r"\textasciitilde{}",
    }
    return "".join(replacements.get(char, char) for char in value)


def _value(raw: str | None, placeholder: str = "") -> str:
    return placeholder if raw is None else escape_tex(raw)


def _metadata_path_value(raw: str | None) -> str:
    return "" if raw is None else escape_tex(raw.replace("\\", "/"))


def project_config(root: Path, document_type: str) -> dict[str, object]:
    stem = {
        "proposal": "proposal",
        "literature-review": "literature-review",
        "thesis": "thesis",
    }[document_type]
    return {
        "$schema": schema_path(root), "schemaVersion": 1, "main": "latex/main.tex",
        "outputRoot": "build", "checkOutputRoot": "build/checks", "citationMode": "linked",
        "wordRefresh": "auto", "deliverables": {
            "pdf": f"build/deliverables/{stem}.pdf",
            "word": f"build/deliverables/{stem}.docx",
        },
    }


def metadata_tex(args, document_type: str) -> str:
    year = "2022" if document_type == "proposal" else "2024"
    return rf"""% UTF-8 metadata for the generated external project.
\yibinsetup{{
  template-year           = {year},
  title                   = {{{_value(args.title, "（填写中文题目）")}}},
  english-title           = {{{_value(args.english_title, "Title in English")}}},
  author                  = {{{_value(args.author, "（填写姓名）")}}},
  student-id              = {{{_value(args.student_id, "（填写学号）")}}},
  college                 = {{{_value(args.college, "（填写学院）")}}},
  major                   = {{{_value(args.major, "（填写专业）")}}},
  grade                   = {{{_value(args.grade, "（填写年级）")}}},
  class-name              = {{{_value(args.class_name)}}},
  advisor                 = {{{_value(args.advisor, "（填写校内导师）")}}},
  advisor-title           = {{{_value(args.advisor_title, "（填写职称）")}}},
  external-advisor        = {{{_value(args.external_advisor)}}},
  external-advisor-title  = {{{_value(args.external_advisor_title)}}},
  version                 = {{{_value(args.version)}}},
  date                    = {{{_value(args.date)}}},
  secrecy                 = {args.secrecy},
  declassify-year         = {{{_value(args.declassify_year)}}},
  logo                    = {{builtin}},
  author-signature        = {{{_metadata_path_value(args.author_signature)}}},
  advisor-signature       = {{{_metadata_path_value(args.advisor_signature)}}},
  signature-background    = {args.signature_background}
}}
"""


def _main_tex(document_type: str, discipline: str) -> str:
    if document_type == "proposal":
        return rf"""% !TeX program = xelatex
\documentclass[proposal,{discipline}]{{yibinthesis}}
\input{{metadata}}
\addbibresource{{references.bib}}
\begin{{document}}
\makeyibinproposal
\yibinproposalfield{{significance}}{{请填写选题意义。}}
\yibinproposalfield{{research-status}}{{请填写国内外研究现状概述。}}
\yibinproposalfield{{research-content}}{{请填写主要研究内容。}}
\yibinproposalfield{{research-approach}}{{请填写研究思路、方法、技术路线与可行性论证。}}
\yibinproposalfield{{schedule}}{{请填写研究工作安排及进度。}}
\yibinproposalfield{{references}}{{\printyibinproposalbibliography}}
\yibinproposalfield{{advisor-opinion}}{{}}
\end{{document}}
"""
    if document_type == "literature-review":
        return rf"""% !TeX program = xelatex
\documentclass[literature-review,{discipline}]{{yibinthesis}}
\input{{metadata}}
\addbibresource{{references.bib}}
\begin{{document}}
\makeyibinliteraturereviewcover
\input{{chapters/body}}
\printyibinbibliography
\end{{document}}
"""
    return rf"""% !TeX program = xelatex
\documentclass[thesis,{discipline}]{{yibinthesis}}
\input{{metadata}}
\addbibresource{{references.bib}}
\begin{{document}}
\makeyibincover
\makeyibindeclarations
\frontmatter
\input{{chapters/00-abstract-cn}}
\input{{chapters/01-abstract-en}}
\clearpage
\tableofcontents
\mainmatter
\input{{chapters/10-introduction}}
\input{{chapters/20-body}}
\input{{chapters/30-conclusion}}
\backmatter
\input{{chapters/99-acknowledgements}}
\printyibinnotes
\printyibinbibliography
\input{{chapters/90-appendix}}
\end{{document}}
"""


def _body_files(document_type: str) -> Mapping[str, str]:
    if document_type == "literature-review":
        return {"chapters/body.tex": "%% 文献综述正文（一般不少于 3000 字）\n请在此填写综述正文；引言部分可以不设标题。\n"}
    if document_type == "proposal":
        return {}
    return {
        "chapters/00-abstract-cn.tex": r"\begin{cnabstract}" "\n请在此填写中文摘要（约 250 字）。\n\n" r"\cnkeywords{关键词一；关键词二；关键词三}" "\n" r"\end{cnabstract}" "\n",
        "chapters/01-abstract-en.tex": r"\begin{enabstract}" "\nPlease write the English abstract here.\n\n" r"\enkeywords{keyword one; keyword two; keyword three}" "\n" r"\end{enabstract}" "\n",
        "chapters/10-introduction.tex": r"\yibinopeningchapter" "\n" r"\section{研究背景}" "\n请在此填写绪论内容。\n" r"\section{研究内容与方法}" "\n请在此填写研究内容与方法。\n",
        "chapters/20-body.tex": r"\chapter{正文}" "\n请在此填写论文主体内容。\n",
        "chapters/30-conclusion.tex": r"\yibinclosingchapter" "\n请在此填写结论与展望。\n",
        "chapters/90-appendix.tex": r"\begin{yibinappendices}" "\n% 需要附录时取消注释并填写内容。\n" r"% \yibinappendix{材料清单}" "\n" r"% \yibinappendixsection{说明}" "\n" r"\end{yibinappendices}" "\n",
        "chapters/99-acknowledgements.tex": r"\yibinunnumberedchapter{致谢}" "\n请在此填写致谢。\n",
    }


def scaffold_files(args) -> dict[str, str]:
    files = {
        "latex/main.tex": _main_tex(args.document_type, args.discipline),
        "latex/metadata.tex": metadata_tex(args, args.document_type),
        "latex/references.bib": "% Add verified references in GB/T 7714-2015 BibLaTeX format.\n",
        "README.md": f"# 外部论文项目\n\n这是由 YibinThesis CLI 创建的 {args.document_type} LaTeX 项目。\n\n在本目录执行 yibinthesis check 或 yibinthesis build all。论文源文件位于 latex/。\n",
        ".gitignore": "build/\n*.aux\n*.bbl\n*.bcf\n*.blg\n*.log\n*.run.xml\n*.synctex.gz\n",
        "latex/assets/.gitkeep": "",
    }
    files.update({f"latex/{name}": content for name, content in _body_files(args.document_type).items()})
    return files


def _marker(path: Path) -> dict[str, object] | None:
    if not path.is_file():
        return None
    value = json.loads(read_utf8(path))
    return value if isinstance(value, dict) else None


def create_project(args) -> int:
    target = Path(args.target).expanduser().resolve()
    if is_inside(target, resource_root().resolve()):
        raise CliError(f"论文项目必须创建在工具仓库之外：{target}")
    marker_path = target / GENERATED_MARKER_NAME
    marker = _marker(marker_path)
    owned = set(marker.get("files", [])) if marker else set()
    if target.exists() and any(target.iterdir()) and (not args.force or not owned):
        raise CliError(f"目标目录非空，未覆盖已有内容：{target}；仅可用 --force 覆盖 CLI 自己生成的文件")
    files = scaffold_files(args)
    planned = [target / name for name in sorted(files)] + [target / PROJECT_CONFIG_NAME, marker_path]
    if args.dry_run:
        print("将创建或更新：")
        for path in planned:
            print(f"  {path}")
        return 0
    target.mkdir(parents=True, exist_ok=True)
    generated = []
    for relative, content in files.items():
        destination = target / relative
        name = destination.relative_to(target).as_posix()
        if destination.exists() and name not in owned:
            raise CliError(f"目标文件已存在，拒绝覆盖：{destination}")
        write_utf8(destination, content)
        generated.append(name)
    config = target / PROJECT_CONFIG_NAME
    if config.exists() and PROJECT_CONFIG_NAME not in owned:
        raise CliError(f"目标文件已存在，拒绝覆盖：{config}")
    write_utf8(config, json.dumps(project_config(target, args.document_type), ensure_ascii=False, indent=2) + "\n")
    generated.append(PROJECT_CONFIG_NAME)
    write_utf8(marker_path, json.dumps({"tool": "yibinthesis", "version": CLI_VERSION, "documentType": args.document_type, "discipline": args.discipline, "files": sorted(generated)}, ensure_ascii=False, indent=2) + "\n")
    print(f"已创建 LaTeX 项目：{target}\n入口文件：{target / 'latex' / 'main.tex'}")
    return 0
