"""Argument parser and top-level command dispatch."""

from __future__ import annotations

import argparse
import sys
from typing import Sequence

from . import build
from .runtime import CLI_VERSION, CliError, configure_stdio
from .scaffold import DISCIPLINES, DOCUMENT_TYPES, create_project


def _project_options(parser: argparse.ArgumentParser) -> None:
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--project", help="外部 LaTeX 项目目录（默认：当前目录）")
    group.add_argument("--config", help="显式指定 yibinthesis.project.json")


def _build_options(parser: argparse.ArgumentParser) -> None:
    _project_options(parser)
    parser.add_argument("--main", help="覆盖配置中的 LaTeX 入口")
    parser.add_argument("--output-root", help="覆盖构建输出目录")
    parser.add_argument("--citation-mode", choices=("linked", "native"), help="Word 引用模式")
    parser.add_argument("--word-refresh", choices=("auto", "always", "never"), help="Word 域刷新策略")


def _new_options(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("target", help="项目目录（必须位于工具仓库之外）")
    parser.add_argument("--type", choices=DOCUMENT_TYPES, default="thesis", dest="document_type")
    parser.add_argument("--discipline", choices=DISCIPLINES, default="humanities")
    for option in ("title", "english-title", "author", "student-id", "college", "major", "grade", "class-name", "advisor", "advisor-title", "external-advisor", "external-advisor-title", "version", "date", "declassify-year", "author-signature", "advisor-signature"):
        parser.add_argument(f"--{option}", dest=option.replace("-", "_"))
    parser.add_argument("--secrecy", choices=("public", "confidential"), default="public")
    parser.add_argument("--signature-background", choices=("preserve", "whiten"), default="preserve")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--force", action="store_true")


def make_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="yibinthesis", description="宜宾学院本科毕业材料 LaTeX 工具")
    parser.add_argument("--version", action="version", version=f"%(prog)s {CLI_VERSION}")
    subparsers = parser.add_subparsers(dest="command", required=True)
    new = subparsers.add_parser("new", aliases=["init"], help="新建外部 LaTeX 项目结构")
    _new_options(new)
    build_parser = subparsers.add_parser("build", help="构建外部项目（默认同时生成 PDF 和 DOCX）")
    build_parser.add_argument("format", choices=build.BUILD_FORMATS, nargs="?", default="all")
    _build_options(build_parser)
    for operation, help_text in (("doctor", "检查外部工具链"), ("check", "运行结构和格式检查"), ("clean", "清理外部项目构建产物")):
        item = subparsers.add_parser(operation, help=help_text)
        _build_options(item)
    for format_name in build.BUILD_FORMATS:
        item = subparsers.add_parser(format_name, help=argparse.SUPPRESS)
        _build_options(item)
    subparsers.add_parser("version", help="显示 CLI 版本")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    configure_stdio()
    args = make_parser().parse_args(argv)
    try:
        if args.command in {"new", "init"}:
            return create_project(args)
        if args.command == "version":
            print(f"YibinThesis CLI {CLI_VERSION}")
            return 0
        if args.command == "build":
            return build.run(args, "build")
        if args.command in build.BUILD_FORMATS:
            args.format = args.command
            return build.run(args, "build")
        return build.run(args, args.command)
    except CliError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2
    except OSError as exc:
        print(f"错误：无法启动构建器 ({exc})", file=sys.stderr)
        return 2
    except KeyboardInterrupt:
        print("已取消。", file=sys.stderr)
        return 130
