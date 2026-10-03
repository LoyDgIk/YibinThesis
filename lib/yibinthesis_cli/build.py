"""PowerShell builder adapter used by the CLI."""

from __future__ import annotations

from pathlib import Path
import shutil
import subprocess

from .runtime import PROJECT_CONFIG_NAME, CliError, resource_root

BUILD_FORMATS = ("pdf", "word", "all")


def project_and_config(args) -> tuple[Path, Path]:
    if args.project and args.config:
        raise CliError("--project 与 --config 不能同时使用。")
    if args.config:
        # Keep the caller's absolute spelling. Windows CI may expose the same
        # temporary directory through an 8.3 alias; canonicalizing here would
        # make command inspection disagree with the path supplied by callers.
        config = Path(args.config).expanduser().absolute()
        project = config.parent
    else:
        project = Path(args.project or ".").expanduser().resolve()
        config = project / PROJECT_CONFIG_NAME
    if not project.is_dir():
        raise CliError(f"项目目录不存在：{project}")
    if not config.is_file():
        raise CliError(f"找不到项目配置：{config}；请先运行 yibinthesis new <目录>。")
    return project, config


def _powershell() -> str:
    for candidate in ("pwsh", "powershell"):
        found = shutil.which(candidate)
        if found:
            return found
    raise CliError("未找到 PowerShell（pwsh 或 powershell），无法调用构建器。")


def command_args(args, operation: str) -> tuple[list[str], Path]:
    project, config = project_and_config(args)
    builder = resource_root() / "build.ps1"
    if not builder.is_file():
        raise CliError(f"找不到内置构建器：{builder}")
    command = args.format if operation == "build" else operation
    command_line = [_powershell(), "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(builder), command, "-Config", str(config)]
    if args.main:
        main = Path(args.main).expanduser()
        command_line.extend(("-Main", str((project / main if not main.is_absolute() else main).resolve())))
    if args.output_root:
        output = Path(args.output_root).expanduser()
        command_line.extend(("-OutputRoot", str((project / output if not output.is_absolute() else output).resolve())))
    if getattr(args, "citation_mode", None):
        command_line.extend(("-CitationMode", args.citation_mode))
    if getattr(args, "word_refresh", None):
        command_line.extend(("-WordRefresh", args.word_refresh))
    return command_line, project


def run(args, operation: str) -> int:
    command_line, project = command_args(args, operation)
    return int(subprocess.run(command_line, cwd=project, check=False).returncode)
