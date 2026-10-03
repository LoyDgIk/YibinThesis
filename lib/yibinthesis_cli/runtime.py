"""Shared runtime, path and UTF-8 helpers for the CLI."""

from __future__ import annotations

import os
from pathlib import Path
import sys


CLI_VERSION = "0.3.0"
PROJECT_CONFIG_NAME = "yibinthesis.project.json"
GENERATED_MARKER_NAME = ".yibinthesis-cli.json"


class CliError(RuntimeError):
    """An expected user-facing CLI error."""


def configure_stdio() -> None:
    """Keep user-facing CLI diagnostics UTF-8 on Windows console runners."""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8", errors="replace")


def resource_root() -> Path:
    extracted = getattr(sys, "_MEIPASS", None)
    if extracted:
        return Path(extracted).resolve()
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


def read_utf8(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except OSError as exc:
        raise CliError(f"无法读取文件：{path} ({exc})") from exc


def write_utf8(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("w", encoding="utf-8", newline="\n") as stream:
            stream.write(content)
    except OSError as exc:
        raise CliError(f"无法写入文件：{path} ({exc})") from exc


def is_inside(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
    except ValueError:
        return False
    return True


def schema_path(root: Path) -> str:
    schema = resource_root() / "yibinthesis.project.schema.json"
    try:
        return os.path.relpath(schema, root).replace(os.sep, "/")
    except ValueError:
        return schema.as_posix()
