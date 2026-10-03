"""Modular command-line implementation for YibinThesis."""

from .app import main, make_parser
from .runtime import CLI_VERSION, CliError, resource_root

__all__ = ["CLI_VERSION", "CliError", "main", "make_parser", "resource_root"]
