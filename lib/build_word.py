"""Stable Word-builder entry point.

The implementation lives in :mod:`word_core`; this module keeps the existing
``python lib/build_word.py`` and ``import build_word`` contracts stable for
PowerShell, tests, and signature preparation.
"""

from __future__ import annotations

import importlib
from pathlib import Path
import sys

_LIB_ROOT = str(Path(__file__).resolve().parent)
if _LIB_ROOT not in sys.path:
    sys.path.insert(0, _LIB_ROOT)

_core = importlib.import_module("word_core")

# Preserve public names for callers that import the historical module.
for _name in getattr(_core, "__all__", ()):
    globals()[_name] = getattr(_core, _name)


def __getattr__(name: str):
    return getattr(_core, name)


def main() -> int:
    return _core.main()


if __name__ == "__main__":
    raise SystemExit(main())
