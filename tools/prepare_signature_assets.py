#!/usr/bin/env python3
"""Prepare PDF-side signature images from metadata without changing source assets."""

from __future__ import annotations

import argparse
from pathlib import Path

import build_word


OVERRIDE_NAME = "yibinthesis-build-overrides.tex"
OUTPUT_NAMES = {
    "author-signature": "yibinthesis-author-signature.png",
    "advisor-signature": "yibinthesis-advisor-signature.png",
}


def locate_metadata(main_path: Path, project_root: Path) -> Path:
    main_dir = main_path.parent
    allow_project_fallback = build_word.path_is_within(main_path, project_root)
    main_text = main_path.read_text(encoding="utf-8")
    for event in build_word.parse_main_events(main_text):
        if event.kind not in {"input", "include"} or event.phase != "pre" or not event.value:
            continue
        candidate = build_word.resolve_source(
            event.value,
            main_dir,
            project_root,
            allow_project_fallback=allow_project_fallback,
        )
        if "\\yibinsetup" in candidate.read_text(encoding="utf-8"):
            return candidate
    candidates = [main_dir / "metadata.tex"]
    if allow_project_fallback:
        candidates.append(project_root / "metadata.tex")
    metadata = next((path.resolve() for path in candidates if path.is_file()), None)
    if metadata is None:
        raise build_word.BuildError("未找到 metadata.tex，无法准备签名图片。")
    return metadata


def prepare(main_path: Path, project_root: Path, output_dir: Path) -> Path | None:
    metadata_path = locate_metadata(main_path, project_root)
    metadata = build_word.parse_metadata(metadata_path)
    mode = metadata.get("signature-background", "preserve").strip().casefold()
    if mode not in {"preserve", "whiten"}:
        raise build_word.BuildError(
            "元数据 signature-background 只能是 preserve 或 whiten：" + mode
        )

    output_dir.mkdir(parents=True, exist_ok=True)
    override = output_dir / OVERRIDE_NAME
    generated = [output_dir / name for name in OUTPUT_NAMES.values()]
    if mode == "preserve":
        override.unlink(missing_ok=True)
        for path in generated:
            path.unlink(missing_ok=True)
        print("Signature background: preserve (no derived PDF assets)")
        return None

    values: dict[str, str] = {}
    for key, output_name in OUTPUT_NAMES.items():
        source = build_word._resolve_signature_asset(
            metadata,
            key,
            main_path.parent,
            metadata_path.parent,
        )
        target = output_dir / output_name
        if source is None:
            target.unlink(missing_ok=True)
            values[key] = ""
            continue
        build_word.whiten_signature_image(source, target)
        # The PDF output directory is added to TeX's search path by build.ps1.
        # Keeping the override relative avoids absolute-path warnings and keeps
        # generated metadata portable within the build tree.
        values[key] = target.name

    override.write_text(
        "\\yibinsetup{\n"
        f"  author-signature = {{{values['author-signature']}}},\n"
        f"  advisor-signature = {{{values['advisor-signature']}}}\n"
        "}\n",
        encoding="utf-8",
    )
    print(f"Signature background: whiten ({override})")
    return override


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--main", required=True, type=Path)
    parser.add_argument("--project-root", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    args = parser.parse_args()
    prepare(
        args.main.expanduser().resolve(),
        args.project_root.expanduser().resolve(),
        args.output_dir.expanduser().resolve(),
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
