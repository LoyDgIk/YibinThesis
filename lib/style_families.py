"""Single source of truth for the per-document-type Word style families.

Every document type owns a style family: the review and the opening report no
longer borrow the thesis ladder.  Keeping the names in one table stops the same
mapping from being retyped in the Python builders, the PowerShell refresher and
the format audit, where a missed copy used to fail silently.

``install_style_constants`` publishes the historical ``STYLE_*`` names so the
existing call sites keep working, and ``python lib/style_families.py --json``
emits the table for ``tools/refresh_word.ps1``.
"""

from __future__ import annotations

import json
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Mapping


@dataclass(frozen=True)
class StyleFamily:
    """One document type's style names plus the roles it overrides.

    ``overrides`` maps a thesis role to the role that replaces it, so a family
    can fold two thesis styles into one (the review renders the first paragraph
    with its body style).
    """

    document_type: str
    prefix: str
    roles: Mapping[str, str]
    overrides: Mapping[str, str] = field(default_factory=dict)

    def name(self, role: str) -> str:
        try:
            return self.roles[role]
        except KeyError:
            raise KeyError(
                f"{self.document_type} has no '{role}' style; "
                f"known roles: {', '.join(sorted(self.roles))}"
            ) from None

    @property
    def headings(self) -> tuple[str, str, str, str]:
        return (
            self.name("heading_1"),
            self.name("heading_2"),
            self.name("heading_3"),
            self.name("heading_4"),
        )


THESIS = StyleFamily(
    "thesis",
    "宜宾论文",
    {
        "body": "宜宾论文-正文",
        "first_paragraph": "宜宾论文-首段",
        "list_body": "宜宾论文-列表正文",
        "heading_1": "宜宾论文-一级标题",
        "heading_2": "宜宾论文-二级标题",
        "heading_3": "宜宾论文-三级标题",
        "heading_4": "宜宾论文-四级标题",
        "unnumbered_heading": "宜宾论文-无编号标题",
        "front_title": "宜宾论文-中文页标题",
        "english_abstract_title": "宜宾论文-英文摘要标题",
        "toc_title": "宜宾论文-目录标题",
        "appendix_heading": "宜宾论文-附录标题",
        "appendix_section": "宜宾论文-附录二级标题",
        "chinese_abstract": "宜宾论文-中文摘要正文",
        "english_abstract": "宜宾论文-英文摘要正文",
        "keywords": "宜宾论文-关键词",
        "figure": "宜宾论文-插图",
        "equation": "宜宾论文-公式",
        "figure_caption": "宜宾论文-图题",
        "table_caption": "宜宾论文-表题",
        "table_continuation": "宜宾论文-续表题",
        "table_text": "宜宾论文-表格正文",
        "table_center": "宜宾论文-表格居中",
        "table_right": "宜宾论文-表格右对齐",
        "table_header": "宜宾论文-表头",
        "table_header_left": "宜宾论文-表头左对齐",
        "table_header_right": "宜宾论文-表头右对齐",
        "bibliography": "宜宾论文-参考文献",
        "notes": "宜宾论文-注释",
        "citation": "宜宾论文-文献上标",
        "three_line_table": "宜宾论文-三线表",
        "table_of_figures": "Table of Figures",
    },
)

PROPOSAL = StyleFamily(
    "proposal",
    "宜宾开题",
    {
        "title": "宜宾开题-标题",
        "subtitle": "宜宾开题-副标题",
        "label": "宜宾开题-栏目",
        "body": "宜宾开题-正文",
        "heading_1": "宜宾开题-一级标题",
        "heading_2": "宜宾开题-二级标题",
        "heading_3": "宜宾开题-三级标题",
        "heading_4": "宜宾开题-四级标题",
        "figure": "宜宾开题-插图",
        "figure_caption": "宜宾开题-图题",
        "table_caption": "宜宾开题-表题",
        "bibliography": "宜宾开题-参考文献",
        "citation": "宜宾开题-文献上标",
        "prompt": "宜宾开题-提示",
        "unnumbered_heading": "宜宾开题-无编号标题",
        "signature": "宜宾开题-签名",
        "form_table": "YibinProposalForm",
    },
    overrides={
        role: role
        for role in (
            "body",
            "heading_1",
            "heading_2",
            "heading_3",
            "heading_4",
            "unnumbered_heading",
            "figure",
            "figure_caption",
            "table_caption",
            "bibliography",
            "citation",
        )
    },
)

LITERATURE_REVIEW = StyleFamily(
    "literature-review",
    "宜宾综述",
    {
        "document_title": "宜宾综述-文档标题",
        "thesis_title": "宜宾综述-论文题目",
        "info_label": "宜宾综述-信息标签",
        "info_value": "宜宾综述-信息值",
        "date": "宜宾综述-日期",
        "cover_field": "宜宾综述-封面填写行",
        "cover_label": "宜宾综述-封面标签",
        "cover_value": "宜宾综述-封面填值",
        "author": "宜宾综述-作者",
        "affiliation": "宜宾综述-作者单位",
        "abstract_title": "宜宾综述-中文摘要标题",
        "english_abstract_title": "宜宾综述-英文摘要标题",
        "chinese_abstract": "宜宾综述-中文摘要正文",
        "english_abstract": "宜宾综述-英文摘要正文",
        "keywords": "宜宾综述-中文关键词",
        "english_keywords": "宜宾综述-英文关键词",
        "body": "宜宾综述-正文",
        "heading_1": "宜宾综述-一级标题",
        "heading_2": "宜宾综述-二级标题",
        "heading_3": "宜宾综述-三级标题",
        "heading_4": "宜宾综述-四级标题",
        "unnumbered_heading": "宜宾综述-无编号标题",
        "bibliography": "宜宾综述-参考文献",
        "info_table": "YibinReviewInfoLayout",
    },
    overrides={
        "body": "body",
        "first_paragraph": "body",
        "heading_1": "heading_1",
        "heading_2": "heading_2",
        "heading_3": "heading_3",
        "heading_4": "heading_4",
        "unnumbered_heading": "unnumbered_heading",
        "front_title": "abstract_title",
        "english_abstract_title": "english_abstract_title",
        "chinese_abstract": "chinese_abstract",
        "english_abstract": "english_abstract",
        "keywords": "keywords",
        "bibliography": "bibliography",
    },
)

# Layout tables are shared by several document types, so they are not part of a
# single family.
SHARED_STYLES = {
    "cover_layout_table": "YibinCoverLayout",
    "front_layout_table": "YibinFrontLayout",
}

FAMILIES: dict[str, StyleFamily] = {
    THESIS.document_type: THESIS,
    PROPOSAL.document_type: PROPOSAL,
    LITERATURE_REVIEW.document_type: LITERATURE_REVIEW,
}


def family_for(document_type: str) -> StyleFamily:
    try:
        return FAMILIES[document_type]
    except KeyError:
        raise KeyError(
            f"unknown document type {document_type!r}; "
            f"known types: {', '.join(sorted(FAMILIES))}"
        ) from None


def remap_from_thesis(family: StyleFamily) -> dict[str, str]:
    """Return the {thesis style: family style} table for one document type."""

    remap: dict[str, str] = {}
    for thesis_role, family_role in family.overrides.items():
        thesis_name = THESIS.name(thesis_role)
        family_name = family.name(family_role)
        if thesis_name != family_name:
            remap[thesis_name] = family_name
    return remap


def install_style_constants(namespace: dict) -> None:
    """Publish the historical ``STYLE_*`` names into a module namespace."""

    for role, name in THESIS.roles.items():
        namespace[f"STYLE_{role.upper()}"] = name
    for role, name in PROPOSAL.roles.items():
        namespace[f"STYLE_PROPOSAL_{role.upper()}"] = name
    for role, name in LITERATURE_REVIEW.roles.items():
        namespace[f"STYLE_REVIEW_{role.upper()}"] = name
    for role, name in SHARED_STYLES.items():
        namespace[f"STYLE_{role.upper()}"] = name


def json_payload() -> dict[str, object]:
    """Serialise the table for the PowerShell refresher."""

    return {
        "shared": dict(SHARED_STYLES),
        "families": {
            document_type: {
                "prefix": family.prefix,
                "styles": dict(sorted(family.roles.items())),
                "headings": list(family.headings),
                "citation": family.roles.get("citation"),
                "overrides": dict(sorted(family.overrides.items())),
            }
            for document_type, family in sorted(FAMILIES.items())
        },
    }


def main(argv: list[str] | None = None) -> int:
    arguments = sys.argv[1:] if argv is None else argv
    if not arguments or arguments[0] != "--json":
        print("usage: python lib/style_families.py --json [path]", file=sys.stderr)
        return 2
    payload = json.dumps(json_payload(), ensure_ascii=False, indent=2) + "\n"
    if len(arguments) > 1:
        # Writing here keeps the encoding correct; piping through a shell on
        # Windows would re-encode the CJK style names.
        Path(arguments[1]).write_text(payload, encoding="utf-8", newline="\n")
    else:
        sys.stdout.write(payload)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
