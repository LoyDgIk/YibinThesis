"""Guard the single style-family table shared by Python and PowerShell."""

from __future__ import annotations

import json
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "lib"))

import style_families as STYLES  # noqa: E402


class StyleFamilyTableTests(unittest.TestCase):
    def test_committed_json_matches_the_module(self) -> None:
        table = ROOT / "lib" / "style-families.json"
        self.assertTrue(table.is_file(), "style-families.json must be committed")
        committed = json.loads(table.read_text(encoding="utf-8"))
        self.assertEqual(
            committed,
            json.loads(json.dumps(STYLES.json_payload(), ensure_ascii=False)),
            "lib/style-families.json is stale; regenerate it with "
            "python lib/style_families.py --json lib/style-families.json",
        )

    def test_every_family_declares_the_heading_ladder(self) -> None:
        for document_type, family in STYLES.FAMILIES.items():
            with self.subTest(document_type=document_type):
                self.assertEqual(len(family.headings), 4)
                self.assertEqual(len(set(family.headings)), 4)
                for role in ("body", "bibliography"):
                    self.assertTrue(family.name(role))

    def test_style_names_are_unique_across_families(self) -> None:
        seen: dict[str, str] = {}
        for document_type, family in STYLES.FAMILIES.items():
            for role, name in family.roles.items():
                owner = seen.setdefault(name, document_type)
                self.assertEqual(
                    owner,
                    document_type,
                    f"{name!r} is claimed by {owner} and {document_type}",
                )

    def test_overrides_only_reference_shared_roles(self) -> None:
        for document_type, family in STYLES.FAMILIES.items():
            with self.subTest(document_type=document_type):
                for thesis_role, family_role in family.overrides.items():
                    self.assertIn(thesis_role, STYLES.THESIS.roles)
                    self.assertIn(family_role, family.roles)

    def test_remap_skips_roles_the_family_does_not_rename(self) -> None:
        remap = STYLES.remap_from_thesis(STYLES.PROPOSAL)
        self.assertEqual(
            remap[STYLES.THESIS.name("body")], STYLES.PROPOSAL.name("body")
        )
        self.assertNotIn(STYLES.THESIS.name("keywords"), remap)
        self.assertEqual(
            remap[STYLES.THESIS.name("citation")], STYLES.PROPOSAL.name("citation")
        )

    def test_review_folds_the_first_paragraph_into_its_body(self) -> None:
        remap = STYLES.remap_from_thesis(STYLES.LITERATURE_REVIEW)
        self.assertEqual(
            remap[STYLES.THESIS.name("first_paragraph")],
            STYLES.LITERATURE_REVIEW.name("body"),
        )
        self.assertEqual(
            remap[STYLES.THESIS.name("front_title")],
            STYLES.LITERATURE_REVIEW.name("abstract_title"),
        )

    def test_install_publishes_the_historical_constant_names(self) -> None:
        namespace: dict[str, object] = {}
        STYLES.install_style_constants(namespace)
        for expected in (
            "STYLE_BODY",
            "STYLE_HEADING_1",
            "STYLE_PROPOSAL_CITATION",
            "STYLE_PROPOSAL_FORM_TABLE",
            "STYLE_REVIEW_BIBLIOGRAPHY",
            "STYLE_COVER_LAYOUT_TABLE",
            "STYLE_THREE_LINE_TABLE",
        ):
            self.assertIn(expected, namespace)
        self.assertEqual(namespace["STYLE_BODY"], "宜宾论文-正文")
        self.assertEqual(namespace["STYLE_PROPOSAL_CITATION"], "宜宾开题-文献上标")

    def test_unknown_document_type_is_rejected(self) -> None:
        with self.assertRaises(KeyError):
            STYLES.family_for("dissertation")

    def test_powershell_refresher_reads_the_table(self) -> None:
        script = (ROOT / "tools" / "refresh_word.ps1").read_text(encoding="utf-8")
        self.assertIn("style-families.json", script)
        # The citation style must come from the table, not a second literal.
        self.assertNotIn("'宜宾开题-文献上标'", script)


if __name__ == "__main__":
    unittest.main()
