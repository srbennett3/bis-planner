"""Tests for BIS CmpRaw / Equip formula builders (see generate_bis_planner.py)."""

import os
import sys
import unittest

SCRIPT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, SCRIPT_DIR)

from generate_bis_planner import (  # noqa: E402
    HELPER_DIFF_COUNT,
    SHEET_FIELDS,
    STAT_COLUMNS,
    _build_cmp_raw_formula,
    _build_equip_name_formula,
    _build_stat_diff_helper_formula,
    _col_letter,
)


class TestCmpRawFormula(unittest.TestCase):
    def test_stat_columns_excludes_spell_dmg_heal(self):
        self.assertNotIn("Spell Dmg/Heal", STAT_COLUMNS)

    def test_cmp_raw_formula_structure(self):
        first_h = len(SHEET_FIELDS) + 1
        hletters = [_col_letter(first_h + i) for i in range(HELPER_DIFF_COUNT)]
        f = _build_cmp_raw_formula(10, "K", hletters, "ItemDB!A$2:Z$500", 25)
        self.assertTrue(f.startswith("="))
        self.assertNotIn("LET(", f)
        self.assertIn(f"{hletters[0]}10", f)
        self.assertIn("TEXTJOIN", f)
        self.assertIn("ROUND(ABS(", f)
        self.assertIn("IFERROR(VLOOKUP(K10", f)
        # Equip #N/A must not break IF(K="",...)
        self.assertIn('IF(IFERROR(K10,"")="","Current Equipment Not Specified"', f)
        self.assertNotIn("Spell Dmg/Heal", f)
        self.assertLess(len(f), 8192)

    def test_stat_diff_helper_formula(self):
        h = _build_stat_diff_helper_formula(10, "K", "ItemDB!A$2:Z$500", 0)
        self.assertTrue(h.startswith("="))
        self.assertIn("IFERROR(VLOOKUP(D10", h)
        self.assertIn("IFERROR(VLOOKUP(K10", h)

    def test_cmp_raw_formula_each_stat_suffix_present(self):
        first_h = 15
        hletters = [_col_letter(first_h + i) for i in range(HELPER_DIFF_COUNT)]
        f = _build_cmp_raw_formula(4, "K", hletters, "ItemDB!A$2:Z$100", 20)
        for key in STAT_COLUMNS:
            with self.subTest(key=key):
                self.assertIn(f'&" {key}"', f)


class TestEquipFormula(unittest.TestCase):
    def test_equip_formula_basic(self):
        spec_order = ["Holy", "Protection", "Retribution"]
        ce_spec_rows = {
            "Holy": (10, 25),
            "Protection": (27, 42),
            "Retribution": (44, 59),
        }
        f = _build_equip_name_formula(12, ce_spec_rows, spec_order)
        self.assertTrue(f.startswith("="))
        self.assertIn("IFERROR(IFS(", f)
        self.assertIn('B12="Holy"', f)
        self.assertIn("Current Equipment", f)


if __name__ == "__main__":
    unittest.main()
