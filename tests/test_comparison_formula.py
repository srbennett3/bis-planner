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
    _build_equip2_name_formula,
    _build_equip_name_formula,
    _build_stat_diff_helper_formula,
    _col_letter,
)


class TestCmpRawFormula(unittest.TestCase):
    def test_stat_columns_excludes_spell_dmg_heal(self):
        self.assertNotIn("Spell Dmg/Heal", STAT_COLUMNS)

    def test_cmp_raw_formula_structure(self):
        first_h = len(SHEET_FIELDS) + 1
        h1 = [_col_letter(first_h + i) for i in range(HELPER_DIFF_COUNT)]
        h2_start = first_h + HELPER_DIFF_COUNT
        h2 = [_col_letter(h2_start + i) for i in range(HELPER_DIFF_COUNT)]
        f = _build_cmp_raw_formula(10, "M", "N", h1, h2, "ItemDB!A$2:Z$500", 25)
        self.assertTrue(f.startswith("="))
        self.assertNotIn("LET(", f)
        self.assertIn("TEXTJOIN", f)
        self.assertIn(f"{h1[0]}10:{h1[-1]}10", f)
        self.assertIn(f"{h2[0]}10:{h2[-1]}10", f)
        self.assertNotIn("ROUND(ABS(", f)
        self.assertIn("IFERROR(VLOOKUP(M10", f)
        self.assertIn("IFERROR(VLOOKUP(N10", f)
        self.assertIn('OR(C10="Ring",C10="Trinket")', f)
        self.assertIn("Slot 1 Comparison:", f)
        self.assertIn("Slot 2 comparison", f)
        self.assertIn('TRIM(IFERROR(D10,""))=TRIM(IFERROR(M10,""))', f)
        self.assertIn('"Equipped slot 1:"&CHAR(10)&', f)
        self.assertIn('"Equipped slot 2"&CHAR(10)&', f)
        self.assertIn('"Currently Equipped:"&CHAR(10)&', f)
        self.assertIn('SUBSTITUTE(SUBSTITUTE(TRIM(IFERROR(VLOOKUP(M10', f)
        self.assertNotIn("Spell Dmg/Heal", f)
        self.assertLess(len(f), 16384)

    def test_stat_diff_helper_formula(self):
        h = _build_stat_diff_helper_formula(10, "K", "ItemDB!A$2:Z$500", 0)
        self.assertTrue(h.startswith("="))
        self.assertIn("IFERROR(VLOOKUP(D10", h)
        self.assertIn("IFERROR(VLOOKUP(K10", h)
        self.assertIn("ROUND(ABS(", h)
        self.assertIn("TRIM(IFERROR(K10", h)

    def test_cmp_raw_formula_each_stat_suffix_present(self):
        first_h = 15
        h1 = [_col_letter(first_h + i) for i in range(HELPER_DIFF_COUNT)]
        h2 = [_col_letter(first_h + HELPER_DIFF_COUNT + i) for i in range(HELPER_DIFF_COUNT)]
        f = _build_cmp_raw_formula(4, "M", "N", h1, h2, "ItemDB!A$2:Z$100", 20)
        self.assertIn("TEXTJOIN", f)
        self.assertIn(f"{h1[0]}4:{h1[-1]}4", f)
        for key in STAT_COLUMNS:
            with self.subTest(key=key):
                hh = _build_stat_diff_helper_formula(4, "K", "ItemDB!A$2:Z$100", STAT_COLUMNS.index(key))
                self.assertIn(f'&" {key}"', hh)


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
        self.assertIn("Ring 1", f)
        self.assertIn("Trinket 1", f)

    def test_equip2_formula_basic(self):
        spec_order = ["Holy", "Protection", "Retribution"]
        ce_spec_rows = {
            "Holy": (10, 25),
            "Protection": (27, 42),
            "Retribution": (44, 59),
        }
        f = _build_equip2_name_formula(12, ce_spec_rows, spec_order)
        self.assertTrue(f.startswith("="))
        self.assertIn("Ring 2", f)
        self.assertIn("Trinket 2", f)
        self.assertIn('OR(C12="Ring",C12="Trinket")', f)


if __name__ == "__main__":
    unittest.main()
