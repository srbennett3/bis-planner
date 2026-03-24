"""
Regression: BIS Planner export no longer emits per-row Δ helper columns or CmpRaw formulas
(Apps Script builds CmpRaw for Google Sheets).
"""

from pathlib import Path


def test_generate_bis_planner_source_has_no_delta_helpers():
    root = Path(__file__).resolve().parents[1]
    text = (root / "generate_bis_planner.py").read_text(encoding="utf-8")
    assert "HELPER_DIFF_COUNT" not in text
    assert "_build_stat_diff_helper_formula" not in text
    assert "_build_cmp_raw_formula" not in text
    assert "Δ1 " not in text
    assert "bp_total_cols = len(SHEET_FIELDS) + 2" not in text
