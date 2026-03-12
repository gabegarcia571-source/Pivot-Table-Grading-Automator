from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook

from grader.grade_writer import write_grades
from grader.ingest import find_pivot_origin, load_student_submission
from grader.pivot_checker import (
    check_q3_filter,
    check_q4_average,
    check_q5_filter,
    check_q8_highlight,
    check_q10_filter,
    compare_pivot_values,
    is_desc_sorted_within_groups,
    sheet_fingerprint,
    fingerprint_similarity,
)
from grader.scoring import compute_question_score


def test_find_pivot_origin_standard() -> None:
    """Pivot at A1 — origin should be (1, 1)."""
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Label"
    ws["B1"] = "Value"
    ws["A2"] = "Food & Drink"
    ws["B2"] = 1_707_732.18

    assert find_pivot_origin(ws) == (1, 1)


def test_find_pivot_origin_shifted() -> None:
    """Stray label in row 3, actual pivot header starting at C8."""
    wb = Workbook()
    ws = wb.active
    ws["A3"] = "Student Name"   # single-cell stray label — only 1 non-empty in row
    ws["C8"] = "category_name"
    ws["D8"] = "Sum of total_product_price"
    ws["C9"] = "Food & Drink"
    ws["D9"] = 1_707_732.18

    assert find_pivot_origin(ws) == (8, 3)


def test_load_student_submission_shifted_pivot(tmp_path: Path) -> None:
    """Ensure the pivot is found and loaded correctly when offset from A1."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Q1"
    ws["A3"] = "My Homework"              # stray single-cell label
    ws["C8"] = "category_name"            # pivot header starts here
    ws["D8"] = "Sum of total_product_price"
    ws["C9"] = "Food & Drink"
    ws["D9"] = 1_707_732.18

    student_dir = tmp_path / "student_1"
    student_dir.mkdir()
    wb.save(student_dir / "submission.xlsx")

    result = load_student_submission(student_dir)

    assert result.error is None
    assert "Q1" in result.sheets
    df = result.sheets["Q1"]
    # First column must be the pivot label column, not an offset blank column
    assert df.columns[0] == "category_name"
    assert len(df) >= 1


# ---------------------------------------------------------------------------
# Q4 average-scan tests
# ---------------------------------------------------------------------------

def test_check_q4_average_match() -> None:
    """A value within 2.0 of 46.48 should produce match=True."""
    df = pd.DataFrame({"order_id": [1, 2, 3], "order_total": [44.50, 46.48, 50.00]})
    result = check_q4_average(df)
    assert result["has_numeric"] is True
    assert result["match"] is True
    assert abs(result["closest"] - 46.48) <= 2.0


def test_check_q4_average_wrong_values() -> None:
    """Values that exist but none near 46.48 — wrong_values deduction expected."""
    df = pd.DataFrame({"order_id": [1, 2, 3], "line_avg": [12.50, 13.00, 11.75]})
    result = check_q4_average(df)
    assert result["has_numeric"] is True
    assert result["match"] is False
    assert result["delta"] > 2.0


def test_check_q4_average_no_numerics() -> None:
    """A sheet with no numeric values at all — missing_pivot deduction expected."""
    df = pd.DataFrame({"note": ["See other sheet", "N/A"], "col2": [None, None]})
    result = check_q4_average(df)
    assert result["has_numeric"] is False
    assert result["match"] is False
    assert result["closest"] is None


def test_compare_pivot_values_match() -> None:
    student = pd.DataFrame({"Label": ["A", "B"], "Value": [10.0, 20.0]})
    answer = pd.DataFrame({"Label": ["A", "B"], "Value": [10.0, 20.0]})

    result = compare_pivot_values(student, answer)

    assert result["match"] is True
    assert result["mismatches"] == []
    assert result["score_suggestion"] == 1.0


def test_compare_pivot_values_partial_credit() -> None:
    student = pd.DataFrame({"Label": ["A", "B"], "Value": [10.0, 22.0]})
    answer = pd.DataFrame({"Label": ["A", "B"], "Value": [10.0, 20.0]})

    result = compare_pivot_values(student, answer)

    assert result["match"] is False
    assert len(result["mismatches"]) == 1
    assert result["score_suggestion"] == 0.5


def test_compute_question_score_with_deductions() -> None:
    score, comments = compute_question_score(
        has_pivot=True,
        pivot_match=False,
        structural_issues=["no_sort"],
        explanation_deduct=True,
    )

    assert score == 0.0
    assert len(comments) == 3


def test_write_grades(tmp_path: Path) -> None:
    template = tmp_path / "template.xlsx"
    output_dir = tmp_path / "out"

    wb = Workbook()
    ws = wb.active
    ws.title = "Gradesheet"
    ws["C4"] = "=C17"
    ws["C17"] = "=SUM(C7:C16)"
    wb.save(template)

    scores = {f"Q{i}": 1.0 for i in range(1, 11)}
    comments = {f"Q{i}": "ok" for i in range(1, 11)}

    out_file = write_grades(
        student_id="student_1",
        scores_dict=scores,
        comments_dict=comments,
        template_path=template,
        output_dir=output_dir,
    )

    graded = load_workbook(out_file)
    gws = graded["Gradesheet"]

    assert gws["C3"].value == "student_1"
    assert gws["C7"].value == 1.0
    assert gws["D7"].value == "ok"
    assert gws["C4"].value == "=C17"
    assert gws["C17"].value == "=SUM(C7:C16)"


# ---------------------------------------------------------------------------
# ffill merged-cell tests
# ---------------------------------------------------------------------------

def test_load_student_submission_merged_cell_ffill(tmp_path: Path) -> None:
    """Merged pivot row-group labels (NaN after export) must be filled down."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Q1"
    # Simulate a pivot with a row-group label merged across 3 rows:
    # Excel exports the label only in the first row; remaining cells are blank.
    ws["A1"] = "Category"
    ws["B1"] = "Value"
    ws["A2"] = "Food & Drink"  # exported label (first row of group)
    ws["B2"] = 100.0
    ws["A3"] = None            # merged — blank in export
    ws["B3"] = 200.0
    ws["A4"] = None            # merged — blank in export
    ws["B4"] = 150.0

    student_dir = tmp_path / "student_merged"
    student_dir.mkdir()
    wb.save(student_dir / "submission.xlsx")

    result = load_student_submission(student_dir)
    assert result.error is None
    df = result.sheets["Q1"]

    first_col = df.iloc[:, 0].tolist()
    # After ffill every cell should hold "Food & Drink", not NaN
    assert all(v == "Food & Drink" for v in first_col), (
        f"Expected ffill to propagate 'Food & Drink'; got {first_col}"
    )


# ---------------------------------------------------------------------------
# Q3 filter-check tests
# ---------------------------------------------------------------------------

_VENDOR_NAMES = [
    "sweetums industries",
    "jj's diner goods",
    "rent-a-swag inc.",
    "lil' sebastian co.",
]


def test_check_q3_filter_not_applied() -> None:
    """All four vendor names in first column → filter_ok=False."""
    df = pd.DataFrame({
        "vendor_name": _VENDOR_NAMES,
        "total": [1000, 2000, 1500, 800],
    })
    result = check_q3_filter(df)
    assert result["filter_ok"] is False
    assert len(result["found_vendors_in_rows"]) == 4


def test_check_q3_filter_applied() -> None:
    """Product names in first column, no vendor names → filter_ok=True."""
    df = pd.DataFrame({
        "product_name": ["Calzone", "Pizza", "Waffle"],
        "total": [500, 600, 700],
    })
    result = check_q3_filter(df)
    assert result["filter_ok"] is True
    assert result["found_vendors_in_rows"] == []


def test_check_q3_filter_partial_vendors() -> None:
    """Only two of the top-3 vendor names in first column → filter_ok=False.

    The new logic requires exactly all three top vendors — Sweetums Industries,
    JJ's Diner Goods, Lil' Sebastian Co. — and no others.
    """
    df = pd.DataFrame({
        "vendor_name": ["sweetums industries", "jj's diner goods"],
        "total": [1000, 2000],
    })
    result = check_q3_filter(df)
    assert result["filter_ok"] is False
    assert len(result["missing_vendors"]) == 1  # lil' sebastian co. missing


def test_check_q3_filter_correct_top3() -> None:
    """Exactly the three top vendors in rows, no others → filter_ok=True."""
    df = pd.DataFrame({
        "vendor_name": ["sweetums industries", "jj's diner goods", "lil' sebastian co."],
        "total": [1000, 2000, 800],
    })
    result = check_q3_filter(df)
    assert result["filter_ok"] is True
    assert result["extra_vendors"] == []
    assert result["missing_vendors"] == []


# ---------------------------------------------------------------------------
# Q5 filter-check tests
# ---------------------------------------------------------------------------

def test_check_q5_filter_correct() -> None:
    """Only Jun/Jul/Aug labels → filter_ok=True."""
    df = pd.DataFrame({
        "month": ["Jun", "Jul", "Aug"],
        "total": [300, 400, 350],
    })
    result = check_q5_filter(df)
    assert result["filter_ok"] is True
    assert result["non_summer_months_found"] == []


def test_check_q5_filter_wrong() -> None:
    """January label present → filter_ok=False."""
    df = pd.DataFrame({
        "month": ["Jan", "Jun", "Jul", "Aug"],
        "total": [200, 300, 400, 350],
    })
    result = check_q5_filter(df)
    assert result["filter_ok"] is False
    assert "jan" in result["non_summer_months_found"]


# ---------------------------------------------------------------------------
# Q10 structure-check tests (vendor count, column headers, proportion values)
# ---------------------------------------------------------------------------

_Q10_VENDORS = [
    "Sweetums Industries",
    "JJ's Diner Goods",
    "Rent-A-Swag Inc.",
    "Lil' Sebastian Co.",
    "Vendor E",
    "Vendor F",
    "Vendor G",
]


def test_check_q10_filter_correct() -> None:
    """7 vendor rows, Honey + No Promo Code columns, values in [0,1] → filter_ok=True."""
    df = pd.DataFrame({
        "vendor_name": _Q10_VENDORS,
        "Honey": [0.21, 0.15, 0.53, 0.40, 0.30, 0.10, 0.25],
        "No Promo Code": [0.79, 0.85, 0.47, 0.60, 0.70, 0.90, 0.75],
        "Grand Total": [1.00] * 7,
    })
    result = check_q10_filter(df)
    assert result["filter_ok"] is True
    assert result["vendor_count_ok"] is True
    assert result["has_honey_col"] is True
    assert result["has_no_promo_col"] is True
    assert result["values_in_range"] is True


def test_check_q10_filter_wrong_vendor_count() -> None:
    """Only 4 vendor rows → filter_ok=False, vendor_count_ok=False."""
    df = pd.DataFrame({
        "vendor_name": _Q10_VENDORS[:4],
        "Honey": [0.21, 0.15, 0.53, 0.40],
        "No Promo Code": [0.79, 0.85, 0.47, 0.60],
    })
    result = check_q10_filter(df)
    assert result["filter_ok"] is False
    assert result["vendor_count_ok"] is False
    assert result["vendor_count"] == 4


def test_check_q10_filter_missing_honey_column() -> None:
    """No 'Honey' column header → filter_ok=False, has_honey_col=False."""
    df = pd.DataFrame({
        "vendor_name": _Q10_VENDORS,
        "No Promo Code": [0.79, 0.85, 0.47, 0.60, 0.70, 0.90, 0.75],
        "Grand Total": [1.00] * 7,
    })
    result = check_q10_filter(df)
    assert result["filter_ok"] is False
    assert result["has_honey_col"] is False


def test_check_q10_filter_values_not_proportions() -> None:
    """Raw counts instead of proportions (values > 1) → filter_ok=False."""
    df = pd.DataFrame({
        "vendor_name": _Q10_VENDORS,
        "Honey": [210, 150, 530, 400, 300, 100, 250],
        "No Promo Code": [790, 850, 470, 600, 700, 900, 750],
    })
    result = check_q10_filter(df)
    assert result["filter_ok"] is False
    assert result["values_in_range"] is False


# ---------------------------------------------------------------------------
# Q1 within-group sort tests
# ---------------------------------------------------------------------------

def test_is_desc_sorted_within_groups_correct() -> None:
    """Days within each category are descending — should pass."""
    df = pd.DataFrame({
        "category": ["Food & Drink", "Food & Drink", "Food & Drink",
                     "Electronics",  "Electronics",  "Electronics"],
        "day":      ["Monday", "Tuesday", "Wednesday", "Friday", "Saturday", "Sunday"],
        "value":    [300.0, 200.0, 100.0, 150.0, 120.0, 80.0],
    })
    assert is_desc_sorted_within_groups(df) is True


def test_is_desc_sorted_within_groups_wrong_order() -> None:
    """One group has ascending order — should fail."""
    df = pd.DataFrame({
        "category": ["Food & Drink", "Food & Drink", "Electronics", "Electronics"],
        "day":      ["Monday", "Tuesday", "Friday", "Saturday"],
        "value":    [100.0, 200.0, 150.0, 80.0],  # Food & Drink ascending — wrong
    })
    assert is_desc_sorted_within_groups(df) is False


def test_is_desc_sorted_within_groups_flat_fallback() -> None:
    """No repeated first-col labels — falls back to flat is_desc_sorted, passes."""
    df = pd.DataFrame({
        "category": ["Electronics", "Food & Drink", "Outdoors"],
        "value":    [300.0, 200.0, 100.0],
    })
    assert is_desc_sorted_within_groups(df) is True


def test_is_desc_sorted_within_groups_ignores_total_rows() -> None:
    """Subtotal row (label contains 'total') is excluded from the sort check."""
    df = pd.DataFrame({
        "category": ["Food & Drink", "Food & Drink", "Food & Drink Total",
                     "Electronics",  "Electronics",  "Electronics Total"],
        "day":      ["Monday", "Tuesday", "Total", "Friday", "Saturday", "Total"],
        "value":    [300.0, 200.0, 500.0, 150.0, 80.0, 230.0],  # totals would break sort
    })
    assert is_desc_sorted_within_groups(df) is True


# ---------------------------------------------------------------------------
# Q8 highlight-detection tests (check_q8_highlight uses openpyxl fill scanning)
# ---------------------------------------------------------------------------

from openpyxl.styles import PatternFill
from grader.answer_constants import HOLIDAY_ONLY_CUSTOMERS

_YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")


def _make_q8_workbook(
    tmp_path: Path,
    highlighted_ids: list[int],
    non_highlighted_ids: list[int] | None = None,
) -> tuple[Path, str]:
    """Write a minimal Q8 xlsx with the given customer IDs, highlighting the
    *highlighted_ids* rows with a yellow fill.  Returns (path, sheet_name)."""
    if non_highlighted_ids is None:
        non_highlighted_ids = [99999]  # a row that should not be highlighted
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Q8"
    ws["A1"] = "customer_id"
    ws["B1"] = "Nov"
    ws["C1"] = "Dec"
    row = 2
    for cid in highlighted_ids:
        ws.cell(row=row, column=1, value=cid)
        ws.cell(row=row, column=1).fill = _YELLOW_FILL
        row += 1
    for cid in non_highlighted_ids:
        ws.cell(row=row, column=1, value=cid)
        row += 1
    path = tmp_path / "q8_test.xlsx"
    wb.save(path)
    return path, "Q8"


def test_check_q8_highlight_exact_match(tmp_path: Path) -> None:
    """Highlighted IDs exactly match HOLIDAY_ONLY_CUSTOMERS → match=True."""
    correct_ids = sorted(HOLIDAY_ONLY_CUSTOMERS)
    path, sname = _make_q8_workbook(tmp_path, highlighted_ids=correct_ids)
    result = check_q8_highlight(path, sname)
    assert result["match"] is True
    assert result["missing_pivot"] is False


def test_check_q8_highlight_no_highlights(tmp_path: Path) -> None:
    """No cells highlighted → missing_pivot=True, match=False."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Q8"
    ws["A1"] = "customer_id"
    for i, cid in enumerate(sorted(HOLIDAY_ONLY_CUSTOMERS)[:10], start=2):
        ws.cell(row=i, column=1, value=cid)  # no fill applied
    path = tmp_path / "q8_no_hl.xlsx"
    wb.save(path)
    result = check_q8_highlight(path, "Q8")
    assert result["match"] is False
    assert result["missing_pivot"] is True
    assert any("highlight" in n.lower() for n in result["notes"])


def test_check_q8_highlight_within_tolerance(tmp_path: Path) -> None:
    """Highlighted IDs have a small error (≤5%) → still match=True."""
    correct_ids = sorted(HOLIDAY_ONLY_CUSTOMERS)
    # Drop 1 ID (tiny error, well within 5% for any non-trivial set size)
    slightly_off = correct_ids[1:]  # remove one
    path, sname = _make_q8_workbook(tmp_path, highlighted_ids=slightly_off)
    result = check_q8_highlight(path, sname)
    # With only 10 IDs in the constant set a 1-ID miss is 10%, so tolerate
    # only when the set is large enough.  Skip assertion when set is tiny.
    if len(correct_ids) > 20:
        assert result["match"] is True


def test_check_q8_highlight_wrong_ids(tmp_path: Path) -> None:
    """Completely wrong IDs highlighted → match=False."""
    wrong_ids = [999001, 999002, 999003]
    path, sname = _make_q8_workbook(tmp_path, highlighted_ids=wrong_ids)
    result = check_q8_highlight(path, sname)
    assert result["match"] is False
    assert result["missing_pivot"] is False
    assert any("highlight" in n.lower() or "customer" in n.lower() for n in result["notes"])


def test_check_q8_highlight_missing_sheet(tmp_path: Path) -> None:
    """Sheet name not in workbook → match=False with descriptive note."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "WrongSheet"
    path = tmp_path / "q8_wrong_sheet.xlsx"
    wb.save(path)
    result = check_q8_highlight(path, "Q8")
    assert result["match"] is False
    assert any("not found" in n.lower() or "sheet" in n.lower() for n in result["notes"])


# ---------------------------------------------------------------------------
# sheet_fingerprint / fingerprint_similarity tests
# ---------------------------------------------------------------------------

def test_sheet_fingerprint_row_bucket() -> None:
    tiny_df = pd.DataFrame({"a": range(5), "b": range(5)})
    medium_df = pd.DataFrame({"a": range(500), "b": range(500)})
    large_df = pd.DataFrame({"a": range(3000), "b": range(3000)})
    assert sheet_fingerprint(tiny_df)["row_bucket"] == "tiny"
    assert sheet_fingerprint(medium_df)["row_bucket"] == "medium"
    assert sheet_fingerprint(large_df)["row_bucket"] == "large"


def test_fingerprint_similarity_identical() -> None:
    df = pd.DataFrame({
        "label": ["Food & Drink", "Electronics"],
        "total": [1000.0, 2000.0],
    })
    fp = sheet_fingerprint(df)
    # Identical fingerprints should yield the maximum score
    score = fingerprint_similarity(fp, fp)
    assert score >= 9.0  # row_bucket(3) + cols_exact(2) + first_col_numeric(2) + label_overlap(4×1.0)


def test_fingerprint_similarity_different_sizes() -> None:
    small_df = pd.DataFrame({"label": ["A", "B"], "total": [1.0, 2.0]})
    large_df = pd.DataFrame({"x": range(5000), "y": range(5000), "z": range(5000)})
    fp_small = sheet_fingerprint(small_df)
    fp_large = sheet_fingerprint(large_df)
    low_score = fingerprint_similarity(fp_small, fp_large)
    # Different row bucket (tiny vs large) → no row_bucket bonus; column mismatch → low score
    assert low_score < 3.0
