from __future__ import annotations

from typing import Any

import pandas as pd

from grader.pivot_checker import compare_pivot_values, evaluate_highlight_formatting
from grader.utils.normalize import normalize_label

_Q10_EXPECTED_VENDOR_COUNT = 14


def check_q10_filter(df: pd.DataFrame) -> dict[str, Any]:
    if df.empty:
        return {
            "filter_ok": False,
            "vendor_count": 0,
            "vendor_count_ok": False,
            "has_honey_col": False,
            "has_no_promo_col": False,
            "numeric_present": False,
            "values_in_range": False,
        }

    first_col_norm = [
        normalize_label(str(v).strip())
        for v in df.iloc[:, 0].dropna().astype(str)
    ]
    data_labels = [
        v for v in first_col_norm if v and "total" not in v and v not in {"honey", "no promo code"}
    ]
    vendor_count = len(set(data_labels))
    vendor_count_ok = vendor_count == _Q10_EXPECTED_VENDOR_COUNT

    col_headers_lower = {str(c).strip().lower() for c in df.columns}
    has_honey = any("honey" in h for h in col_headers_lower)
    has_no_promo = any("no promo" in h for h in col_headers_lower)

    # Nested submissions often encode Honey/No Promo as row labels, not columns.
    row_has_honey = any(v == "honey" for v in first_col_norm)
    row_has_no_promo = any("no promo" in v for v in first_col_norm)
    has_honey = has_honey or row_has_honey
    has_no_promo = has_no_promo or row_has_no_promo

    numerics: list[float] = []
    for col in df.columns[1:]:
        converted = pd.to_numeric(df[col], errors="coerce").dropna()
        numerics.extend(float(v) for v in converted)
    numeric_present = bool(numerics)
    values_in_range = bool(numerics) and all(0.0 <= v <= 1.0 for v in numerics)

    column_ok = has_honey and has_no_promo and numeric_present and values_in_range

    return {
        "filter_ok": column_ok and vendor_count_ok,
        "vendor_count": vendor_count,
        "vendor_count_ok": vendor_count_ok,
        "has_honey_col": has_honey,
        "has_no_promo_col": has_no_promo,
        "numeric_present": numeric_present,
        "values_in_range": values_in_range,
    }


def _contains_multiple_items_marker(df: pd.DataFrame) -> bool:
    all_strings: list[str] = []
    all_strings.extend(str(c) for c in df.columns)
    for col in df.columns:
        all_strings.extend(df[col].dropna().astype(str).tolist())
    return any("multiple items" in normalize_label(v) for v in all_strings)


def _first_numeric_col(df: pd.DataFrame) -> str | None:
    for col in df.columns[1:]:
        if pd.to_numeric(df[col], errors="coerce").notna().any():
            return str(col)
    return None


def _normalize_q10_nested_raw(df: pd.DataFrame) -> pd.DataFrame | None:
    value_col = _first_numeric_col(df)
    if value_col is None:
        return None

    rows: list[dict[str, Any]] = []
    current_vendor: str | None = None
    current_honey: float | None = None
    current_no_promo: float | None = None

    def flush_vendor() -> None:
        nonlocal current_vendor, current_honey, current_no_promo
        if current_vendor is None or current_honey is None or current_no_promo is None:
            current_vendor = None
            current_honey = None
            current_no_promo = None
            return
        total = current_honey + current_no_promo
        if total > 0:
            rows.append(
                {
                    "Row Labels": current_vendor,
                    "Honey": current_honey / total,
                    "No Promo Code": current_no_promo / total,
                    "Grand Total": 1.0,
                }
            )
        current_vendor = None
        current_honey = None
        current_no_promo = None

    for _, row in df.iterrows():
        label = normalize_label("" if pd.isna(row.iloc[0]) else str(row.iloc[0]))
        if not label or label == "row labels" or "total" in label:
            continue

        value = pd.to_numeric(row[value_col], errors="coerce")
        if pd.isna(value):
            continue
        value_f = float(value)

        if label == "honey":
            current_honey = value_f
            continue
        if "no promo" in label:
            current_no_promo = value_f
            continue

        # New vendor block starts; flush previous vendor if complete.
        if current_vendor is not None:
            flush_vendor()
        current_vendor = label

    if current_vendor is not None:
        flush_vendor()

    if not rows:
        return None
    return pd.DataFrame(rows)


def grade_question(
    student_df: pd.DataFrame,
    answer_df: pd.DataFrame,
    question_cfg: dict[str, Any],
    workbook_path: Any = None,
    sheet_name: str | None = None,
    qid: str = "Q10",
) -> dict[str, Any]:
    structural_issues: list[str] = []
    value_issues: list[str] = []

    if student_df.empty:
        formatting_score, formatting_issues = evaluate_highlight_formatting(workbook_path, sheet_name)
        return {
            "structural_score": 0.0,
            "value_score": 0.0,
            "formatting_score": formatting_score,
            "explanation_score": 1.0,
            "structural_issues": ["Missing pivot table"],
            "value_issues": ["Missing pivot table"],
            "formatting_issues": formatting_issues,
            "explanation_issues": [],
        }

    filter_result = check_q10_filter(student_df)
    structural_score = 1.0 if filter_result["filter_ok"] else 0.0
    if not filter_result["filter_ok"]:
        structural_issues.append("Incorrect filter")
        if not filter_result.get("vendor_count_ok", False):
            structural_issues.append(
                f"Expected {_Q10_EXPECTED_VENDOR_COUNT} vendors after filter, found {filter_result.get('vendor_count', 0)}."
            )
        if not filter_result.get("has_honey_col", False) or not filter_result.get("has_no_promo_col", False):
            structural_issues.append("Missing Honey/No Promo Code split.")
        if not filter_result.get("values_in_range", False):
            structural_issues.append("Values out of range (expected proportions between 0 and 1).")

    if not _contains_multiple_items_marker(student_df):
        structural_score = 0.0
        structural_issues.append("Incorrect filter")

    value_result = compare_pivot_values(student_df, answer_df)
    used_raw_dollar_normalization = False
    if not value_result["match"]:
        normalized_df = _normalize_q10_nested_raw(student_df)
        if normalized_df is not None:
            value_result = compare_pivot_values(normalized_df, answer_df)
            used_raw_dollar_normalization = value_result["match"]

    value_score = 1.0 if value_result["match"] else 0.0
    if used_raw_dollar_normalization:
        value_issues.append("Values shown as raw dollars, correct % relationship verified.")
    elif not value_result["match"] and value_result["mismatches"]:
        first = value_result["mismatches"][0]
        value_issues.append(
            f"Mismatch at {first['label']}: expected {first['expected']}, actual {first['actual']}."
        )

    formatting_score, formatting_issues = evaluate_highlight_formatting(workbook_path, sheet_name)

    return {
        "structural_score": structural_score,
        "value_score": value_score,
        "formatting_score": formatting_score,
        "explanation_score": 1.0,
        "structural_issues": structural_issues,
        "value_issues": value_issues,
        "formatting_issues": formatting_issues,
        "explanation_issues": [],
    }
