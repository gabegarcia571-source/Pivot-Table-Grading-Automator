from __future__ import annotations

from typing import Any

import pandas as pd

_Q10_EXPECTED_VENDOR_COUNT = 7
_Q10_VALUE_COLS = frozenset({"honey", "no promo code"})


def check_q10_filter(df: pd.DataFrame) -> dict[str, Any]:
    if df.empty:
        return {
            "filter_ok": False,
            "vendor_count": 0,
            "vendor_count_ok": False,
            "has_honey_col": False,
            "has_no_promo_col": False,
            "values_in_range": False,
        }

    first_col_vals = df.iloc[:, 0].dropna().astype(str).str.strip().str.lower()
    data_labels = [v for v in first_col_vals if "total" not in v and v != ""]
    vendor_count = len(set(data_labels))
    vendor_count_ok = vendor_count == _Q10_EXPECTED_VENDOR_COUNT

    col_headers_lower = {str(c).strip().lower() for c in df.columns}
    has_honey = any("honey" in h for h in col_headers_lower)
    has_no_promo = any("no promo" in h for h in col_headers_lower)

    numerics: list[float] = []
    for col in df.columns[1:]:
        converted = pd.to_numeric(df[col], errors="coerce").dropna()
        numerics.extend(float(v) for v in converted)
    values_in_range = bool(numerics) and all(0.0 <= v <= 1.0 for v in numerics)

    column_ok = has_honey and has_no_promo and values_in_range

    return {
        "filter_ok": column_ok and vendor_count_ok,
        "vendor_count": vendor_count,
        "vendor_count_ok": vendor_count_ok,
        "has_honey_col": has_honey,
        "has_no_promo_col": has_no_promo,
        "values_in_range": values_in_range,
    }


def grade_question(
    student_df: pd.DataFrame,
    answer_df: pd.DataFrame,
    question_cfg: dict[str, Any],
) -> dict[str, Any]:
    # TODO: Preserve current behavior until month-filter validation is implemented.
    # Q10 is intentionally forced to NEEDS_REVIEW.
    return {
        "structural_score": 1.0,
        "value_score": 0.0,
        "formatting_score": 1.0,
        "explanation_score": 1.0,
        "structural_issues": [],
        "value_issues": ["NEEDS_REVIEW: Q10 month-filter validation is incomplete"],
        "formatting_issues": [],
        "explanation_issues": [],
    }
