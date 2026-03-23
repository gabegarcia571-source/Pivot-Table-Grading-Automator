from __future__ import annotations

from typing import Any

import pandas as pd

from grader.pivot_checker import evaluate_highlight_formatting
from grader.utils.normalize import normalize_label


def _detect_measure(df: pd.DataFrame) -> str:
    tokens: list[str] = []
    tokens.extend(normalize_label(str(col)) for col in df.columns)
    # Some exports place measure labels in top data rows instead of headers.
    scan_rows = min(20, len(df))
    for row_idx in range(scan_rows):
        for value in df.iloc[row_idx].tolist():
            if value is None or pd.isna(value):
                continue
            tokens.append(normalize_label(str(value)))

    joined = " ".join(tokens)
    has_count = "count" in joined
    has_sum = "sum" in joined or "total product price" in joined
    if has_count and not has_sum:
        return "count_only"
    if has_sum:
        return "sum_present"
    return "unknown"

def grade_question(
    student_df: pd.DataFrame,
    answer_df: pd.DataFrame,
    question_cfg: dict[str, Any],
    workbook_path: Any = None,
    sheet_name: str | None = None,
    qid: str = "Q8",
) -> dict[str, Any]:
    structural_issues: list[str] = []
    value_issues: list[str] = []
    formatting_issues: list[str] = []

    if student_df.empty:
        return {
            "structural_score": 0.0,
            "value_score": 0.0,
            "formatting_score": 0.0,
            "explanation_score": 1.0,
            "structural_issues": ["Missing pivot table"],
            "value_issues": ["Missing pivot table"],
            "formatting_issues": ["Missing highlight"],
            "explanation_issues": [],
        }

    # Structural: pivot must have customer-level granularity
    # (more than 1000 rows) and between 3 and 14 columns
    structural_score = 1.0
    if len(student_df) < 1000 or not (3 <= len(student_df.columns) <= 14):
        structural_score = 0.0
        structural_issues.append("Incorrect filter")

    measure = _detect_measure(student_df)
    measure_ok = measure != "count_only"
    if not measure_ok:
        value_issues.append("Wrong measure: used Count instead of Sum of total_product_price")

    # Value: structure plus measure selection must match the target setup.
    value_score = 1.0 if structural_score == 1.0 and measure_ok else 0.0

    # Formatting: any highlight on the sheet = full credit
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
