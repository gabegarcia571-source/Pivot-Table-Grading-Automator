from __future__ import annotations

from typing import Any

import pandas as pd

from grader.pivot_checker import evaluate_highlight_formatting
from grader.qualitative_grader import grade_explanation
from grader.utils.normalize import normalize_label

_AUDIENCES: tuple[str, ...] = ("adults", "families community", "kids teens")
_CATEGORIES: tuple[str, ...] = (
    "clothing accessories",
    "entertainment media",
    "food drink",
    "home lifestyle",
    "outdoor recreation",
)


def _extract_answer_pct_map(df: pd.DataFrame) -> dict[str, float]:
    if df.empty:
        return {}

    # Most answer sheets encode audience labels in the first row, not headers.
    if len(df) < 2:
        return {}

    header_row = df.iloc[0]
    audience_col_idx: dict[str, int] = {}
    for idx, value in enumerate(header_row):
        label = normalize_label("" if pd.isna(value) else str(value))
        if "families" in label:
            audience_col_idx["families community"] = idx
        elif "kids" in label:
            audience_col_idx["kids teens"] = idx
        elif "adults" in label:
            audience_col_idx["adults"] = idx

    if len(audience_col_idx) < 2:
        return {}

    out: dict[str, float] = {}
    for row_idx in range(1, len(df)):
        row = df.iloc[row_idx]
        category = normalize_label("" if pd.isna(row.iloc[0]) else str(row.iloc[0]))
        if category not in _CATEGORIES:
            continue
        for audience, col_idx in audience_col_idx.items():
            value = pd.to_numeric(row.iloc[col_idx], errors="coerce")
            if pd.isna(value):
                continue
            out[f"{category}::{audience}"] = float(value)

    # Ensure explicit zeros remain represented for known category/audience cells.
    for category in _CATEGORIES:
        for audience in _AUDIENCES:
            out.setdefault(f"{category}::{audience}", 0.0)
    return out


def _extract_student_pct_map(df: pd.DataFrame) -> tuple[dict[str, float], bool]:
    if df.empty or len(df.columns) < 2:
        return {}, False

    first_numeric_col: str | None = None
    for col in df.columns[1:]:
        if pd.to_numeric(df[col], errors="coerce").notna().any():
            first_numeric_col = str(col)
            break
    if first_numeric_col is None:
        return {}, False

    current_category: str | None = None
    category_totals: dict[str, float] = {}
    audience_values: dict[str, dict[str, float]] = {}

    for _, row in df.iterrows():
        label = normalize_label("" if pd.isna(row.iloc[0]) else str(row.iloc[0]))
        if not label or "total" in label or label == "row labels":
            continue

        value = pd.to_numeric(row[first_numeric_col], errors="coerce")
        if pd.isna(value):
            continue

        if label in _CATEGORIES:
            current_category = label
            category_totals[label] = float(value)
            audience_values.setdefault(label, {})
            continue

        if current_category is None:
            continue

        if "families" in label:
            audience_values[current_category]["families community"] = float(value)
        elif "kids" in label:
            audience_values[current_category]["kids teens"] = float(value)
        elif "adults" in label:
            audience_values[current_category]["adults"] = float(value)

    if not audience_values:
        return {}, False

    raw_max = max((abs(v) for cats in audience_values.values() for v in cats.values()), default=0.0)
    looks_raw_dollars = raw_max > 2.0

    out: dict[str, float] = {}
    for category, aud_map in audience_values.items():
        if looks_raw_dollars:
            total = category_totals.get(category, 0.0)
            if total <= 0:
                continue
            for audience, raw_val in aud_map.items():
                out[f"{category}::{audience}"] = raw_val / total
        else:
            for audience, pct in aud_map.items():
                out[f"{category}::{audience}"] = pct

    # Preserve zero entries so they are graded explicitly, not skipped as missing.
    for category in _CATEGORIES:
        for audience in _AUDIENCES:
            out.setdefault(f"{category}::{audience}", 0.0)

    return out, looks_raw_dollars


def _compare_maps(student_map: dict[str, float], answer_map: dict[str, float], tol: float = 1e-2) -> dict[str, Any]:
    if not student_map or not answer_map:
        return {
            "match": False,
            "mismatches": [
                {
                    "label": "<table>",
                    "expected": "non-empty comparable pivot",
                    "actual": "missing or non-numeric values",
                }
            ],
        }

    mismatches: list[dict[str, Any]] = []
    for label, expected in answer_map.items():
        actual = student_map.get(label)
        if actual is None:
            mismatches.append({"label": label, "expected": expected, "actual": None})
            continue
        if abs(float(actual) - float(expected)) > tol:
            mismatches.append({"label": label, "expected": float(expected), "actual": float(actual)})

    return {"match": len(mismatches) == 0, "mismatches": mismatches}


def _extract_explanation_text(df: pd.DataFrame) -> str:
    text_values: list[str] = []
    for col in df.columns:
        series = df[col].dropna().astype(str)
        for value in series:
            value = value.strip()
            if len(value.split()) >= 6:
                text_values.append(value)
    return "\n".join(text_values)


def grade_question(
    student_df: pd.DataFrame,
    answer_df: pd.DataFrame,
    question_cfg: dict[str, Any],
    workbook_path: Any = None,
    sheet_name: str | None = None,
    qid: str = "Q7",
) -> dict[str, Any]:
    value_issues: list[str] = []
    explanation_issues: list[str] = []

    if student_df.empty:
        formatting_score, formatting_issues = evaluate_highlight_formatting(workbook_path, sheet_name)
        return {
            "structural_score": 1.0,
            "value_score": 0.0,
            "formatting_score": formatting_score,
            "explanation_score": 1.0,
            "structural_issues": [],
            "value_issues": ["Missing pivot table"],
            "formatting_issues": formatting_issues,
            "explanation_issues": [],
        }

    answer_map = _extract_answer_pct_map(answer_df)
    student_map, used_raw_dollar_normalization = _extract_student_pct_map(student_df)
    value_result = _compare_maps(student_map, answer_map)

    value_score = 1.0 if value_result["match"] else 0.0
    if value_result["match"] and used_raw_dollar_normalization:
        value_issues.append("Values shown as raw dollars; correct % relationship verified.")
    elif not value_result["match"] and value_result["mismatches"]:
        first = value_result["mismatches"][0]
        value_issues.append(
            f"Mismatch at {first['label']}: expected {first['expected']}, actual {first['actual']}."
        )

    explanation_score = 1.0
    if question_cfg.get("explanation_required"):
        rubric_text = question_cfg.get("explanation_rubric", "")
        student_text = _extract_explanation_text(student_df)
        llm_result = grade_explanation(qid, student_text, rubric_text)
        if llm_result.get("needs_review", False):
            explanation_score = 0.0
            reason = str(llm_result.get("brief_reason", "Manual review needed"))
            explanation_issues.append(f"NEEDS_REVIEW: {reason}")
        elif llm_result.get("deduct_explanation", False):
            explanation_score = 0.0
            explanation_issues.append(str(llm_result.get("brief_reason", "Needs more detail")))

    formatting_score, formatting_issues = evaluate_highlight_formatting(workbook_path, sheet_name)

    return {
        "structural_score": 1.0,
        "value_score": value_score,
        "formatting_score": formatting_score,
        "explanation_score": explanation_score,
        "structural_issues": [],
        "value_issues": value_issues,
        "formatting_issues": formatting_issues,
        "explanation_issues": explanation_issues,
    }
