from __future__ import annotations

from typing import Any

import pandas as pd

from grader.pivot_checker import compare_pivot_values
from grader.qualitative_grader import grade_explanation


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
    qid: str = "Q6",
) -> dict[str, Any]:
    value_issues: list[str] = []
    explanation_issues: list[str] = []

    if student_df.empty:
        return {
            "structural_score": 1.0,
            "value_score": 0.0,
            "formatting_score": 1.0,
            "explanation_score": 1.0,
            "structural_issues": [],
            "value_issues": ["Missing pivot table"],
            "formatting_issues": [],
            "explanation_issues": [],
        }

    value_result = compare_pivot_values(student_df, answer_df)
    value_score = 1.0 if value_result["match"] else 0.0
    if not value_result["match"] and value_result["mismatches"]:
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
            explanation_issues.append(str(llm_result.get("brief_reason", "NEEDS_REVIEW")))
        elif llm_result.get("deduct_explanation", False):
            explanation_score = 0.0
            explanation_issues.append(str(llm_result.get("brief_reason", "Needs more detail")))

    return {
        "structural_score": 1.0,
        "value_score": value_score,
        "formatting_score": 1.0,
        "explanation_score": explanation_score,
        "structural_issues": [],
        "value_issues": value_issues,
        "formatting_issues": [],
        "explanation_issues": explanation_issues,
    }
