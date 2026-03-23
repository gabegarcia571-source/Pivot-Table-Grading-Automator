from __future__ import annotations

from typing import Any

import pandas as pd

from grader.pivot_checker import evaluate_highlight_formatting
from grader.qualitative_grader import grade_explanation
from grader.utils.normalize import normalize_label

_REQUIRED_BOTH_PROMO_IDS: frozenset[int] = frozenset({4, 12, 27})


def _extract_explanation_text(df: pd.DataFrame) -> str:
    text_values: list[str] = []
    for col in df.columns:
        series = df[col].dropna().astype(str)
        for value in series:
            value = value.strip()
            if len(value.split()) >= 6:
                text_values.append(value)
    return "\n".join(text_values)


def _safe_int(value: Any) -> int | None:
    try:
        num = int(float(value))
    except (TypeError, ValueError):
        return None
    return num if num > 0 else None


def _detect_crosstab_columns(df: pd.DataFrame) -> tuple[int, int] | None:
    # Header row often contains "Row Labels", "Honey", "No Promo Code" values.
    for _, row in df.iterrows():
        honey_idx = -1
        no_promo_idx = -1
        for idx, value in enumerate(row):
            label = normalize_label("" if pd.isna(value) else str(value))
            if "honey" in label:
                honey_idx = idx
            if "no promo" in label:
                no_promo_idx = idx
        if honey_idx >= 0 and no_promo_idx >= 0:
            return honey_idx, no_promo_idx
    return None


def _extract_both_ids_crosstab(df: pd.DataFrame, honey_idx: int, no_promo_idx: int) -> set[int]:
    both_ids: set[int] = set()
    for _, row in df.iterrows():
        customer_id = _safe_int(row.iloc[0])
        if customer_id is None:
            continue

        honey = pd.to_numeric(row.iloc[honey_idx], errors="coerce")
        no_promo = pd.to_numeric(row.iloc[no_promo_idx], errors="coerce")
        if pd.isna(honey) or pd.isna(no_promo):
            continue
        if float(honey) > 0 and float(no_promo) > 0:
            both_ids.add(customer_id)
    return both_ids


def _extract_both_ids_nested(df: pd.DataFrame) -> set[int]:
    both_flags: dict[int, dict[str, bool]] = {}
    current_customer: int | None = None

    for _, row in df.iterrows():
        label = "" if pd.isna(row.iloc[0]) else str(row.iloc[0]).strip()
        norm = normalize_label(label)
        if not norm:
            continue

        maybe_customer = _safe_int(label)
        if maybe_customer is not None:
            current_customer = maybe_customer
            both_flags.setdefault(current_customer, {"honey": False, "no_promo": False})
            continue

        if current_customer is None:
            continue

        if "honey" in norm:
            both_flags[current_customer]["honey"] = True
        elif "no promo" in norm:
            both_flags[current_customer]["no_promo"] = True

    return {
        cid
        for cid, flags in both_flags.items()
        if flags.get("honey") and flags.get("no_promo")
    }


def _extract_both_ids(df: pd.DataFrame) -> set[int]:
    crosstab_cols = _detect_crosstab_columns(df)
    if crosstab_cols is not None:
        return _extract_both_ids_crosstab(df, crosstab_cols[0], crosstab_cols[1])
    return _extract_both_ids_nested(df)


def _layout_style(df: pd.DataFrame) -> str:
    if _detect_crosstab_columns(df) is not None:
        return "matrix"
    return "flat_or_nested"


def _detect_measure(df: pd.DataFrame) -> str:
    tokens: list[str] = []
    tokens.extend(normalize_label(str(col)) for col in df.columns)
    scan_rows = min(20, len(df))
    for row_idx in range(scan_rows):
        for value in df.iloc[row_idx].tolist():
            if value is None or pd.isna(value):
                continue
            tokens.append(normalize_label(str(value)))

    joined = " ".join(tokens)
    if "count of promo code" in joined or "count of promo_code" in joined:
        return "count_promo_code"
    if "count of order id" in joined or "count of order_id" in joined:
        return "count_order_id"
    return "unknown"


def grade_question(
    student_df: pd.DataFrame,
    answer_df: pd.DataFrame,
    question_cfg: dict[str, Any],
    workbook_path: Any = None,
    sheet_name: str | None = None,
    qid: str = "Q9",
) -> dict[str, Any]:
    structural_issues: list[str] = []
    value_issues: list[str] = []
    explanation_issues: list[str] = []
    needs_review = False

    if student_df.empty:
        formatting_score, formatting_issues = evaluate_highlight_formatting(workbook_path, sheet_name)
        return {
            "structural_score": 1.0,
            "value_score": 0.0,
            "formatting_score": formatting_score,
            "explanation_score": 1.0,
            "structural_issues": structural_issues,
            "value_issues": ["Missing pivot table"],
            "formatting_issues": formatting_issues,
            "explanation_issues": [],
            "needs_review": needs_review,
        }

    student_layout = _layout_style(student_df)
    answer_layout = _layout_style(answer_df)

    if student_layout == "matrix":
        needs_review = True
        structural_issues.append(
            "NEEDS_REVIEW: alternate layout detected (customer rows with promo-code columns)"
        )

    student_measure = _detect_measure(student_df)
    answer_measure = _detect_measure(answer_df)
    measure_mismatch = (
        student_measure != "unknown"
        and answer_measure != "unknown"
        and student_measure != answer_measure
    )

    student_both_ids = _extract_both_ids(student_df)
    answer_both_ids = _extract_both_ids(answer_df)
    required_ids = _REQUIRED_BOTH_PROMO_IDS | answer_both_ids.intersection(_REQUIRED_BOTH_PROMO_IDS)

    missing_required = sorted(required_ids - student_both_ids)
    value_score = 1.0 if (not missing_required and not measure_mismatch) else 0.0
    if needs_review:
        value_score = 0.0
    if missing_required:
        if student_layout != answer_layout:
            value_issues.append(
                f"Structure differs from answer key ({student_layout} vs {answer_layout}); unable to align all required IDs."
            )
        value_issues.append(
            "Missing key customers (4, 12, 27)"
            + (f": {', '.join(str(v) for v in missing_required)}." if missing_required else ".")
        )
    if measure_mismatch:
        value_issues.append("Wrong measure: used Count of order_id instead of Count of promo_code")

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
        "structural_issues": structural_issues,
        "value_issues": value_issues,
        "formatting_issues": formatting_issues,
        "explanation_issues": explanation_issues,
        "needs_review": needs_review,
    }
