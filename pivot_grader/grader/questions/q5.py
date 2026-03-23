from __future__ import annotations

from typing import Any

import pandas as pd

from grader.pivot_checker import compare_pivot_values, evaluate_highlight_formatting
from grader.utils.normalize import normalize_label

_ALL_MONTH_NAMES: frozenset[str] = frozenset({
	"jan", "feb", "mar", "apr", "may", "jun",
	"jul", "aug", "sep", "oct", "nov", "dec",
	"january", "february", "march", "april", "june", "july", "august",
	"september", "october", "november", "december",
})

_SUMMER_MONTHS: frozenset[str] = frozenset({
	"jun", "jul", "aug", "june", "july", "august",
})

_NON_SUMMER_MONTHS: frozenset[str] = _ALL_MONTH_NAMES - _SUMMER_MONTHS


def _sheet_labels(df: pd.DataFrame) -> frozenset[str]:
	labels: set[str] = set()
	if not df.empty:
		for col in df.columns:
			for val in df[col].dropna().astype(str):
				labels.add(normalize_label(val.strip()))
	for col in df.columns:
		labels.add(normalize_label(str(col).strip()))
	return frozenset(labels)


def check_q5_filter(df: pd.DataFrame) -> dict[str, Any]:
	labels = _sheet_labels(df)
	non_summer_found = sorted(_NON_SUMMER_MONTHS & labels)
	summer_found = sorted(_SUMMER_MONTHS & labels)
	return {
		"filter_ok": len(non_summer_found) == 0 and len(summer_found) > 0,
		"non_summer_months_found": non_summer_found,
		"summer_months_found": summer_found,
	}


def _filter_q5_answer_rows(answer_df: pd.DataFrame) -> pd.DataFrame:
	"""Remove month sub-rows so value checks compare category totals only."""
	if answer_df.empty or len(answer_df.columns) < 1:
		return answer_df

	month_labels = _NON_SUMMER_MONTHS | _SUMMER_MONTHS
	first_col = answer_df.columns[0]
	labels = answer_df[first_col].apply(
		lambda val: "" if pd.isna(val) else normalize_label(str(val).strip())
	)
	filtered = answer_df[~labels.isin(month_labels)].copy()
	return filtered if not filtered.empty else answer_df


def grade_question(
	student_df: pd.DataFrame,
	answer_df: pd.DataFrame,
	question_cfg: dict[str, Any],
	workbook_path: Any = None,
	sheet_name: str | None = None,
) -> dict[str, Any]:
	structural_issues: list[str] = []
	value_issues: list[str] = []

	if student_df.empty:
		formatting_score, formatting_issues = evaluate_highlight_formatting(workbook_path, sheet_name)
		structural_issues.append("Missing pivot table")
		value_issues.append("Missing pivot table")
		return {
			"structural_score": 0.0,
			"value_score": 0.0,
			"formatting_score": formatting_score,
			"explanation_score": 1.0,
			"structural_issues": structural_issues,
			"value_issues": value_issues,
			"formatting_issues": formatting_issues,
			"explanation_issues": [],
		}

	structural_score = 1.0
	filter_result = check_q5_filter(student_df)
	if not filter_result["filter_ok"]:
		structural_score = 0.0
		if filter_result["non_summer_months_found"] and filter_result["summer_months_found"]:
			structural_issues.append(
				"Month filter incorrect - non-summer months found: "
				+ ", ".join(filter_result["non_summer_months_found"])
			)
		elif not filter_result["summer_months_found"]:
			structural_issues.append("Month filter not applied - no summer months found in pivot")

	filtered_answer_df = _filter_q5_answer_rows(answer_df)
	value_result = compare_pivot_values(student_df, filtered_answer_df)
	value_score = 1.0 if value_result["match"] else 0.0
	if not value_result["match"] and value_result["mismatches"]:
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
