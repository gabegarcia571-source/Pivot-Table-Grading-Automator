from __future__ import annotations

from typing import Any

import pandas as pd

from grader.pivot_checker import compare_pivot_values
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
		for val in df.iloc[:, 0].dropna().astype(str):
			labels.add(normalize_label(val.strip()))
	for col in df.columns:
		labels.add(normalize_label(str(col).strip()))
	return frozenset(labels)


def check_q5_filter(df: pd.DataFrame) -> dict[str, Any]:
	labels = _sheet_labels(df)
	non_summer_found = sorted(_NON_SUMMER_MONTHS & labels)
	return {
		"filter_ok": len(non_summer_found) == 0,
		"non_summer_months_found": non_summer_found,
	}


def grade_question(
	student_df: pd.DataFrame,
	answer_df: pd.DataFrame,
	question_cfg: dict[str, Any],
) -> dict[str, Any]:
	structural_issues: list[str] = []
	value_issues: list[str] = []

	if student_df.empty:
		structural_issues.append("Missing pivot table")
		value_issues.append("Missing pivot table")
		return {
			"structural_score": 0.0,
			"value_score": 0.0,
			"formatting_score": 1.0,
			"explanation_score": 1.0,
			"structural_issues": structural_issues,
			"value_issues": value_issues,
			"formatting_issues": [],
			"explanation_issues": [],
		}

	structural_score = 1.0
	filter_result = check_q5_filter(student_df)
	if not filter_result["filter_ok"]:
		structural_score = 0.0
		structural_issues.append(
			"Month filter incorrect - non-summer months found: "
			+ ", ".join(filter_result["non_summer_months_found"])
			+ "."
		)

	value_result = compare_pivot_values(student_df, answer_df)
	value_score = 1.0 if value_result["match"] else 0.0
	if not value_result["match"] and value_result["mismatches"]:
		first = value_result["mismatches"][0]
		value_issues.append(
			f"Mismatch at {first['label']}: expected {first['expected']}, actual {first['actual']}."
		)

	return {
		"structural_score": structural_score,
		"value_score": value_score,
		"formatting_score": 1.0,
		"explanation_score": 1.0,
		"structural_issues": structural_issues,
		"value_issues": value_issues,
		"formatting_issues": [],
		"explanation_issues": [],
	}
