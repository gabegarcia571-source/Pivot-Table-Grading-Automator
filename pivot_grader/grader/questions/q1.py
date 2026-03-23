from __future__ import annotations

from typing import Any

import pandas as pd

from grader.pivot_checker import (
	compare_pivot_values,
	evaluate_highlight_formatting,
	is_desc_sorted_within_groups,
	is_group_order_desc,
)


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
	if question_cfg.get("sort_required") and (
		not is_desc_sorted_within_groups(student_df) or not is_group_order_desc(student_df, answer_df)
	):
		structural_score = 0.0
		structural_issues.append("Incorrect sort")

	value_result = compare_pivot_values(student_df, answer_df)
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
