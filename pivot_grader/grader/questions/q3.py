from __future__ import annotations

from typing import Any

import pandas as pd

from grader.pivot_checker import compare_pivot_values_subset, evaluate_highlight_formatting, has_multiple_items_marker
from grader.utils.normalize import normalize_label

_KNOWN_VENDORS: frozenset[str] = frozenset({
	"sweetums industries",
	"jj's diner goods",
	"rent-a-swag inc.",
	"lil' sebastian co.",
})

_TOP3_VENDORS: frozenset[str] = frozenset({
	"sweetums industries",
	"jj's diner goods",
	"lil' sebastian co.",
})


def evaluate_q3_structure(df: pd.DataFrame) -> dict[str, Any]:
	if df.empty:
		return {
			"structure_ok": False,
			"style": "missing",
			"found_vendors_in_rows": [],
			"missing_vendors": sorted(_TOP3_VENDORS),
			"extra_vendors": [],
			"has_multiple_items_marker": False,
		}

	first_col_vals = df.iloc[:, 0].dropna().astype(str).map(normalize_label)
	known_normalized = {v: normalize_label(v) for v in _KNOWN_VENDORS}
	vendors_in_rows = {
		raw
		for raw, norm in known_normalized.items()
		if first_col_vals.str.contains(norm, regex=False).any()
	}

	present_top3 = vendors_in_rows & _TOP3_VENDORS
	extra_vendors = vendors_in_rows - _TOP3_VENDORS
	missing_vendors = _TOP3_VENDORS - present_top3

	has_marker = has_multiple_items_marker(df) or any(
		"multiple items" in str(col).strip().lower() for col in df.columns
	)

	no_vendor_rows = len(vendors_in_rows) == 0
	row_label_vendor_style = vendors_in_rows == _TOP3_VENDORS
	structure_ok = no_vendor_rows

	if no_vendor_rows:
		style = "column_or_report_filter"
	elif row_label_vendor_style:
		style = "row_label_vendor"
	else:
		style = "invalid_vendor_rows"

	return {
		"structure_ok": structure_ok,
		"style": style,
		"found_vendors_in_rows": sorted(present_top3 | extra_vendors),
		"missing_vendors": sorted(missing_vendors),
		"extra_vendors": sorted(extra_vendors),
		"has_multiple_items_marker": has_marker,
	}


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

	q3_structure = evaluate_q3_structure(student_df)
	structural_score = 1.0 if q3_structure["structure_ok"] else 0.0
	if not q3_structure["structure_ok"]:
		structural_issues.append("Incorrect filter")

	required_labels: list[str] = []
	if not answer_df.empty:
		for val in answer_df.iloc[:, 0].dropna().astype(str):
			label = val.strip()
			if label and "total" not in label.lower():
				required_labels.append(label)
			if len(required_labels) == 3:
				break

	value_result = compare_pivot_values_subset(
		student_df,
		answer_df,
		required_labels=required_labels,
		ignore_labels=q3_structure.get("found_vendors_in_rows", []),
	)
	value_score = 1.0 if value_result["match"] else 0.0

	if not value_result["match"] and value_result["mismatches"]:
		first = value_result["mismatches"][0]
		value_issues.append(
			f"Mismatch at {first['label']}: expected {first['expected']}, actual {first['actual']}."
		)
	if not value_result["match"] and value_result.get("missing_required"):
		value_issues.append(
			"Missing required top product(s): " + ", ".join(value_result["missing_required"]) + "."
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
