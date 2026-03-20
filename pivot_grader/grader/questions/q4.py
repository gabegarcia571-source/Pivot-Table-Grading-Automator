from __future__ import annotations

from typing import Any

import pandas as pd

from grader.qualitative_grader import grade_explanation

Q4_TARGET = 46.48
Q4_TOLERANCE = 2.0


def check_q4_average(
	df: pd.DataFrame,
	target: float = Q4_TARGET,
	tolerance: float = Q4_TOLERANCE,
) -> dict[str, Any]:
	numerics: list[float] = []

	for col in df.columns:
		converted = pd.to_numeric(df[col], errors="coerce").dropna()
		numerics.extend(float(v) for v in converted)

	if not numerics:
		return {"has_numeric": False, "match": False, "closest": None, "delta": None}

	closest = min(numerics, key=lambda v: abs(v - target))
	delta = abs(closest - target)
	return {
		"has_numeric": True,
		"match": delta <= tolerance,
		"closest": closest,
		"delta": delta,
	}


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
	qid: str = "Q4",
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

	q4 = check_q4_average(student_df)
	if not q4["has_numeric"]:
		value_score = 0.0
		value_issues.append("Missing pivot table")
	elif not q4["match"]:
		value_score = 0.0
		value_issues.append("Incorrect value")
	else:
		value_score = 1.0

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
