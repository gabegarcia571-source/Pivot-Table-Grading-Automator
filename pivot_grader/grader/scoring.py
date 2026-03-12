from __future__ import annotations

DEDUCTIONS = {
    "missing_pivot": {"points": -1.0, "comment": "No pivot table submitted for this question."},
    "no_sort": {"points": -0.3, "comment": "Table not sorted as required by the question."},
    "extra_columns": {
        "points": -0.3,
        "comment": "Pivot table contains extra/unnecessary columns.",
    },
    "answer_not_highlighted": {
        "points": -0.5,
        "comment": "Correct answer is present but not highlighted.",
    },
    "wrong_values": {
        "points": -0.7,
        "comment": "Incorrect columns or data values; correct answer not present. Partial credit awarded.",
    },
    "bad_explanation": {
        "points": -0.3,
        "comment": "Explanation is missing, incorrect, or inconsistent with the analysis.",
    },
}


def compute_question_score(
    has_pivot: bool,
    pivot_match: bool,
    structural_issues: list[str],
    explanation_deduct: bool,
) -> tuple[float, list[str]]:
    if not has_pivot:
        return 0.0, ["No pivot table submitted."]

    score = 1.0
    comments: list[str] = []

    for issue_key in structural_issues:
        if issue_key in DEDUCTIONS:
            score += DEDUCTIONS[issue_key]["points"]
            comments.append(DEDUCTIONS[issue_key]["comment"])

    if not pivot_match:
        score += DEDUCTIONS["wrong_values"]["points"]
        comments.append(DEDUCTIONS["wrong_values"]["comment"])

    if explanation_deduct:
        score += DEDUCTIONS["bad_explanation"]["points"]
        comments.append(DEDUCTIONS["bad_explanation"]["comment"])

    return max(0.0, round(score, 1)), comments
