from __future__ import annotations

from typing import Any

DEDUCTIONS = {
    "missing_pivot": {"points": -1.0, "comment": "Missing pivot table"},
    "no_sort": {"points": -0.3, "comment": "Incorrect sort"},
    "incorrect_filter": {"points": -0.3, "comment": "Incorrect filter"},
    "extra_columns": {
        "points": -0.3,
        "comment": "Extra columns",
    },
    "answer_not_highlighted": {
        "points": -0.5,
        "comment": "Missing highlight",
    },
    "wrong_values": {
        "points": -0.7,
        "comment": "Incorrect value",
    },
    "bad_explanation": {
        "points": -0.3,
        "comment": "Needs more detail",
    },
}


def _limit_words(text: str, max_words: int) -> str:
    words = text.split()
    if len(words) <= max_words:
        return text
    return " ".join(words[:max_words])


def _normalize_comment(text: str) -> str:
    low = text.strip().lower()
    if not low:
        return ""
    if "highlight" in low:
        return "Missing highlight"
    if "filter" in low:
        return "Incorrect filter"
    if "sort" in low:
        return "Incorrect sort"
    if "inconsistent" in low and "analysis" in low:
        return "Answer inconsistent with analysis"
    if "directly address" in low or "off-topic" in low:
        return "Should more directly address question"
    if "explanation" in low:
        return "Needs more detail"
    if "manual review" in low:
        return "Manual review needed"
    if "pivot" in low and ("missing" in low or "no" in low):
        return "Missing pivot table"
    if "mismatch" in low or "expected" in low or "incorrect" in low or "value" in low:
        return "Incorrect value"
    return text.strip()


def format_short_comments(comments: list[str], max_words: int = 15) -> str:
    """Normalize comments into short rubric-style phrases.

    Removes duplicates, keeps first-seen order, and joins with semicolons.
    """
    seen: set[str] = set()
    out: list[str] = []
    for raw in comments:
        norm = _normalize_comment(raw)
        if not norm:
            continue
        norm = _limit_words(norm, max_words=max_words)
        if norm in seen:
            continue
        seen.add(norm)
        out.append(norm)
    return "; ".join(out)


def assemble_score(contract: dict[str, Any]) -> tuple[float, list[str]]:
    """Assemble final score/comments from question-module contract output.

    Uses the same deduction magnitudes as compute_question_score:
      - structural: 0.3
      - value: 0.7
      - explanation: 0.3
    Formatting carries the highlight deduction weight from the rubric.
    """
    structural_score = float(contract.get("structural_score", 0.0))
    value_score = float(contract.get("value_score", 0.0))
    formatting_score = float(contract.get("formatting_score", 1.0))
    explanation_score = float(contract.get("explanation_score", 1.0))

    # Clamp to [0, 1] to keep scoring stable even if a module returns out-of-range values.
    structural_score = max(0.0, min(1.0, structural_score))
    value_score = max(0.0, min(1.0, value_score))
    formatting_score = max(0.0, min(1.0, formatting_score))
    explanation_score = max(0.0, min(1.0, explanation_score))

    structural_weight = abs(float(DEDUCTIONS["incorrect_filter"]["points"]))
    value_weight = abs(float(DEDUCTIONS["wrong_values"]["points"]))
    explanation_weight = abs(float(DEDUCTIONS["bad_explanation"]["points"]))
    formatting_weight = abs(float(DEDUCTIONS["answer_not_highlighted"]["points"]))

    score = 1.0
    score -= structural_weight * (1.0 - structural_score)
    score -= value_weight * (1.0 - value_score)
    score -= explanation_weight * (1.0 - explanation_score)
    score -= formatting_weight * (1.0 - formatting_score)

    comments: list[str] = []
    for key in (
        "structural_issues",
        "value_issues",
        "formatting_issues",
        "explanation_issues",
    ):
        for issue in contract.get(key, []) or []:
            text = str(issue).strip()
            if text:
                comments.append(text)

    return max(0.0, round(score, 1)), comments
