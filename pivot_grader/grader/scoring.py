from __future__ import annotations

DEDUCTIONS = {
    "missing_pivot": {"points": -1.0, "comment": "Missing pivot table"},
    "no_sort": {"points": -0.3, "comment": "Incorrect sort"},
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


def compute_question_score(
    has_pivot: bool,
    pivot_match: bool,
    structural_issues: list[str],
    explanation_deduct: bool,
) -> tuple[float, list[str]]:
    if not has_pivot:
        return 0.0, [DEDUCTIONS["missing_pivot"]["comment"]]

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
