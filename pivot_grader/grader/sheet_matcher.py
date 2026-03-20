from __future__ import annotations

import re

import pandas as pd

from grader.pivot_checker import fingerprint_similarity, sheet_fingerprint

_MIN_MATCH_CONFIDENCE = 3.0


def _normalize_name(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", value.lower())


def _qnum(qid: str) -> str:
    return re.sub(r"[^0-9]", "", qid)


def _sheet_match_rank(sheet_name: str, qid: str) -> int | None:
    """Return a rank (lower is better) for sheet_name matching qid.

    0: exact normalized match (e.g., Q1, q1, Q1 )
    1: explicit question pattern (e.g., Question 1, Question1)
    2: contains q<num> token anywhere
    None: no match
    """
    qnum = _qnum(qid)
    if not qnum:
        return None

    raw = sheet_name.strip().lower()
    norm = _normalize_name(raw)
    if norm == f"q{qnum}":
        return 0
    if norm == f"question{qnum}":
        return 1
    if re.search(rf"\bq\s*0*{qnum}\b", raw):
        return 2
    if re.search(rf"\bquestion\s*0*{qnum}\b", raw):
        return 2
    return None


def match_sheets_to_questions(
    sheets: dict[str, pd.DataFrame],
    answer_key: dict[str, pd.DataFrame],
    question_ids: list[str],
    warnings: list[str],
) -> tuple[dict[str, pd.DataFrame | None], dict[str, str | None]]:
    """Match student sheets to question IDs.

    Priority:
      1) Direct/simple sheet-name matching (Q1, q1, Question 1, etc.)
      2) Fingerprint fallback if no name match found.
    """
    student_fps = {name: sheet_fingerprint(df) for name, df in sheets.items()}
    answer_fps = {
        qid: sheet_fingerprint(df)
        for qid, df in answer_key.items()
        if not df.empty
    }

    matched: dict[str, pd.DataFrame | None] = {}
    matched_names: dict[str, str | None] = {}
    used: set[str] = set()

    # 1) Prefer simple, reliable name-based mapping.
    for qid in question_ids:
        best_name: str | None = None
        best_rank: int | None = None
        for name in sheets:
            if name in used:
                continue
            rank = _sheet_match_rank(name, qid)
            if rank is None:
                continue
            if best_rank is None or rank < best_rank:
                best_rank = rank
                best_name = name

        if best_name is not None:
            matched[qid] = sheets[best_name]
            matched_names[qid] = best_name
            used.add(best_name)
        else:
            matched[qid] = None
            matched_names[qid] = None

    # 2) Fallback to fingerprint similarity for anything still unmatched.
    for qid in question_ids:
        if matched.get(qid) is not None:
            continue

        answer_fp = answer_fps.get(qid)
        if answer_fp is None:
            matched[qid] = None
            matched_names[qid] = None
            continue

        best_name: str | None = None
        best_score: float = 0.0
        for name, sfp in student_fps.items():
            if name in used:
                continue
            score = fingerprint_similarity(sfp, answer_fp)
            if score > best_score:
                best_score = score
                best_name = name

        if best_name and best_score >= _MIN_MATCH_CONFIDENCE:
            matched[qid] = sheets[best_name]
            matched_names[qid] = best_name
            used.add(best_name)
        else:
            matched[qid] = None
            matched_names[qid] = None
            warnings.append(
                f"{qid}: no sheet matched confidently (best_score={best_score:.1f}). "
                "Flagged for human review — grade manually."
            )

    return matched, matched_names
