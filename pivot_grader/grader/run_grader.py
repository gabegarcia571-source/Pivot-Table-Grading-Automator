from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import pandas as pd

from grader.grade_writer import write_grades
from grader.ingest import SubmissionResult, load_answer_key, load_student_submission
from grader.pivot_checker import (
    check_q3_filter,
    check_q4_average,
    check_q5_filter,
    check_q8_highlight,
    check_q10_filter,
    compare_pivot_values,
    fingerprint_similarity,
    is_desc_sorted,
    is_desc_sorted_within_groups,
    sheet_fingerprint,
)
from grader.qualitative_grader import grade_explanation
from grader.scoring import compute_question_score

EXPLANATION_QUESTIONS = {"Q4", "Q6", "Q7", "Q9"}

_MIN_MATCH_CONFIDENCE = 3.0


def _match_sheets_to_questions(
    sheets: dict[str, pd.DataFrame],
    answer_key: dict[str, pd.DataFrame],
    question_ids: list[str],
    warnings: list[str],
) -> tuple[dict[str, pd.DataFrame | None], dict[str, str | None]]:
    """Match student sheets to question IDs by content fingerprint.

    Each student sheet is assigned to at most one question.  If no sheet scores
    above _MIN_MATCH_CONFIDENCE for a question, the question is flagged for
    human review rather than automatically failed.

    Returns a tuple of (matched_dfs, matched_sheet_names).
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

    for qid in question_ids:
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


def _extract_explanation_text(df: pd.DataFrame) -> str:
    text_values: list[str] = []
    for col in df.columns:
        series = df[col].dropna().astype(str)
        for value in series:
            value = value.strip()
            if len(value.split()) >= 6:  # avoid single labels and noise
                text_values.append(value)
    return "\n".join(text_values)


def _evaluate_pivot_match(
    qid: str,
    student_df: pd.DataFrame,
    answer_df: pd.DataFrame,
    question_cfg: dict[str, Any],
    workbook_path: Any = None,
    sheet_name: str | None = None,
) -> tuple[bool, list[str]]:
    notes: list[str] = []

    # Q8 — highlight detection via openpyxl; skip DataFrame comparison entirely
    if qid == "Q8":
        if workbook_path and sheet_name:
            q8 = check_q8_highlight(workbook_path, sheet_name)
        else:
            q8 = {
                "match": False,
                "missing_pivot": False,
                "notes": [
                    "Workbook path unavailable for Q8 highlight detection — grade manually."
                ],
            }
        notes.extend(q8["notes"])
        return q8["match"], notes

    # Filter verification — return early on failure (wrong_values deduction applies)
    if qid == "Q3":
        fc = check_q3_filter(student_df)
        if not fc["filter_ok"]:
            missing = fc.get("missing_vendors", [])
            extra = fc.get("extra_vendors", [])
            parts: list[str] = []
            if extra:
                parts.append(f"unexpected vendor(s) in rows: {', '.join(extra)}")
            if missing:
                parts.append(f"missing top-3 vendor(s): {', '.join(missing)}")
            notes.append(
                f"Vendor filter incorrect — {'; '.join(parts) or 'wrong vendor row labels'}."
            )
            return False, notes

    if qid == "Q5":
        fc = check_q5_filter(student_df)
        if not fc["filter_ok"]:
            notes.append(
                f"Month filter incorrect — non-summer months found: "
                f"{', '.join(fc['non_summer_months_found'])}."
            )
            return False, notes

    if qid == "Q10":
        fc = check_q10_filter(student_df)
        if not fc["filter_ok"]:
            parts2: list[str] = []
            if not fc.get("vendor_count_ok"):
                parts2.append(
                    f"expected 7 vendor rows, found {fc.get('vendor_count', '?')}"
                )
            if not fc.get("has_honey_col"):
                parts2.append("missing 'Honey' column")
            if not fc.get("has_no_promo_col"):
                parts2.append("missing 'No Promo Code' column")
            if not fc.get("values_in_range"):
                parts2.append("cell values are not proportions (expected 0–1)")
            notes.append(
                f"Q10 structure incorrect — {'; '.join(parts2) or 'wrong structure'}. "
                "NOTE: month filter validation requires manual review."
            )
            return False, notes

    result = compare_pivot_values(student_df, answer_df)

    if not result["match"] and result["mismatches"]:
        first = result["mismatches"][0]
        notes.append(
            f"Mismatch at {first['label']}: expected {first['expected']}, actual {first['actual']}."
        )

    return bool(result["match"]), notes


def _evaluate_structural_issues(qid: str, student_df: pd.DataFrame, question_cfg: dict[str, Any]) -> list[str]:
    issues: list[str] = []

    if question_cfg.get("sort_required"):
        # Q1 is a nested pivot (category × day); check sort within each category group
        if qid == "Q1":
            sort_ok = is_desc_sorted_within_groups(student_df)
        else:
            sort_ok = is_desc_sorted(student_df)
        if not sort_ok:
            issues.append("no_sort")

    return issues


def _grade_one_question(
    qid: str,
    question_cfg: dict[str, Any],
    student_df: pd.DataFrame | None,
    answer_key: dict[str, pd.DataFrame],
    review_flagged: bool = False,
    workbook_path: Any = None,
    sheet_name: str | None = None,
) -> tuple[float, str]:
    answer_df = answer_key.get(qid)

    if student_df is None or student_df.empty:
        if review_flagged:
            return 0.0, (
                "HUMAN REVIEW NEEDED — no sheet could be matched to this question "
                "automatically. Grade manually."
            )
        score, comments = compute_question_score(
            has_pivot=False,
            pivot_match=False,
            structural_issues=[],
            explanation_deduct=False,
        )
        return score, " ".join(comments)

    if answer_df is None:
        score, comments = compute_question_score(
            has_pivot=False,
            pivot_match=False,
            structural_issues=[],
            explanation_deduct=False,
        )
        return score, " ".join(comments)

    # ------------------------------------------------------------------
    # Q4 — special numeric scan instead of structural pivot comparison.
    # The sheet contains ~28,990 per-order rows; we only need to verify
    # that the student computed the correct average (~46.48) somewhere.
    # ------------------------------------------------------------------
    if qid == "Q4":
        q4 = check_q4_average(student_df)
        has_pivot = q4["has_numeric"]
        pivot_match = q4["match"]
        pivot_notes: list[str] = []
        if not has_pivot:
            pivot_notes.append("No numeric values found on Q4 sheet.")
        elif not pivot_match:
            pivot_notes.append(
                f"Average not near expected ~46.48. "
                f"Closest value found: {q4['closest']:.4f} (delta {q4['delta']:.4f}). "
                "Student may have averaged individual line items instead of per-order totals."
            )
        explanation_deduct = False
        explanation_note = ""
        if question_cfg.get("explanation_required"):
            rubric_text = question_cfg.get("explanation_rubric", "")
            student_text = _extract_explanation_text(student_df)
            llm_result = grade_explanation(qid, student_text, rubric_text)
            explanation_deduct = bool(llm_result["deduct_explanation"])
            explanation_note = str(llm_result.get("brief_reason", ""))
        score, comments = compute_question_score(
            has_pivot=has_pivot,
            pivot_match=pivot_match,
            structural_issues=[],
            explanation_deduct=explanation_deduct,
        )
        all_comments = comments + pivot_notes
        if explanation_note:
            all_comments.append(f"Explanation check: {explanation_note}")
        return score, " ".join(all_comments)

    structural_issues = _evaluate_structural_issues(qid, student_df, question_cfg)
    pivot_match, pivot_notes = _evaluate_pivot_match(
        qid, student_df, answer_df, question_cfg,
        workbook_path=workbook_path, sheet_name=sheet_name,
    )

    explanation_deduct = False
    explanation_note = ""

    if qid in EXPLANATION_QUESTIONS and question_cfg.get("explanation_required"):
        rubric_text = question_cfg.get("explanation_rubric", "")
        student_text = _extract_explanation_text(student_df)
        llm_result = grade_explanation(qid, student_text, rubric_text)
        explanation_deduct = bool(llm_result["deduct_explanation"])
        explanation_note = str(llm_result.get("brief_reason", ""))

    score, comments = compute_question_score(
        has_pivot=True,
        pivot_match=pivot_match,
        structural_issues=structural_issues,
        explanation_deduct=explanation_deduct,
    )

    all_comments = comments + pivot_notes
    if explanation_note:
        all_comments.append(f"Explanation check: {explanation_note}")

    return score, " ".join(all_comments)


def run_all(
    submissions_dir: str | Path,
    answer_key_path: str | Path,
    template_path: str | Path,
    output_dir: str | Path,
) -> dict[str, Any]:
    """Grade all student folders and emit per-student logs + run summary."""
    submissions_path = Path(submissions_dir)
    output_root = Path(output_dir)
    logs_dir = output_root / "logs"
    grades_dir = output_root / "grades"
    logs_dir.mkdir(parents=True, exist_ok=True)
    grades_dir.mkdir(parents=True, exist_ok=True)

    rubric_path = Path(__file__).resolve().parents[1] / "configs" / "rubric.json"
    rubric = json.loads(rubric_path.read_text(encoding="utf-8"))

    answer_key = load_answer_key(answer_key_path)

    summary: dict[str, Any] = {
        "timestamp_utc": datetime.now(timezone.utc).isoformat(),
        "assignment": rubric.get("assignment"),
        "max_score": rubric.get("max_score", 10),
        "students": [],
        "success_count": 0,
        "failure_count": 0,
    }

    student_folders = sorted([p for p in submissions_path.iterdir() if p.is_dir()])

    for folder in student_folders:
        submission = load_student_submission(folder)
        log_path = logs_dir / f"{submission.student_id}.log"

        try:
            if submission.error:
                raise ValueError(submission.error)

            scores: dict[str, float] = {}
            comments: dict[str, str] = {}

            match_warnings: list[str] = []
            matched_sheets, matched_names = _match_sheets_to_questions(
                submission.sheets,
                answer_key,
                [q["id"] for q in rubric.get("questions", [])],
                match_warnings,
            )

            for question_cfg in rubric.get("questions", []):
                qid = question_cfg["id"]
                student_df = matched_sheets.get(qid)
                review_flagged = any(w.startswith(f"{qid}:") for w in match_warnings)
                score, note = _grade_one_question(
                    qid, question_cfg, student_df, answer_key, review_flagged,
                    workbook_path=submission.workbook_path,
                    sheet_name=matched_names.get(qid),
                )
                scores[qid] = score
                comments[qid] = note

            grade_file = write_grades(
                student_id=submission.student_id,
                scores_dict=scores,
                comments_dict=comments,
                template_path=template_path,
                output_dir=grades_dir,
            )

            total = round(sum(scores.values()), 1)
            log_lines = [
                f"student_id={submission.student_id}",
                f"status=success",
                f"total_score={total}",
                f"grade_file={grade_file}",
            ]
            for w in match_warnings:
                log_lines.append(f"WARN: {w}")
            for qid in sorted(scores):
                log_lines.append(f"{qid}: {scores[qid]} | {comments[qid]}")

            log_path.write_text("\n".join(log_lines) + "\n", encoding="utf-8")

            summary["students"].append(
                {
                    "student_id": submission.student_id,
                    "status": "success",
                    "total_score": total,
                    "grade_file": str(grade_file),
                }
            )
            summary["success_count"] += 1

        except Exception as exc:  # noqa: BLE001
            log_path.write_text(
                f"student_id={submission.student_id}\nstatus=failure\nerror={exc}\n",
                encoding="utf-8",
            )
            summary["students"].append(
                {
                    "student_id": submission.student_id,
                    "status": "failure",
                    "error": str(exc),
                }
            )
            summary["failure_count"] += 1

    summary_path = logs_dir / "run_summary.json"
    summary_path.write_text(json.dumps(summary, indent=2), encoding="utf-8")
    return summary


if __name__ == "__main__":
    repo_root = Path(__file__).resolve().parents[1]
    run_all(
        submissions_dir=repo_root / "submissions",
        answer_key_path=repo_root / "answer_key" / "GryzzlSales2024_Answer_Key.xlsx",
        template_path=repo_root / "templates" / "Homework_3_Gradesheet_Template.xlsx",
        output_dir=repo_root / "outputs",
    )
