from __future__ import annotations

import atexit
import json
import os
import re
import signal
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
from grader.scoring import compute_question_score, format_short_comments

EXPLANATION_QUESTIONS = {"Q4", "Q6", "Q7", "Q9"}
NEEDS_REVIEW = "NEEDS_REVIEW"

_MIN_MATCH_CONFIDENCE = 3.0
_TRACE_MODE = os.getenv("PIVOT_TRACE_MODE", "0").strip().lower() in {"1", "true", "yes", "on"}
_TRACE_TIMEOUT_SEC = int(os.getenv("PIVOT_TRACE_TIMEOUT_SEC", "0") or "0")
_TRACE_LAST_STEP = "init"
_TRACE_COMPLETED = False


def _trace_enabled() -> bool:
    return _TRACE_MODE


def _trace_timestamp() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _trace(step: str, message: str = "") -> None:
    global _TRACE_LAST_STEP
    if not _trace_enabled():
        return
    _TRACE_LAST_STEP = step
    suffix = f" | {message}" if message else ""
    print(f"[TRACE {_trace_timestamp()}] {step}{suffix}")


def _trace_branch(qid: str, branch: str, message: str = "") -> None:
    if not _trace_enabled():
        return
    detail = f" | {message}" if message else ""
    print(f"[TRACE {_trace_timestamp()}] branch {qid}: {branch}{detail}")


def _trace_last_completed(reason: str) -> None:
    if not _trace_enabled():
        return
    print(f"[TRACE {_trace_timestamp()}] LAST_COMPLETED_STEP ({reason}): {_TRACE_LAST_STEP}")


def _setup_trace_guards() -> None:
    if not _trace_enabled():
        return

    def _signal_handler(signum: int, _frame: Any) -> None:
        reason = f"signal={signum}"
        if signum == signal.SIGALRM:
            reason = "timeout"
        _trace_last_completed(reason)
        raise SystemExit(124)

    for sig in (signal.SIGINT, signal.SIGTERM):
        signal.signal(sig, _signal_handler)
    signal.signal(signal.SIGALRM, _signal_handler)

    if _TRACE_TIMEOUT_SEC > 0:
        signal.alarm(_TRACE_TIMEOUT_SEC)


@atexit.register
def _trace_on_exit() -> None:
    if _trace_enabled() and not _TRACE_COMPLETED:
        _trace_last_completed("process_exit")


def _match_debug_enabled() -> bool:
    return os.getenv("PIVOT_MATCH_DEBUG", "0").strip().lower() in {"1", "true", "yes", "on"}


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


def _match_sheets_to_questions(
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
    if _match_debug_enabled():
        print("[match] Student sheet names:", sorted(sheets.keys()))
    _trace("question mapping", f"begin mapping for {len(question_ids)} questions")

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
            _trace("question mapping", f"{qid} -> {best_name} (direct name match)")
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
            _trace(
                "question mapping",
                f"{qid} -> {best_name} (fallback path, similarity={best_score:.1f})",
            )
        else:
            matched[qid] = None
            matched_names[qid] = None
            warnings.append(
                f"{qid}: no sheet matched confidently (best_score={best_score:.1f}). "
                "Flagged for human review — grade manually."
            )
            _trace(
                "question mapping",
                f"{qid} unmapped question (best similarity={best_score:.1f})",
            )

    if _match_debug_enabled():
        final_map = {qid: matched_names.get(qid) for qid in question_ids}
        print("[match] Final Q->Sheet mapping:", final_map)

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
) -> tuple[bool, list[str], bool]:
    notes: list[str] = []
    needs_review = False
    _trace("pivot extraction", f"{qid} start")

    # Q8 — highlight detection via openpyxl; skip DataFrame comparison entirely
    if qid == "Q8":
        if workbook_path and sheet_name:
            q8 = check_q8_highlight(workbook_path, sheet_name)
        else:
            q8 = {
                "match": False,
                "missing_pivot": False,
                "needs_review": True,
                "notes": [
                    f"{NEEDS_REVIEW}: workbook path unavailable for Q8 highlight detection"
                ],
            }
        needs_review = bool(q8.get("needs_review", False))
        notes.extend(q8["notes"])
        _trace("pivot extraction", f"{qid} complete (match={q8['match']}, needs_review={needs_review})")
        return q8["match"], notes, needs_review

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
            _trace("pivot extraction", f"{qid} filter mismatch")
            return False, notes, needs_review

    if qid == "Q5":
        fc = check_q5_filter(student_df)
        if not fc["filter_ok"]:
            notes.append(
                f"Month filter incorrect — non-summer months found: "
                f"{', '.join(fc['non_summer_months_found'])}."
            )
            _trace("pivot extraction", f"{qid} filter mismatch")
            return False, notes, needs_review

    if qid == "Q10":
        notes.append(f"{NEEDS_REVIEW}: Q10 month-filter validation is incomplete")
        _trace("pivot extraction", f"{qid} requires manual review")
        return False, notes, True

    result = compare_pivot_values(student_df, answer_df)

    if not result["match"] and result["mismatches"]:
        first = result["mismatches"][0]
        notes.append(
            f"Mismatch at {first['label']}: expected {first['expected']}, actual {first['actual']}."
        )

    _trace("pivot extraction", f"{qid} complete (match={bool(result['match'])})")

    return bool(result["match"]), notes, needs_review


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
    comment_max_words: int = 15,
) -> tuple[float | None, str]:
    _trace("per-question grading", f"{qid} start")
    answer_df = answer_key.get(qid)

    if student_df is None or student_df.empty:
        if review_flagged:
            _trace_branch(qid, "unmapped question", "NEEDS_REVIEW")
            return None, f"{NEEDS_REVIEW}: unmapped question"
        score, comments = compute_question_score(
            has_pivot=False,
            pivot_match=False,
            structural_issues=[],
            explanation_deduct=False,
        )
        _trace_branch(qid, "missing pivot", "no student sheet")
        return score, format_short_comments(comments, max_words=comment_max_words)

    if answer_df is None:
        _trace_branch(qid, "NEEDS_REVIEW", "answer key sheet missing")
        return None, f"{NEEDS_REVIEW}: answer key sheet missing"

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
            pivot_notes.append("Missing pivot table")
        elif not pivot_match:
            pivot_notes.append("Incorrect value")
        explanation_deduct = False
        explanation_note = ""
        if question_cfg.get("explanation_required"):
            _trace("explanation grading", f"{qid} start")
            rubric_text = question_cfg.get("explanation_rubric", "")
            student_text = _extract_explanation_text(student_df)
            llm_result = grade_explanation(qid, student_text, rubric_text)
            if llm_result.get("needs_review", False):
                _trace_branch(qid, "NEEDS_REVIEW", "explanation grading unavailable")
                return None, str(llm_result.get("brief_reason", f"{NEEDS_REVIEW}: explanation grading unavailable"))
            explanation_deduct = bool(llm_result["deduct_explanation"])
            explanation_note = str(llm_result.get("brief_reason", ""))
            _trace("explanation grading", f"{qid} complete (deduct={explanation_deduct})")
        score, comments = compute_question_score(
            has_pivot=has_pivot,
            pivot_match=pivot_match,
            structural_issues=[],
            explanation_deduct=explanation_deduct,
        )
        all_comments = comments + pivot_notes
        if explanation_deduct and explanation_note:
            all_comments.append(explanation_note)
        if score == 1.0:
            _trace_branch(qid, "perfect")
        else:
            _trace_branch(qid, "deduction", f"score={score}")
        return score, format_short_comments(all_comments, max_words=comment_max_words)

    structural_issues = _evaluate_structural_issues(qid, student_df, question_cfg)
    pivot_match, pivot_notes, pivot_needs_review = _evaluate_pivot_match(
        qid, student_df, answer_df, question_cfg,
        workbook_path=workbook_path, sheet_name=sheet_name,
    )

    if pivot_needs_review:
        for note in pivot_notes:
            if NEEDS_REVIEW in note:
                _trace_branch(qid, "NEEDS_REVIEW", note)
                return None, note
        _trace_branch(qid, "NEEDS_REVIEW", "incomplete validation")
        return None, f"{NEEDS_REVIEW}: incomplete validation"

    explanation_deduct = False
    explanation_note = ""

    if qid in EXPLANATION_QUESTIONS and question_cfg.get("explanation_required"):
        _trace("explanation grading", f"{qid} start")
        rubric_text = question_cfg.get("explanation_rubric", "")
        student_text = _extract_explanation_text(student_df)
        llm_result = grade_explanation(qid, student_text, rubric_text)
        if llm_result.get("needs_review", False):
            _trace_branch(qid, "NEEDS_REVIEW", "explanation grading unavailable")
            return None, str(llm_result.get("brief_reason", f"{NEEDS_REVIEW}: explanation grading unavailable"))
        explanation_deduct = bool(llm_result["deduct_explanation"])
        explanation_note = str(llm_result.get("brief_reason", ""))
        _trace("explanation grading", f"{qid} complete (deduct={explanation_deduct})")

    score, comments = compute_question_score(
        has_pivot=True,
        pivot_match=pivot_match,
        structural_issues=structural_issues,
        explanation_deduct=explanation_deduct,
    )

    all_comments = comments + pivot_notes
    if explanation_deduct and explanation_note:
        all_comments.append(explanation_note)

    if score == 1.0:
        _trace_branch(qid, "perfect")
    else:
        _trace_branch(qid, "deduction", f"score={score}")

    return score, format_short_comments(all_comments, max_words=comment_max_words)


def run_all(
    submissions_dir: str | Path,
    answer_key_path: str | Path,
    template_path: str | Path,
    output_dir: str | Path,
) -> dict[str, Any]:
    """Grade all student submissions and emit per-student logs + run summary."""
    global _TRACE_COMPLETED
    _setup_trace_guards()
    _trace("submission discovery", "run_all start")

    submissions_path = Path(submissions_dir)
    output_root = Path(output_dir)
    logs_dir = output_root / "logs"
    grades_dir = output_root / "grades"
    logs_dir.mkdir(parents=True, exist_ok=True)
    grades_dir.mkdir(parents=True, exist_ok=True)

    rubric_path = Path(__file__).resolve().parents[1] / "configs" / "rubric.json"
    rubric = json.loads(rubric_path.read_text(encoding="utf-8"))
    comment_cfg = rubric.get("comment_style", {})
    comment_max_words = int(comment_cfg.get("max_words_per_comment", 15))

    answer_key = load_answer_key(answer_key_path)

    summary: dict[str, Any] = {
        "timestamp_utc": datetime.now(timezone.utc).isoformat(),
        "assignment": rubric.get("assignment"),
        "max_score": rubric.get("max_score", 10),
        "students": [],
        "success_count": 0,
        "failure_count": 0,
    }

    # Prefer direct workbook discovery so classes that submit many files inside
    # a single folder are processed one workbook per student.
    student_workbooks = sorted(
        [
            p for p in submissions_path.rglob("*.xlsx")
            if not p.name.startswith("~$") and "grade" not in p.name.lower()
        ]
    )
    submission_targets: list[Path]
    if student_workbooks:
        submission_targets = student_workbooks
    else:
        submission_targets = sorted([p for p in submissions_path.iterdir() if p.is_dir()])
    _trace("submission discovery", f"found {len(submission_targets)} target(s)")

    single_submission_mode = _trace_enabled()
    processed_first_valid = False
    for target in submission_targets:
        if single_submission_mode and processed_first_valid:
            _trace("submission discovery", "single-submission trace mode active; skipping remaining targets")
            break

        _trace("submission discovery", f"loading target {target}")
        submission = load_student_submission(target)
        if not submission.error:
            processed_first_valid = True
        log_path = logs_dir / f"{submission.student_id}.log"

        try:
            if submission.error:
                _trace("sheet loading", f"{submission.student_id} invalid submission: {submission.error}")
                raise ValueError(submission.error)

            _trace("sheet loading", f"{submission.student_id} loaded {len(submission.sheets)} sheet(s)")

            scores: dict[str, float | None] = {}
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
                    comment_max_words=comment_max_words,
                )
                scores[qid] = score
                comments[qid] = note

            _trace("score calculation", f"{submission.student_id} computing total")
            _trace("output writing", f"{submission.student_id} writing grade workbook")
            grade_file = write_grades(
                student_id=submission.student_id,
                scores_dict=scores,
                comments_dict=comments,
                template_path=template_path,
                output_dir=grades_dir,
                max_comment_words=comment_max_words,
            )
            _trace("output writing", f"{submission.student_id} wrote {grade_file}")

            total = round(sum(v for v in scores.values() if v is not None), 1)
            needs_review_questions = sorted([qid for qid, val in scores.items() if val is None])
            log_lines = [
                f"student_id={submission.student_id}",
                f"status={'needs_review' if needs_review_questions else 'success'}",
                f"total_score={total}",
                f"grade_file={grade_file}",
            ]
            if needs_review_questions:
                log_lines.append(f"needs_review_questions={','.join(needs_review_questions)}")
            for w in match_warnings:
                log_lines.append(f"WARN: {w}")
            for qid in sorted(scores):
                log_lines.append(f"{qid}: {scores[qid]} | {comments[qid]}")

            log_path.write_text("\n".join(log_lines) + "\n", encoding="utf-8")

            summary["students"].append(
                {
                    "student_id": submission.student_id,
                    "status": "needs_review" if needs_review_questions else "success",
                    "total_score": total,
                    "grade_file": str(grade_file),
                    "needs_review_questions": needs_review_questions,
                }
            )
            summary["success_count"] += 1
            _trace(
                "final result",
                f"student_id={submission.student_id} status={'needs_review' if needs_review_questions else 'success'} total={total}",
            )

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
            _trace("final result", f"student_id={submission.student_id} status=failure error={exc}")

    summary_path = logs_dir / "run_summary.json"
    summary_path.write_text(json.dumps(summary, indent=2), encoding="utf-8")

    if _trace_enabled() and _TRACE_TIMEOUT_SEC > 0:
        signal.alarm(0)
    _TRACE_COMPLETED = True
    return summary


if __name__ == "__main__":
    repo_root = Path(__file__).resolve().parents[1]
    run_all(
        submissions_dir=repo_root / "submissions",
        answer_key_path=repo_root / "answer_key" / "GryzzlSales2024_Answer_Key.xlsx",
        template_path=repo_root / "templates" / "Homework_3_Gradesheet_Template.xlsx",
        output_dir=repo_root / "outputs",
    )
