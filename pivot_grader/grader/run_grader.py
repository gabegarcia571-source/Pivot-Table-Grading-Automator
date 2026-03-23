from __future__ import annotations

import argparse
import importlib
import inspect
import atexit
import json
import os
import signal
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
import argparse

import pandas as pd

from grader.grade_writer import write_grades
from grader.ingest import load_answer_key, load_student_submission
from grader.sheet_matcher import match_sheets_to_questions
from grader.scoring import assemble_score, format_short_comments

NEEDS_REVIEW = "NEEDS_REVIEW"

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


def _qnum(qid: str) -> str:
    return "".join(ch for ch in qid if ch.isdigit())


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
        _trace_branch(qid, "missing pivot", "no student sheet")
        return 0.0, "Missing pivot table"

    if answer_df is None:
        _trace_branch(qid, "NEEDS_REVIEW", "answer key sheet missing")
        return None, f"{NEEDS_REVIEW}: answer key sheet missing"

    qnum = _qnum(qid)
    if not qnum:
        _trace_branch(qid, "NEEDS_REVIEW", "invalid question id format")
        return None, f"{NEEDS_REVIEW}: invalid question id {qid}"

    module_name = f"grader.questions.q{qnum}"
    try:
        module = importlib.import_module(module_name)
    except Exception as exc:  # noqa: BLE001
        _trace_branch(qid, "NEEDS_REVIEW", f"import failure: {exc}")
        return None, f"{NEEDS_REVIEW}: failed to import {module_name}: {exc}"

    grade_fn = getattr(module, "grade_question", None)
    if grade_fn is None:
        _trace_branch(qid, "NEEDS_REVIEW", "grade_question missing")
        return None, f"{NEEDS_REVIEW}: {module_name}.grade_question missing"

    params = inspect.signature(grade_fn).parameters
    call_kwargs: dict[str, Any] = {
        "student_df": student_df,
        "answer_df": answer_df,
        "question_cfg": question_cfg,
    }
    if "qid" in params:
        call_kwargs["qid"] = qid
    if "workbook_path" in params:
        call_kwargs["workbook_path"] = workbook_path
    if "sheet_name" in params:
        call_kwargs["sheet_name"] = sheet_name

    try:
        contract = grade_fn(**call_kwargs)
    except TypeError:
        # Compatibility fallback for question modules that still expose a 3-arg signature.
        contract = grade_fn(student_df, answer_df, question_cfg)

    if not isinstance(contract, dict):
        _trace_branch(qid, "NEEDS_REVIEW", "invalid contract type")
        return None, f"{NEEDS_REVIEW}: invalid grading contract for {qid}"

    # Structural/value NEEDS_REVIEW requires full manual grading.
    for key in ("value_issues", "structural_issues"):
        for issue in (contract.get(key) or []):
            issue_text = str(issue)
            if issue_text.startswith(NEEDS_REVIEW):
                _trace_branch(qid, "NEEDS_REVIEW", issue_text)
                return None, issue_text

    explanation_issues = [str(v) for v in (contract.get("explanation_issues") or [])]
    if any(issue.startswith(NEEDS_REVIEW) for issue in explanation_issues):
        structural_score = max(0.0, min(1.0, float(contract.get("structural_score", 0.0))))
        value_score = max(0.0, min(1.0, float(contract.get("value_score", 0.0))))
        formatting_score = max(0.0, min(1.0, float(contract.get("formatting_score", 1.0))))
        partial = 1.0
        partial -= 0.3 * (1.0 - structural_score)
        partial -= 0.7 * (1.0 - value_score)
        partial -= 0.5 * (1.0 - formatting_score)
        partial = max(0.0, round(partial, 1))
        message = f"{NEEDS_REVIEW}: explanation pending manual grade"
        _trace_branch(qid, "partial review", f"score={partial}")
        return partial, message

    score, comments = assemble_score(contract)
    short_comment = format_short_comments(comments, max_words=comment_max_words)

    if score == 1.0:
        _trace_branch(qid, "perfect")
    else:
        _trace_branch(qid, "deduction", f"score={score}")

    return score, short_comment


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
    submission_targets: list[Path]
    if submissions_path.is_file():
        submission_targets = [submissions_path]
    else:
        student_workbooks = sorted(
            [
                p for p in submissions_path.rglob("*.xlsx")
                if not p.name.startswith("~$") and "grade" not in p.name.lower()
            ]
        )
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
            matched_sheets, matched_names = match_sheets_to_questions(
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


def _first_existing_path(candidates: list[Path]) -> Path:
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0]


def _parse_args(repo_root: Path) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run pivot-table grader for one file or a submissions folder.")
    parser.add_argument(
        "--student",
        "--submissions",
        dest="submissions_dir",
        default=str(repo_root / "submissions"),
        help="Path to a student workbook (.xlsx) or folder containing submissions.",
    )
    parser.add_argument(
        "--answer_key",
        default=str(
            _first_existing_path(
                [
                    repo_root / "answer_key" / "GryzzlSales2024 - Answer Key.xlsx",
                    repo_root / "answer_key" / "GryzzlSales2024_Answer_Key.xlsx",
                ]
            )
        ),
        help="Path to the answer-key workbook.",
    )
    parser.add_argument(
        "--template",
        default=str(
            _first_existing_path(
                [
                    repo_root / "templates" / "Homework 3 Gradesheet Template.xlsx",
                    repo_root / "templates" / "Homework_3_Gradesheet_Template.xlsx",
                ]
            )
        ),
        help="Path to gradesheet template workbook.",
    )
    parser.add_argument(
        "--output_dir",
        default=str(repo_root / "outputs"),
        help="Output directory for generated grades and logs.",
    )
    return parser.parse_args()


if __name__ == "__main__":
    repo_root = Path(__file__).resolve().parents[1]
    args = _parse_args(repo_root)
    run_all(
        submissions_dir=Path(args.submissions_dir),
        answer_key_path=Path(args.answer_key),
        template_path=Path(args.template),
        output_dir=Path(args.output_dir),
    )
