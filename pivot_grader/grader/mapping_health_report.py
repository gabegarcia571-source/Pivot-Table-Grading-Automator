from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from grader.ingest import load_answer_key, load_student_submission
from grader.run_grader import _match_sheets_to_questions


QUESTION_IDS = [f"Q{i}" for i in range(1, 11)]


@dataclass
class MappingResult:
    file_name: str
    mapped_count: int
    status: str  # FULL | PARTIAL | FAILED
    sheet_names: list[str]
    missing_qs: list[str]
    error: str | None = None


def _find_answer_key(answer_key_dir: Path) -> Path:
    xlsx_files = sorted([p for p in answer_key_dir.glob("*.xlsx") if not p.name.startswith("~$")])
    if not xlsx_files:
        raise FileNotFoundError(f"No .xlsx files found in answer_key directory: {answer_key_dir}")

    # Prefer filenames containing "answer key", else fallback to first workbook.
    for p in xlsx_files:
        if "answer" in p.name.lower() and "key" in p.name.lower():
            return p
    return xlsx_files[0]


def _classify(mapped_count: int, unreadable: bool) -> str:
    if unreadable or mapped_count == 0:
        return "FAILED"
    if mapped_count == 10:
        return "FULL"
    return "PARTIAL"


def build_mapping_health_report(
    submissions_dir: Path,
    answer_key_path: Path,
) -> tuple[dict[str, Any], list[MappingResult]]:
    answer_key = load_answer_key(answer_key_path)

    workbooks = sorted(
        [
            p
            for p in submissions_dir.rglob("*.xlsx")
            if not p.name.startswith("~$") and "grade" not in p.name.lower()
        ]
    )

    results: list[MappingResult] = []

    for workbook in workbooks:
        submission = load_student_submission(workbook)

        if submission.error:
            results.append(
                MappingResult(
                    file_name=workbook.name,
                    mapped_count=0,
                    status="FAILED",
                    sheet_names=[],
                    missing_qs=QUESTION_IDS.copy(),
                    error=submission.error,
                )
            )
            continue

        warnings: list[str] = []
        _, matched_names = _match_sheets_to_questions(
            submission.sheets,
            answer_key,
            QUESTION_IDS,
            warnings,
        )
        mapped_count = sum(1 for qid in QUESTION_IDS if matched_names.get(qid) is not None)
        missing_qs = [qid for qid in QUESTION_IDS if matched_names.get(qid) is None]
        status = _classify(mapped_count, unreadable=False)

        results.append(
            MappingResult(
                file_name=workbook.name,
                mapped_count=mapped_count,
                status=status,
                sheet_names=sorted(submission.sheets.keys()),
                missing_qs=missing_qs,
            )
        )

    summary = {
        "total_students": len(results),
        "full_count": sum(1 for r in results if r.status == "FULL"),
        "partial_count": sum(1 for r in results if r.status == "PARTIAL"),
        "failed_count": sum(1 for r in results if r.status == "FAILED"),
    }
    return summary, results


def print_mapping_health_report(summary: dict[str, Any], results: list[MappingResult]) -> None:
    print("MAPPING HEALTH REPORT")
    print("=" * 80)
    print(f"Total students: {summary['total_students']}")
    print(f"FULL (10/10):   {summary['full_count']}")
    print(f"PARTIAL (1-9):  {summary['partial_count']}")
    print(f"FAILED (0/10):  {summary['failed_count']}")

    problem_cases = [r for r in results if r.status in {"PARTIAL", "FAILED"}]
    problem_cases.sort(key=lambda r: (r.mapped_count, r.file_name.lower()))

    print("\nPROBLEM CASES")
    print("=" * 80)
    if not problem_cases:
        print("None")
        return

    for row in problem_cases:
        print(f"File: {row.file_name}")
        print(f"Status: {row.status} ({row.mapped_count}/10)")
        print(f"Sheet names: {', '.join(row.sheet_names) if row.sheet_names else '(none)'}")
        print(f"Missing Qs: {', '.join(row.missing_qs) if row.missing_qs else '(none)'}")
        if row.error:
            print(f"Error: {row.error}")
        print("-" * 80)


def main() -> None:
    repo_root = Path(__file__).resolve().parents[1]
    submissions_dir = repo_root / "submissions"
    answer_key_dir = repo_root / "answer_key"
    answer_key_path = _find_answer_key(answer_key_dir)

    summary, results = build_mapping_health_report(
        submissions_dir=submissions_dir,
        answer_key_path=answer_key_path,
    )
    print_mapping_health_report(summary, results)


if __name__ == "__main__":
    main()
