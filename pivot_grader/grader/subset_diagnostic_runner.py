from __future__ import annotations

import argparse
import json
import shutil
import zipfile
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path

from grader.ingest import load_student_submission
from grader.run_grader import run_all


@dataclass
class IngestionDiagnostic:
    filename: str
    sheets_detected: int
    pivot_tables_detected: int | None
    ingestion_succeeded: bool
    error: str | None


def _discover_submission_files(submissions_dir: Path) -> list[Path]:
    return sorted(
        [
            p
            for p in submissions_dir.rglob("*.xlsx")
            if p.is_file() and not p.name.startswith("~$") and "grade" not in p.name.lower()
        ]
    )


def _count_pivot_tables(workbook_path: Path | None) -> int | None:
    """Best-effort pivot count from OOXML parts; returns None when unavailable."""
    if workbook_path is None or not workbook_path.exists() or workbook_path.suffix.lower() != ".xlsx":
        return None

    try:
        with zipfile.ZipFile(workbook_path) as zf:
            return sum(
                1
                for info in zf.infolist()
                if info.filename.startswith("xl/pivotTables/") and info.filename.endswith(".xml")
            )
    except Exception:  # noqa: BLE001
        return None


def _collect_ingestion_diagnostics(files: list[Path]) -> list[IngestionDiagnostic]:
    diagnostics: list[IngestionDiagnostic] = []
    for file_path in files:
        result = load_student_submission(file_path)
        diagnostics.append(
            IngestionDiagnostic(
                filename=file_path.name,
                sheets_detected=len(result.sheets),
                pivot_tables_detected=_count_pivot_tables(result.workbook_path or file_path),
                ingestion_succeeded=result.error is None,
                error=result.error,
            )
        )
    return diagnostics


def _stage_subset_files(files: list[Path], subset_dir: Path) -> None:
    subset_dir.mkdir(parents=True, exist_ok=True)
    for file_path in files:
        shutil.copy2(file_path, subset_dir / file_path.name)


def _resolve_default_workbook(directory: Path, preferred_names: list[str]) -> Path:
    for name in preferred_names:
        candidate = directory / name
        if candidate.exists():
            return candidate

    all_xlsx = sorted([p for p in directory.glob("*.xlsx") if p.is_file()])
    if not all_xlsx:
        raise FileNotFoundError(f"No .xlsx files found in {directory}")
    return all_xlsx[0]


def run_subset_diagnostic(
    submissions_dir: Path,
    answer_key_path: Path,
    template_path: Path,
    output_dir: Path,
    limit: int,
) -> dict:
    all_files = _discover_submission_files(submissions_dir)
    selected_files = all_files[:limit]

    if not selected_files:
        raise RuntimeError(f"No submission Excel files found in {submissions_dir}")

    diagnostics = _collect_ingestion_diagnostics(selected_files)

    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    run_dir = output_dir / "subset_runs" / f"subset_{timestamp}"
    subset_inputs_dir = run_dir / "submissions_subset"
    subset_outputs_dir = run_dir / "outputs"
    run_dir.mkdir(parents=True, exist_ok=True)

    _stage_subset_files(selected_files, subset_inputs_dir)

    grading_summary = run_all(
        submissions_dir=subset_inputs_dir,
        answer_key_path=answer_key_path,
        template_path=template_path,
        output_dir=subset_outputs_dir,
    )

    failed = [d for d in diagnostics if not d.ingestion_succeeded]
    successful = [d for d in diagnostics if d.ingestion_succeeded]

    result = {
        "timestamp_utc": datetime.now(timezone.utc).isoformat(),
        "files_tested": [p.name for p in selected_files],
        "successful_ingestion": len(successful),
        "failed_ingestion": len(failed),
        "diagnostics": [asdict(d) for d in diagnostics],
        "subset_run_dir": str(run_dir),
        "grading_summary": grading_summary,
    }

    diagnostic_path = run_dir / "subset_diagnostic_summary.json"
    diagnostic_path.write_text(json.dumps(result, indent=2), encoding="utf-8")

    return result


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run grading pipeline on first N submissions with ingestion diagnostics.",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=5,
        help="Number of submissions to test (recommended 3-5).",
    )
    parser.add_argument("--submissions-dir", type=Path, default=None)
    parser.add_argument("--answer-key", type=Path, default=None)
    parser.add_argument("--template", type=Path, default=None)
    parser.add_argument("--output-dir", type=Path, default=None)
    return parser.parse_args()


def main() -> None:
    args = _parse_args()

    if args.limit < 1:
        raise ValueError("--limit must be >= 1")

    repo_root = Path(__file__).resolve().parents[1]
    submissions_dir = args.submissions_dir or (repo_root / "submissions")
    answer_key = args.answer_key or _resolve_default_workbook(
        repo_root / "answer_key",
        ["GryzzlSales2024_Answer_Key.xlsx", "GryzzlSales2024 - Answer Key.xlsx"],
    )
    template = args.template or _resolve_default_workbook(
        repo_root / "templates",
        ["Homework_3_Gradesheet_Template.xlsx", "Homework 3 Gradesheet Template.xlsx"],
    )
    output_dir = args.output_dir or (repo_root / "outputs")

    result = run_subset_diagnostic(
        submissions_dir=submissions_dir,
        answer_key_path=answer_key,
        template_path=template,
        output_dir=output_dir,
        limit=args.limit,
    )

    print("Subset diagnostic run complete")
    print(f"Files tested ({len(result['files_tested'])}):")
    for name in result["files_tested"]:
        print(f"- {name}")
    print(f"Successful ingestion: {result['successful_ingestion']}")
    print(f"Failed ingestion: {result['failed_ingestion']}")
    print(f"Summary: {Path(result['subset_run_dir']) / 'subset_diagnostic_summary.json'}")


if __name__ == "__main__":
    main()