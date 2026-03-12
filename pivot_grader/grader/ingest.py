from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl
import pandas as pd


@dataclass
class SubmissionResult:
    student_id: str
    sheets: dict[str, pd.DataFrame]
    error: str | None = None
    workbook_path: Path | None = None


def _normalize_frame(df: pd.DataFrame) -> pd.DataFrame:
    cleaned = df.dropna(how="all").dropna(axis=1, how="all")
    cleaned.columns = [str(col).strip() for col in cleaned.columns]
    return cleaned


def find_pivot_origin(ws) -> tuple[int, int]:
    """
    Scan an openpyxl worksheet for the first non-empty string cell in a row
    that has at least 2 non-empty values — more likely a table header than a
    stray label.  Falls back to the very first non-empty string cell found.
    Returns (row, col) as 1-indexed openpyxl coordinates.
    """
    first_string_pos: tuple[int, int] | None = None

    for row_cells in ws.iter_rows():
        non_empty_in_row = sum(
            1 for cell in row_cells
            if cell.value is not None and str(cell.value).strip()
        )
        for cell in row_cells:
            if cell.value and isinstance(cell.value, str) and cell.value.strip():
                if first_string_pos is None:
                    first_string_pos = (cell.row, cell.column)
                # Prefer a row with ≥2 non-empty cells — much more likely to be a header
                if non_empty_in_row >= 2:
                    return cell.row, cell.column

    return first_string_pos or (1, 1)


def load_answer_key(path: str | Path) -> dict[str, pd.DataFrame]:
    """Load answer key workbook with header on row 3 (skip first 2 blank rows)."""
    workbook_path = Path(path)
    if not workbook_path.exists():
        raise FileNotFoundError(f"Answer key not found: {workbook_path}")

    sheets = pd.read_excel(workbook_path, sheet_name=None, header=2, engine="openpyxl")
    return {name: _normalize_frame(df) for name, df in sheets.items()}


def load_student_submission(folder_path: str | Path) -> SubmissionResult:
    """Load all sheets from the first xlsx file found in the student folder.

    For each sheet, openpyxl is used in read-only/stream mode to locate the
    pivot table origin (the first header-like row).  pandas then reads from
    that row onward, and the DataFrame is trimmed to the pivot's starting
    column so that students who offset their tables still compare correctly.
    """
    folder = Path(folder_path)
    student_id = folder.name

    try:
        xlsx_files = sorted(
            [
                p
                for p in folder.glob("*.xlsx")
                if not p.name.startswith("~$") and "grade" not in p.name.lower()
            ]
        )
        if not xlsx_files:
            return SubmissionResult(student_id=student_id, sheets={}, error="No .xlsx submission found.")

        workbook = xlsx_files[0]

        # Stream-scan once with openpyxl to find each sheet's pivot origin.
        opxl_wb = openpyxl.load_workbook(workbook, read_only=True, data_only=True)
        origins: dict[str, tuple[int, int]] = {
            name: find_pivot_origin(opxl_wb[name]) for name in opxl_wb.sheetnames
        }
        opxl_wb.close()

        sheets: dict[str, pd.DataFrame] = {}
        for sheet_name, (start_row, start_col) in origins.items():
            df = pd.read_excel(
                workbook,
                sheet_name=sheet_name,
                skiprows=start_row - 1,   # skip blank rows above the pivot
                header=0,
                engine="openpyxl",
            )
            if start_col > 1:
                df = df.iloc[:, start_col - 1:]   # trim blank columns to the left
            normalized = _normalize_frame(df.dropna(how="all"))
            # Forward-fill the first column to expand merged pivot row-group labels
            if not normalized.empty:
                normalized.iloc[:, 0] = normalized.iloc[:, 0].ffill()
            sheets[sheet_name] = normalized

        return SubmissionResult(student_id=student_id, sheets=sheets, workbook_path=workbook)
    except Exception as exc:  # noqa: BLE001
        return SubmissionResult(student_id=student_id, sheets={}, error=f"Failed to load submission: {exc}")
