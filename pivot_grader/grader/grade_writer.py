from __future__ import annotations

import shutil
from pathlib import Path

from openpyxl import load_workbook


def write_grades(
    student_id: str,
    scores_dict: dict[str, float],
    comments_dict: dict[str, str],
    template_path: str | Path,
    output_dir: str | Path,
) -> Path:
    """Write student grading results into a template copy."""
    template = Path(template_path)
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    output_path = out_dir / f"{student_id}_grade.xlsx"
    shutil.copy2(template, output_path)

    wb = load_workbook(output_path)
    ws = wb["Gradesheet"]

    ws["C3"] = student_id

    for i in range(1, 11):
        qid = f"Q{i}"
        row = 6 + i  # Q1->7 ... Q10->16
        ws[f"C{row}"] = float(scores_dict.get(qid, 0.0))
        ws[f"D{row}"] = comments_dict.get(qid, "")

    wb.save(output_path)
    return output_path
