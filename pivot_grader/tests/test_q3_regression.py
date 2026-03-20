from __future__ import annotations

import json
from pathlib import Path

import pytest

from grader.ingest import load_answer_key, load_student_submission
from grader.questions.q3 import evaluate_q3_structure
from grader.run_grader import _grade_one_question
from grader.sheet_matcher import match_sheets_to_questions

ROOT_DIR = Path(__file__).resolve().parents[1]
ANSWER_KEY_PATH = ROOT_DIR / "answer_key" / "GryzzlSales2024 - Answer Key.xlsx"
SUBSET_DIR = ROOT_DIR / "outputs" / "subset_runs" / "subset_20260319T182330Z" / "submissions_subset"
RUBRIC_PATH = ROOT_DIR / "configs" / "rubric.json"


@pytest.mark.skipif(not ANSWER_KEY_PATH.exists(), reason="Answer key workbook not found")
@pytest.mark.skipif(not SUBSET_DIR.exists(), reason="Subset submissions folder not found")
def test_q3_regression_bautista_and_azimipour() -> None:
    rubric = json.loads(RUBRIC_PATH.read_text(encoding="utf-8"))
    q3_cfg = next(q for q in rubric["questions"] if q.get("id") == "Q3")

    answer_key = load_answer_key(ANSWER_KEY_PATH)

    cases = {
        "bautista": {
            "filename": "bautistapinedatere_1291030_18792748_GryzzlSales2024 TereBautista.xlsb.xlsx",
            "expect_structure_ok": False,
            "expect_value_match": True,
            "expect_score": 0.7,
        },
        "azimipour": {
            "filename": "azimipouralessio_1276652_18809839_Azimipour_GryzzlSales2024.xlsx",
            "expect_structure_ok": True,
            "expect_value_match": True,
            "expect_score": 1.0,
        },
    }

    for name, case in cases.items():
        workbook_path = SUBSET_DIR / case["filename"]
        assert workbook_path.exists(), f"Missing regression workbook for {name}: {workbook_path}"

        submission = load_student_submission(workbook_path)
        assert submission.error is None, f"Failed to load {name}: {submission.error}"

        warnings: list[str] = []
        matched, matched_names = match_sheets_to_questions(
            submission.sheets,
            answer_key,
            ["Q3"],
            warnings,
        )
        assert not warnings, f"Unexpected mapping warnings for {name}: {warnings}"

        student_df = matched.get("Q3")
        assert student_df is not None, f"Q3 sheet did not map for {name}"

        structure = evaluate_q3_structure(student_df)
        structure_ok = bool(structure["structure_ok"])

        score, _ = _grade_one_question(
            "Q3",
            q3_cfg,
            student_df,
            answer_key,
            review_flagged=False,
            workbook_path=submission.workbook_path,
            sheet_name=matched_names.get("Q3"),
        )

        assert structure_ok is case["expect_structure_ok"], (
            f"{name} structure_ok expected {case['expect_structure_ok']} got {structure_ok}"
        )
        if case["expect_value_match"]:
            if case["expect_structure_ok"]:
                assert score == pytest.approx(1.0, abs=1e-9), (
                    f"{name} expected value_match=True to produce full credit when structure_ok=True"
                )
            else:
                assert score == pytest.approx(0.7, abs=1e-9), (
                    f"{name} expected value_match=True with structure deduction only"
                )
        assert score is not None
        assert score == pytest.approx(case["expect_score"], abs=1e-9), (
            f"{name} score expected {case['expect_score']} got {score}"
        )
