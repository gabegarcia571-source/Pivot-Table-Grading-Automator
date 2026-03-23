from __future__ import annotations

import importlib
from types import SimpleNamespace
from typing import Any

import pandas as pd

from grader.run_grader import _grade_one_question
from grader.scoring import assemble_score


class _FakeModule:
    def __init__(self, contract: dict[str, Any]) -> None:
        self._contract = contract

    def grade_question(
        self,
        student_df: pd.DataFrame,
        answer_df: pd.DataFrame,
        question_cfg: dict[str, Any],
        workbook_path: Any = None,
        sheet_name: str | None = None,
        qid: str = "QX",
    ) -> dict[str, Any]:
        return self._contract


def test_assemble_score_applies_formatting_deduction() -> None:
    contract = {
        "structural_score": 1.0,
        "value_score": 1.0,
        "formatting_score": 0.0,
        "explanation_score": 1.0,
        "structural_issues": [],
        "value_issues": [],
        "formatting_issues": ["Missing highlight"],
        "explanation_issues": [],
    }

    score, comments = assemble_score(contract)

    assert score == 0.5
    assert "Missing highlight" in comments


def test_grade_one_question_explanation_needs_review_returns_partial(monkeypatch: Any) -> None:
    contract = {
        "structural_score": 1.0,
        "value_score": 1.0,
        "formatting_score": 1.0,
        "explanation_score": 0.0,
        "structural_issues": [],
        "value_issues": [],
        "formatting_issues": [],
        "explanation_issues": ["NEEDS_REVIEW: explanation grading unavailable (LLM/API)"],
    }

    def _fake_import(_name: str) -> Any:
        return _FakeModule(contract)

    monkeypatch.setattr(importlib, "import_module", _fake_import)

    student_df = pd.DataFrame({"Row Labels": ["A"], "Value": [1.0]})
    answer_key = {"Q4": pd.DataFrame({"Row Labels": ["A"], "Value": [1.0]})}

    score, comment = _grade_one_question(
        qid="Q4",
        question_cfg={},
        student_df=student_df,
        answer_key=answer_key,
        review_flagged=False,
    )

    assert score == 1.0
    assert comment == "NEEDS_REVIEW: explanation pending manual grade"


def test_grade_one_question_value_needs_review_returns_none(monkeypatch: Any) -> None:
    contract = {
        "structural_score": 1.0,
        "value_score": 1.0,
        "formatting_score": 1.0,
        "explanation_score": 1.0,
        "structural_issues": [],
        "value_issues": ["NEEDS_REVIEW: ambiguous numeric conversion"],
        "formatting_issues": [],
        "explanation_issues": [],
    }

    def _fake_import(_name: str) -> Any:
        return _FakeModule(contract)

    monkeypatch.setattr(importlib, "import_module", _fake_import)

    student_df = pd.DataFrame({"Row Labels": ["A"], "Value": [1.0]})
    answer_key = {"Q10": pd.DataFrame({"Row Labels": ["A"], "Value": [1.0]})}

    score, comment = _grade_one_question(
        qid="Q10",
        question_cfg={},
        student_df=student_df,
        answer_key=answer_key,
        review_flagged=False,
    )

    assert score is None
    assert comment.startswith("NEEDS_REVIEW")
