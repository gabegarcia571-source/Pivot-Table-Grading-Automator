from __future__ import annotations

import json
import os
import re
from typing import Any

try:
    from anthropic import Anthropic
except Exception:  # noqa: BLE001
    Anthropic = None  # type: ignore[assignment]


EXPLANATION_PROMPT = """
You are a strict rubric-based grader for a business analytics course.

Question ID: {question_id}
Grading rubric: {rubric_text}

Student explanation:
"{student_text}"

Return ONLY valid JSON, no other text:
{{
  "deduct_explanation": <true if explanation is wrong, vague, or missing key logic>,
  "confidence": <float 0.0-1.0>,
    "brief_reason": "<one short phrase>"
}}

Use concise rubric-style comments only. Preferred phrases:
- "Needs more detail"
- "Should more directly address question"
- "Answer inconsistent with analysis"
""".strip()


def _sanitize_student_text(student_text: str) -> str:
    text = student_text or ""
    text = re.sub(r"(?im)^\s*(name|student\s*name)\s*:\s*.*$", "", text)
    text = re.sub(r"(?im)^\s*(id|student\s*id)\s*:\s*.*$", "", text)
    text = re.sub(r"\b\d{7,10}\b", "[REDACTED_ID]", text)
    return text.strip()


def _fallback_bad_explanation(reason: str) -> dict[str, Any]:
    return {
        "deduct_explanation": True,
        "confidence": 0.0,
        "brief_reason": _short_explanation_comment(reason),
        "needs_review": False,
    }


def _fallback_needs_review(reason: str) -> dict[str, Any]:
    return {
        "deduct_explanation": False,
        "confidence": 0.0,
        "brief_reason": f"NEEDS_REVIEW: {reason}",
        "needs_review": True,
    }


def _short_explanation_comment(reason: str) -> str:
    low = (reason or "").strip().lower()
    if "inconsistent" in low and "analysis" in low:
        return "Answer inconsistent with analysis"
    if "direct" in low and "question" in low:
        return "Should more directly address question"
    if "off-topic" in low:
        return "Should more directly address question"
    return "Needs more detail"


def grade_explanation(question_id: str, student_text: str, rubric_text: str) -> dict[str, Any]:
    """Grade explanation via Anthropic with deterministic settings."""
    sanitized = _sanitize_student_text(student_text)
    if not sanitized:
        return _fallback_bad_explanation("No explanation text provided.")

    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key or Anthropic is None:
        return _fallback_needs_review("explanation grading unavailable (LLM/API)")

    prompt = EXPLANATION_PROMPT.format(
        question_id=question_id,
        rubric_text=rubric_text,
        student_text=sanitized,
    )

    client = Anthropic(api_key=api_key)
    model = os.getenv("ANTHROPIC_MODEL", "claude-3-5-haiku-latest")

    try:
        response = client.messages.create(
            model=model,
            temperature=0,
            max_tokens=250,
            messages=[{"role": "user", "content": prompt}],
        )

        raw_text = ""
        for block in response.content:
            if getattr(block, "type", "") == "text":
                raw_text += block.text

        payload = json.loads(raw_text)
        return {
            "deduct_explanation": bool(payload.get("deduct_explanation", True)),
            "confidence": float(payload.get("confidence", 0.0)),
            "brief_reason": _short_explanation_comment(
                str(payload.get("brief_reason", "Needs more detail"))
            ),
            "needs_review": False,
        }
    except Exception as exc:  # noqa: BLE001
        return _fallback_needs_review(f"explanation grading failed ({exc})")
