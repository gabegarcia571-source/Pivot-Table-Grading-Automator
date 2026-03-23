# Combined Rubric Alignment Prompt

Paste this entire prompt as one message. Do not split it.
Complete every change before running any tests.

---

## Context: what the actual rubric says

```
- Perfect pivot = 1.0 point
- Missing pivot = 0 points
- Wrong sort (when required) = -0.3
- Extra columns = -0.3
- Correct answer present but not highlighted = -0.5 (applies to ALL questions)
- Wrong/incorrect values = at most 0.3 points depending on closeness
- Bad explanation (Q4, Q6, Q7, Q9 only) = -0.3
  Comments: "Needs more detail" / "Should more directly address question" /
            "Answer inconsistent with analysis"
```

This prompt fixes three gaps between the rubric and the current code:
1. Highlight check is missing from all questions except Q8
2. Formatting weight is 0.0 — it should be 0.5
3. NEEDS_REVIEW on explanation wipes the whole score — it should only 
   withhold the explanation portion while recording the pivot score

---

## CHANGE 1 — scoring.py: fix formatting weight

This is the only change to scoring.py.

Find:
  formatting_weight = 0.0

Replace with:
  formatting_weight = abs(float(DEDUCTIONS["answer_not_highlighted"]["points"]))

DEDUCTIONS["answer_not_highlighted"]["points"] is already -0.5.
This makes the 0.5 highlight deduction live everywhere with no 
hardcoded magic number.

---

## CHANGE 2 — pivot_checker.py: add shared highlight utilities

Add these two functions to pivot_checker.py.

First, move _cell_is_highlighted() out of q8.py and into pivot_checker.py.
The exact logic must be preserved — do not rewrite it:

  def _cell_is_highlighted(cell: Any) -> bool:
      fill = cell.fill
      if fill is None or fill.fill_type in (None, "none"):
          return False
      fg = fill.fgColor
      if fg is None:
          return False
      if fg.type in ("theme", "indexed"):
          return True   # treat theme/indexed fills as highlighted
      rgb = (fg.rgb if hasattr(fg, "rgb") else "").upper()
      return rgb not in ("", "00000000", "FFFFFFFF", "FF000000")

NOTE: the current version in q8.py returns False for theme/indexed fills.
Change it to return True — many students apply Excel's built-in theme
colors and those should count as highlights.

Then add:

  def has_any_highlight(workbook_path: Any, sheet_name: str) -> bool:
      """Return True if any cell on the sheet has a non-default fill."""
      import openpyxl
      from pathlib import Path as _Path
      try:
          wb = openpyxl.load_workbook(_Path(workbook_path), data_only=True)
      except Exception:
          return False
      if sheet_name not in wb.sheetnames:
          wb.close()
          return False
      ws = wb[sheet_name]
      for row in ws.iter_rows():
          for cell in row:
              if _cell_is_highlighted(cell):
                  wb.close()
                  return True
      wb.close()
      return False

---

## CHANGE 3 — q8.py: simplify to use shared utilities

q8.py currently has its own _cell_is_highlighted() and a complex
customer-ID matching check. Replace the entire module with this:

  - Remove the local _cell_is_highlighted() definition (it now lives 
    in pivot_checker.py)
  - Remove check_q8_highlight() entirely
  - Remove all references to HOLIDAY_ONLY_CUSTOMERS and 
    HOLIDAY_ONLY_CUSTOMERS_COMPLETE
  - grade_question() becomes:

    def grade_question(
        student_df: pd.DataFrame,
        answer_df: pd.DataFrame,
        question_cfg: dict[str, Any],
        workbook_path: Any = None,
        sheet_name: str | None = None,
        qid: str = "Q8",
    ) -> dict[str, Any]:
        from grader.pivot_checker import has_any_highlight

        structural_issues: list[str] = []
        value_issues: list[str] = []
        formatting_issues: list[str] = []

        if student_df.empty:
            return {
                "structural_score": 0.0,
                "value_score": 0.0,
                "formatting_score": 0.0,
                "explanation_score": 1.0,
                "structural_issues": ["Missing pivot table"],
                "value_issues": ["Missing pivot table"],
                "formatting_issues": ["Missing highlight"],
                "explanation_issues": [],
            }

        # Structural: pivot must have customer-level granularity
        # (more than 1000 rows) and between 3 and 14 columns
        structural_score = 1.0
        if len(student_df) < 1000 or not (3 <= len(student_df.columns) <= 14):
            structural_score = 0.0
            structural_issues.append("Incorrect filter")

        # Value: Q8 has no programmatic value check — the pivot structure
        # itself is the answer. Full value credit if structurally valid.
        value_score = structural_score

        # Formatting: any highlight on the sheet = full credit
        formatting_score = 1.0
        if workbook_path and sheet_name:
            if not has_any_highlight(workbook_path, sheet_name):
                formatting_score = 0.0
                formatting_issues.append("Missing highlight")

        return {
            "structural_score": structural_score,
            "value_score": value_score,
            "formatting_score": formatting_score,
            "explanation_score": 1.0,
            "structural_issues": structural_issues,
            "value_issues": value_issues,
            "formatting_issues": formatting_issues,
            "explanation_issues": [],
        }

---

## CHANGE 4 — all other question modules: add highlight check

Update the signature and body of grade_question() in:
  q1.py, q2.py, q3.py, q4.py, q5.py, q6.py, q7.py, q9.py, q10.py

Add to every signature:
  workbook_path: Any = None,
  sheet_name: str | None = None,

Add to every grade_question() body, after all structural/value logic,
just before the return statement:

  from grader.pivot_checker import has_any_highlight
  formatting_score = 1.0
  formatting_issues: list[str] = []
  if workbook_path and sheet_name:
      if not has_any_highlight(workbook_path, sheet_name):
          formatting_score = 0.0
          formatting_issues.append("Missing highlight")

Then include formatting_score and formatting_issues in the returned dict.
Replace any existing hardcoded "formatting_score": 1.0 and 
"formatting_issues": [] with the variables.

---

## CHANGE 5 — run_grader.py: pass workbook_path and sheet_name to all questions

Currently workbook_path and sheet_name are only passed to q8.
Find the call site where grade_question() is called for each question
and ensure workbook_path and sheet_name are passed for every question,
not just Q8.

The call will look something like:
  contract = grade_fn(
      student_df=student_df,
      answer_df=answer_df,
      question_cfg=question_cfg,
      qid=qid,
      workbook_path=submission.workbook_path,
      sheet_name=matched_names.get(qid),
  )

Confirm this is the case for all 10 questions. If not, fix it.

---

## CHANGE 6 — run_grader.py: explanation NEEDS_REVIEW returns partial score

The rubric treats the pivot and the explanation as independent deductions.
A student with a correct pivot and a flagged explanation should receive a
numeric score for the pivot portion while the explanation is held for review.

Find where NEEDS_REVIEW causes the question score to become None.

Change the logic so:
  - NEEDS_REVIEW in explanation_issues only:
      Compute partial score WITHOUT explanation deduction:
        partial = 1.0
        partial -= 0.3 * (1 - structural_score)
        partial -= 0.7 * (1 - value_score)
        partial -= 0.5 * (1 - formatting_score)
        partial = max(0.0, round(partial, 1))
      Return (partial, ["NEEDS_REVIEW: explanation pending manual grade"])

  - NEEDS_REVIEW in structural_issues or value_issues:
      Keep returning (None, [...]) — those need full manual review.

formatting_score is now part of the partial score formula because the
0.5 weight is live after Change 1.

---

## Locked files — do not touch

  - configs/rubric.json
  - grader/ingest.py
  - grader/grade_writer.py
  - grader/qualitative_grader.py

---

## Validation after all changes

1. Run pytest — must pass (minimum 40 tests, no regressions).

2. Re-run grader on the real student file:
   asamoahemmanuel_1285342_18790702_EAsamoah-_GryzzlSales2024_WM.xlsx

3. Print Q1–Q10 scores AND comments from the grade workbook.

4. This student highlighted on every sheet. Expected:
   - formatting_score = 1.0 for ALL questions
   - If any question shows "Missing highlight" for this student, 
     has_any_highlight() has a bug — investigate before marking done.

5. Expected scores for this student:
   Q1:  0.7  | Incorrect sort
   Q2:  1.0  |
   Q3:  1.0  |
   Q4:  partial numeric | NEEDS_REVIEW: explanation pending manual grade
   Q5:  1.0  |
   Q6:  partial numeric | NEEDS_REVIEW: explanation pending manual grade
   Q7:  partial numeric | NEEDS_REVIEW: explanation pending manual grade
   Q8:  1.0  (pivot is valid, highlighting present)
   Q9:  partial numeric | NEEDS_REVIEW: explanation pending manual grade
   Q10: investigate — prior run showed 0.7 which was incorrect.
        Print the full student Q10 df and answer Q10 df so we can 
        diagnose why the value check matched when it should not have.

6. If any score does not match expected, do NOT move on. 
   Open the answer key Excel and the student Excel for that question,
   compare shapes and value types, and fix the comparison logic.
   Report findings before making any fix.
