from __future__ import annotations

from typing import Any

import pandas as pd


def _coerce_label(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _first_numeric_column(df: pd.DataFrame) -> str | None:
    for col in df.columns[1:]:
        converted = pd.to_numeric(df[col], errors="coerce")
        if converted.notna().any():
            return str(col)
    return None


def _normalize_for_compare(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or len(df.columns) < 2:
        return pd.DataFrame(columns=["label", "value"])

    first_col = df.columns[0]
    value_col = _first_numeric_column(df)
    if value_col is None:
        return pd.DataFrame(columns=["label", "value"])

    out = df[[first_col, value_col]].copy()
    out.columns = ["label", "value"]
    out["label"] = out["label"].apply(_coerce_label)
    out["value"] = pd.to_numeric(out["value"], errors="coerce")
    out = out.dropna(subset=["value"])  # keep only comparable rows
    out = out[out["label"] != ""]
    return out


def compare_pivot_values(
    student_df: pd.DataFrame,
    answer_df: pd.DataFrame,
    numeric_tolerance: float = 1e-2,
) -> dict[str, Any]:
    """Compare student and answer pivots by label + value with tolerance."""
    student_norm = _normalize_for_compare(student_df)
    answer_norm = _normalize_for_compare(answer_df)

    if student_norm.empty or answer_norm.empty:
        return {
            "match": False,
            "mismatches": [
                {
                    "label": "<table>",
                    "expected": "non-empty comparable pivot",
                    "actual": "missing or non-numeric values",
                }
            ],
            "score_suggestion": 0.0,
        }

    answer_map = answer_norm.set_index("label")["value"].to_dict()
    student_map = student_norm.set_index("label")["value"].to_dict()

    mismatches: list[dict[str, Any]] = []
    matches = 0

    for label, expected in answer_map.items():
        actual = student_map.get(label)
        if actual is None:
            mismatches.append({"label": label, "expected": expected, "actual": None})
            continue

        if abs(float(actual) - float(expected)) <= numeric_tolerance:
            matches += 1
        else:
            mismatches.append({"label": label, "expected": float(expected), "actual": float(actual)})

    total = max(1, len(answer_map))
    ratio = matches / total
    match = len(mismatches) == 0
    score_suggestion = round(max(0.0, min(1.0, ratio)), 2)

    return {
        "match": match,
        "mismatches": mismatches,
        "score_suggestion": score_suggestion,
    }


def is_desc_sorted(df: pd.DataFrame) -> bool:
    normalized = _normalize_for_compare(df)
    if normalized.empty:
        return False
    values = normalized["value"].tolist()
    return values == sorted(values, reverse=True)


def is_desc_sorted_within_groups(df: pd.DataFrame) -> bool:
    """For nested pivots (col0 = outer group, repeated via ffill; col1+ = inner rows + values):
    verify that within each col0 group the rows are descending by the first numeric value column.

    Falls back to ``is_desc_sorted`` when all col0 values are unique (flat pivot).
    Rows whose outer label contains "total" are treated as subtotals and excluded.
    """
    if df.empty or len(df.columns) < 2:
        return is_desc_sorted(df)

    outer = df.iloc[:, 0].astype(str).str.strip()
    # Flat pivot: every category appears exactly once
    if outer.nunique() == len(outer):
        return is_desc_sorted(df)

    # Find the first numeric column (skip col0)
    value_col: str | None = None
    for col in df.columns[1:]:
        if pd.to_numeric(df[col], errors="coerce").notna().any():
            value_col = col
            break
    if value_col is None:
        return False

    tmp = df[[df.columns[0], value_col]].copy()
    tmp.columns = ["__group__", "__val__"]
    tmp["__group__"] = tmp["__group__"].astype(str).str.strip().str.lower()
    tmp["__val__"] = pd.to_numeric(tmp["__val__"], errors="coerce")
    tmp = tmp.dropna(subset=["__val__"])
    # Exclude subtotal / grand-total rows
    tmp = tmp[~tmp["__group__"].str.contains("total", regex=False)]

    if tmp.empty:
        return False

    for _, grp in tmp.groupby("__group__", sort=False):
        vals = grp["__val__"].tolist()
        if vals != sorted(vals, reverse=True):
            return False
    return True


def has_multiple_items_marker(df: pd.DataFrame) -> bool:
    if df.empty:
        return False
    for col in df.columns:
        series = df[col].astype(str).str.lower()
        if series.str.contains("multiple items", regex=False).any():
            return True
    return False


def check_column_count(df: pd.DataFrame, expected_columns: int) -> bool:
    return df.shape[1] == expected_columns


# ---------------------------------------------------------------------------
# Sheet fingerprinting for content-based sheet-to-question matching
# ---------------------------------------------------------------------------

def sheet_fingerprint(df: pd.DataFrame) -> dict[str, Any]:
    """Lightweight fingerprint for matching a student sheet to an answer-key question."""
    n_rows, n_cols = df.shape

    if n_rows < 20:
        row_bucket = "tiny"
    elif n_rows < 200:
        row_bucket = "small"
    elif n_rows < 2_000:
        row_bucket = "medium"
    elif n_rows < 15_000:
        row_bucket = "large"
    else:
        row_bucket = "xlarge"

    first_col = df.iloc[:, 0].dropna() if not df.empty else pd.Series([], dtype=object)
    if len(first_col) > 0:
        numeric_ratio = float(pd.to_numeric(first_col, errors="coerce").notna().mean())
    else:
        numeric_ratio = 0.0
    first_col_numeric = numeric_ratio > 0.8

    # Collect label set only for small tables (large ones are too expensive)
    first_col_labels: frozenset[str] = frozenset()
    if n_rows < 200:
        first_col_labels = frozenset(
            str(v).strip().lower()
            for v in first_col
            if v is not None and not pd.isna(v)
        )

    return {
        "n_rows": n_rows,
        "n_cols": n_cols,
        "row_bucket": row_bucket,
        "first_col_numeric": first_col_numeric,
        "first_col_labels": first_col_labels,
    }


def fingerprint_similarity(student_fp: dict[str, Any], answer_fp: dict[str, Any]) -> float:
    """Score how closely a student sheet fingerprint matches an answer-key fingerprint."""
    score = 0.0

    if student_fp["row_bucket"] == answer_fp["row_bucket"]:
        score += 3.0

    s_cols, a_cols = student_fp["n_cols"], answer_fp["n_cols"]
    if s_cols == a_cols:
        score += 2.0
    elif abs(s_cols - a_cols) <= 1:
        score += 1.0

    if student_fp["first_col_numeric"] == answer_fp["first_col_numeric"]:
        score += 2.0

    s_labels = student_fp["first_col_labels"]
    a_labels = answer_fp["first_col_labels"]
    if s_labels and a_labels:
        overlap = len(s_labels & a_labels) / max(len(a_labels), 1)
        score += overlap * 4.0

    return score


# ---------------------------------------------------------------------------
# Filter verification — Q3, Q5, Q10
# ---------------------------------------------------------------------------

_KNOWN_VENDORS: frozenset[str] = frozenset({
    "sweetums industries",
    "jj's diner goods",
    "rent-a-swag inc.",
    "lil' sebastian co.",
})

# The top 3 vendors from Q2 that should appear as the only row labels in Q3.
_TOP3_VENDORS: frozenset[str] = frozenset({
    "sweetums industries",
    "jj's diner goods",
    "lil' sebastian co.",
})

_ALL_MONTH_NAMES: frozenset[str] = frozenset({
    "jan", "feb", "mar", "apr", "may", "jun",
    "jul", "aug", "sep", "oct", "nov", "dec",
    "january", "february", "march", "april", "june", "july", "august",
    "september", "october", "november", "december",
})

_SUMMER_MONTHS: frozenset[str] = frozenset({
    "jun", "jul", "aug", "june", "july", "august",
})

_NON_SUMMER_MONTHS: frozenset[str] = _ALL_MONTH_NAMES - _SUMMER_MONTHS


def _sheet_labels(df: pd.DataFrame) -> frozenset[str]:
    """All string labels from the first column and column headers, lowercased."""
    labels: set[str] = set()
    if not df.empty:
        for val in df.iloc[:, 0].dropna().astype(str):
            labels.add(val.strip().lower())
    for col in df.columns:
        labels.add(str(col).strip().lower())
    return frozenset(labels)


def check_q3_filter(df: pd.DataFrame) -> dict[str, Any]:
    """Q3 filter check: validate that exactly the top-3 vendors are row labels.

    openpyxl cannot read the Excel UI 'Multiple Items' filter label, so we
    infer correctness from the row-label set:

    - If NO vendor names appear as row labels at all the vendor field is being
      used as a slicer/filter rather than a row field — pass through to value
      comparison.
    - If vendor names DO appear as row labels, exactly these three must be
      present and no others:
        Sweetums Industries, JJ's Diner Goods, Lil' Sebastian Co.
      Any other vendor (e.g. Rent-A-Swag Inc.) or a missing top-3 vendor
      is treated as a wrong filter.
    """
    if df.empty:
        return {
            "filter_ok": False,
            "found_vendors_in_rows": [],
            "missing_vendors": sorted(_TOP3_VENDORS),
            "extra_vendors": [],
        }

    first_col_vals = df.iloc[:, 0].dropna().astype(str).str.strip().str.lower()
    present_top3 = {v for v in _TOP3_VENDORS if first_col_vals.str.contains(v, regex=False).any()}
    extra_vendors = {
        v for v in (_KNOWN_VENDORS - _TOP3_VENDORS)
        if first_col_vals.str.contains(v, regex=False).any()
    }

    if len(present_top3) == 0 and len(extra_vendors) == 0:
        # Vendor is used as a filter field (not a row field) — no vendor labels present.
        filter_ok = True
    else:
        # Vendor IS in the row field: require exactly the top-3, nothing more.
        filter_ok = len(present_top3) == 3 and len(extra_vendors) == 0

    return {
        "filter_ok": filter_ok,
        "found_vendors_in_rows": sorted(present_top3 | extra_vendors),
        "missing_vendors": sorted(_TOP3_VENDORS - present_top3),
        "extra_vendors": sorted(extra_vendors),
    }


def check_q5_filter(df: pd.DataFrame) -> dict[str, Any]:
    """Q5 filter check: only Jun/Jul/Aug should appear as row or column labels."""
    labels = _sheet_labels(df)
    non_summer_found = sorted(_NON_SUMMER_MONTHS & labels)
    return {
        "filter_ok": len(non_summer_found) == 0,
        "non_summer_months_found": non_summer_found,
    }


# Q10 expects all 7 vendors in the dataset as row labels.
# NOTE: Only 4 vendor names are currently known; extend this set once the full
# dataset reveals the remaining 3 vendors.
_Q10_EXPECTED_VENDOR_COUNT = 7
_Q10_VALUE_COLS = frozenset({"honey", "no promo code"})


def check_q10_filter(df: pd.DataFrame) -> dict[str, Any]:
    """Q10 validation: shape, vendor row labels, and column structure.

    openpyxl cannot read the Excel UI 'Multiple Items' label, so month filter
    validation is intentionally skipped here.
    # TODO: month filter validation requires manual review — verify that the
    # student filtered to the correct non-October/November months.

    Checks performed:
      - Exactly 7 vendor row labels (the full dataset has 7 vendors).
      - Column headers include both 'Honey' and 'No Promo Code'.
      - All numeric values are proportions in [0, 1].
    """
    if df.empty:
        return {
            "filter_ok": False,
            "vendor_count": 0,
            "vendor_count_ok": False,
            "has_honey_col": False,
            "has_no_promo_col": False,
            "values_in_range": False,
        }

    # Count distinct non-total vendor row labels.
    first_col_vals = df.iloc[:, 0].dropna().astype(str).str.strip().str.lower()
    data_labels = [v for v in first_col_vals if "total" not in v and v != ""]
    vendor_count = len(set(data_labels))
    vendor_count_ok = vendor_count == _Q10_EXPECTED_VENDOR_COUNT

    # Check column headers for required value columns.
    col_headers_lower = {str(c).strip().lower() for c in df.columns}
    has_honey = any("honey" in h for h in col_headers_lower)
    has_no_promo = any("no promo" in h for h in col_headers_lower)

    # All numeric cell values in data columns should be proportions [0, 1].
    numerics: list[float] = []
    for col in df.columns[1:]:
        converted = pd.to_numeric(df[col], errors="coerce").dropna()
        numerics.extend(float(v) for v in converted)
    values_in_range = bool(numerics) and all(0.0 <= v <= 1.0 for v in numerics)

    column_ok = has_honey and has_no_promo and values_in_range

    return {
        "filter_ok": column_ok and vendor_count_ok,
        "vendor_count": vendor_count,
        "vendor_count_ok": vendor_count_ok,
        "has_honey_col": has_honey,
        "has_no_promo_col": has_no_promo,
        "values_in_range": values_in_range,
    }


# ---------------------------------------------------------------------------
# Q8 highlight-based validation
# ---------------------------------------------------------------------------


def _cell_is_highlighted(cell: Any) -> bool:
    """Return True if a cell has a non-default, non-white background fill."""
    fill = cell.fill
    if fill is None or fill.fill_type in (None, "none"):
        return False
    fg = fill.fgColor
    if fg is None:
        return False
    # Theme and indexed colors are Excel defaults (headers, alternating rows, etc.)
    if fg.type in ("theme", "indexed"):
        return False
    rgb = (fg.rgb if hasattr(fg, "rgb") else "").upper()
    # Transparent (00000000) or white (FFFFFFFF) are not meaningful highlights
    return rgb not in ("", "00000000", "FFFFFFFF", "FF000000")


def check_q8_highlight(
    workbook_path: Any, sheet_name: str
) -> dict[str, Any]:
    """Grade Q8 by detecting highlighted rows in the student's openpyxl sheet.

    Students must highlight the customer IDs who order ONLY in November and
    December.  The correct set is HOLIDAY_ONLY_CUSTOMERS in answer_constants.

    Grading rules:
      - No highlighted rows at all → missing_pivot deduction.
      - Highlighted IDs match the correct set exactly → full credit.
      - Symmetric error ≤ 5 % of the correct set size → full credit (minor error).
      - Error > 5 % → wrong_values deduction.
    """
    from pathlib import Path as _Path
    import openpyxl as _openpyxl
    from grader.answer_constants import HOLIDAY_ONLY_CUSTOMERS, HOLIDAY_ONLY_CUSTOMERS_COMPLETE

    notes: list[str] = []

    if not HOLIDAY_ONLY_CUSTOMERS_COMPLETE:
        return {
            "match": False,
            "missing_pivot": False,
            "needs_review": True,
            "notes": ["NEEDS_REVIEW: Q8 answer set is incomplete"],
        }

    try:
        wb = _openpyxl.load_workbook(_Path(workbook_path), data_only=True)
    except Exception as exc:
        notes.append("Missing highlight")
        return {"match": False, "missing_pivot": False, "needs_review": False, "notes": notes}

    if sheet_name not in wb.sheetnames:
        notes.append("Missing highlight")
        wb.close()
        return {"match": False, "missing_pivot": False, "needs_review": False, "notes": notes}

    ws = wb[sheet_name]
    highlighted_ids: set[int] = set()

    for row in ws.iter_rows():
        if not row:
            continue
        row_highlighted = any(_cell_is_highlighted(cell) for cell in row)
        if not row_highlighted:
            continue
        # Take the first column cell as the row label (customer_id)
        label_val = row[0].value
        try:
            highlighted_ids.add(int(label_val))
        except (ValueError, TypeError):
            pass  # header rows or non-numeric labels

    wb.close()

    if not highlighted_ids:
        notes.append("Missing highlight")
        return {"match": False, "missing_pivot": True, "needs_review": False, "notes": notes}

    correct_set = HOLIDAY_ONLY_CUSTOMERS
    n_correct = max(1, len(correct_set))
    false_positives = highlighted_ids - correct_set
    false_negatives = correct_set - highlighted_ids
    error_count = len(false_positives) + len(false_negatives)
    error_rate = error_count / n_correct

    if error_rate == 0.0:
        return {"match": True, "missing_pivot": False, "needs_review": False, "notes": notes}

    if error_rate <= 0.05:
        return {"match": True, "missing_pivot": False, "needs_review": False, "notes": notes}

    notes.append("Incorrect value")
    return {"match": False, "missing_pivot": False, "needs_review": False, "notes": notes}


# ---------------------------------------------------------------------------
# Q4 — average per-order price scan
# ---------------------------------------------------------------------------

Q4_TARGET = 46.48
Q4_TOLERANCE = 2.0


def check_q4_average(
    df: pd.DataFrame,
    target: float = Q4_TARGET,
    tolerance: float = Q4_TOLERANCE,
) -> dict[str, Any]:
    """Scan every numeric value on the Q4 sheet for the correct per-order average.

    Returns a dict with:
      has_numeric  — True if any numeric value was found on the sheet at all.
      match        — True if at least one value is within *tolerance* of *target*.
      closest      — the numeric value closest to target (or None if none exist).
      delta        — abs(closest - target), None if no numerics exist.
    """
    numerics: list[float] = []

    for col in df.columns:
        converted = pd.to_numeric(df[col], errors="coerce").dropna()
        numerics.extend(float(v) for v in converted)

    if not numerics:
        return {"has_numeric": False, "match": False, "closest": None, "delta": None}

    closest = min(numerics, key=lambda v: abs(v - target))
    delta = abs(closest - target)
    return {
        "has_numeric": True,
        "match": delta <= tolerance,
        "closest": closest,
        "delta": delta,
    }
