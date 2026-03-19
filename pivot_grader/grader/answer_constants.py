"""
Hardcoded correct-answer sets used for highlight-based grading checks.

TODO: Populate HOLIDAY_ONLY_CUSTOMERS with the full 2,487 customer IDs from
      the answer key (GryzzlSales2024_Answer_Key.xlsx).  The partial set below
      contains only the first few known IDs — replace the entire set once the
      answer key is available.
"""

from __future__ import annotations

# ---------------------------------------------------------------------------
# Q8 — customers who order ONLY in November and December
# ---------------------------------------------------------------------------
# These 2,487 customer IDs represent every customer whose purchase history
# contains transactions exclusively in November and/or December 2024.
# Source: GryzzlSales2024_Answer_Key.xlsx, Q8 sheet (highlighted rows).
#
# TODO: Replace with the full set extracted from the answer key.
HOLIDAY_ONLY_CUSTOMERS: frozenset[int] = frozenset({
    20, 37, 122, 124, 151, 171, 255, 331, 378, 383,
    # ... paste the remaining ~2,477 IDs from the answer key here ...
})

# Safety flag for graders: False means this answer set is known incomplete and
# must not be used for automatic grading.
HOLIDAY_ONLY_CUSTOMERS_COMPLETE = False
