from __future__ import annotations

import unicodedata


def normalize_label(value: str) -> str:
    """Normalize labels for robust matching across punctuation variants."""
    normalized = unicodedata.normalize("NFKD", value).lower()
    normalized = normalized.replace("’", "'").replace("‘", "'").replace("`", "'")
    cleaned = "".join(
        ch if (ch.isalnum() or ch.isspace() or ch == "'") else " "
        for ch in normalized
    )
    return " ".join(cleaned.split())