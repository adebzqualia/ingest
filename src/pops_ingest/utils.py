"""Deterministic serialization, hashing, and Excel-coordinate helpers."""

from __future__ import annotations

from datetime import date, datetime, time, timedelta
from decimal import Decimal
from hashlib import sha256
import json
import math
from pathlib import Path
import re
from typing import Any, Iterable


_SLUG_RE = re.compile(r"[^a-z0-9]+")


def sha256_file(path: Path, chunk_size: int = 1024 * 1024) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        while chunk := handle.read(chunk_size):
            digest.update(chunk)
    return digest.hexdigest()


def canonical_json(value: Any) -> str:
    return json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
        default=json_default,
    )


def stable_hash(value: Any) -> str:
    return sha256(canonical_json(value).encode("utf-8")).hexdigest()


def json_default(value: Any) -> Any:
    encoded, _ = encode_typed_value(value)
    return encoded


def encode_typed_value(value: Any) -> tuple[Any, str]:
    if value is None:
        return None, "blank"
    if isinstance(value, bool):
        return value, "boolean"
    if isinstance(value, datetime):
        return value.isoformat(), "datetime"
    if isinstance(value, date):
        return value.isoformat(), "date"
    if isinstance(value, time):
        return value.isoformat(), "time"
    if isinstance(value, timedelta):
        return value.total_seconds(), "duration_seconds"
    if isinstance(value, Decimal):
        return str(value), "decimal"
    if isinstance(value, int):
        return value, "integer"
    if isinstance(value, float):
        if math.isnan(value):
            return "NaN", "non_finite_number"
        if math.isinf(value):
            return "Infinity" if value > 0 else "-Infinity", "non_finite_number"
        return value, "number"
    if isinstance(value, str):
        return value, "text"
    if hasattr(value, "text"):
        return str(value.text), type(value).__name__
    return str(value), type(value).__name__


def slugify(text: str, fallback: str = "item", max_length: int = 60) -> str:
    slug = _SLUG_RE.sub("-", text.casefold()).strip("-")
    return (slug or fallback)[:max_length].rstrip("-")


def unique_preserving_order(values: Iterable[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        if value not in seen:
            result.append(value)
            seen.add(value)
    return result


def normalize_whitespace(value: object) -> str:
    return " ".join(str(value).replace("\u00a0", " ").split())


def safe_csv_text(value: Any) -> Any:
    """Neutralize spreadsheet-formula injection without changing JSON output."""
    if isinstance(value, str) and value.startswith(("=", "+", "-", "@", "\t", "\r")):
        return "'" + value
    return value

