"""Configuration for workbook ingestion.

The defaults are intentionally conservative: they work well for normal POPS
templates while bounding pathological OOXML packages and inflated worksheets.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass
from hashlib import sha256
import json


@dataclass(frozen=True, slots=True)
class ExtractionConfig:
    """Immutable settings that materially affect extraction."""

    table_score_threshold: float = 0.50
    uncertain_table_threshold: float = 0.35
    horizontal_gap: int = 2
    vertical_gap: int = 1
    max_file_bytes: int = 300 * 1024 * 1024
    max_uncompressed_bytes: int = 2 * 1024 * 1024 * 1024
    max_zip_entries: int = 30_000
    max_entry_compression_ratio: float = 1_000.0
    max_material_cells_per_sheet: int = 1_500_000
    max_table_area: int = 1_000_000
    preview_rows: int = 120
    preview_columns: int = 80
    include_cached_values: bool = True
    include_conditional_format_membership: bool = True
    include_validation_membership: bool = True
    csv_formula_injection_safe: bool = True

    def to_dict(self) -> dict[str, object]:
        return asdict(self)

    @property
    def fingerprint(self) -> str:
        payload = json.dumps(
            self.to_dict(), sort_keys=True, separators=(",", ":")
        ).encode("utf-8")
        return sha256(payload).hexdigest()

