"""Sparse, layout-aware table candidate detection.

The detector deliberately works from :class:`~pops_ingest.models.CellFeature`
objects instead of materialising a rectangular worksheet.  POPS templates
often have inflated used ranges and large styled-but-empty areas; sparse
processing keeps those workbooks bounded while retaining bordered input grids,
merged headings, validations, formulas, comments, and hyperlinks as evidence.

Native Excel table ranges supplied through ``explicit_ranges`` are forced: a
low score can add a warning, but can never remove the range.  Heuristic and
named-range candidates are scored, overlap-deduplicated, and returned in stable
sheet order.
"""

from __future__ import annotations

from bisect import bisect_left, bisect_right
from collections import defaultdict
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass
import math
import re
from typing import TypeAlias

from .config import ExtractionConfig
from .models import Bounds, CellFeature, TableCandidate
from .utils import normalize_whitespace, unique_preserving_order


FeatureCollection: TypeAlias = (
    Mapping[object, CellFeature] | Iterable[CellFeature]
)
RangeInput: TypeAlias = Bounds | TableCandidate | str | Sequence[int] | object


_HEADER_TOKEN_RE = re.compile(
    r"(?:\b20\d{2}\b|\b(?:actual|baseline|budget|forecast|plan|target)\b|"
    r"\bf\d{1,2}\s*20\d{2}\b|\bb\s*20\d{2}\b|\b(?:fte|opex|kpi)\b|"
    r"\b(?:m|k)?(?:eur|€)\b|%|\b(?:nb|number|count)\b|"
    r"(?:\bvs\.?\b|\bversus\b|\bvariance\b|\bdelta\b|[Δ∆]))",
    re.IGNORECASE,
)
_PROSE_BANNER_RE = re.compile(
    r"(?:copy\s+table|please\s+ensure|instructions?|definitions?|notes?:)",
    re.IGNORECASE,
)


@dataclass(slots=True)
class _FeatureIndex:
    """A small row-index over a normalized feature mapping."""

    by_coord: dict[tuple[int, int], CellFeature]
    by_row: dict[int, list[CellFeature]]
    rows: list[int]

    @classmethod
    def build(cls, features: FeatureCollection) -> "_FeatureIndex":
        by_coord: dict[tuple[int, int], CellFeature] = {}
        values: Iterable[CellFeature]
        if isinstance(features, Mapping):
            values = features.values()
        else:
            values = features

        for feature in values:
            if not isinstance(feature, CellFeature):
                raise TypeError(
                    "features must contain CellFeature instances; "
                    f"got {type(feature)!r}"
                )
            by_coord[(feature.row, feature.col)] = feature

        by_row: dict[int, list[CellFeature]] = defaultdict(list)
        for feature in by_coord.values():
            by_row[feature.row].append(feature)
        for row_features in by_row.values():
            row_features.sort(key=lambda item: item.col)
        return cls(by_coord, dict(by_row), sorted(by_row))

    def in_bounds(self, bounds: Bounds) -> list[CellFeature]:
        start = bisect_left(self.rows, bounds.min_row)
        stop = bisect_right(self.rows, bounds.max_row)
        result: list[CellFeature] = []
        for row in self.rows[start:stop]:
            for feature in self.by_row[row]:
                if bounds.min_col <= feature.col <= bounds.max_col:
                    result.append(feature)
        return result


class _UnionFind:
    """Minimal deterministic union-find for sparse coordinates."""

    def __init__(self, items: Iterable[tuple[int, int]]) -> None:
        self.parent = {item: item for item in items}
        self.rank = {item: 0 for item in self.parent}

    def find(self, item: tuple[int, int]) -> tuple[int, int]:
        parent = self.parent[item]
        if parent != item:
            self.parent[item] = self.find(parent)
        return self.parent[item]

    def union(self, left: tuple[int, int], right: tuple[int, int]) -> None:
        root_left = self.find(left)
        root_right = self.find(right)
        if root_left == root_right:
            return
        rank_left = self.rank[root_left]
        rank_right = self.rank[root_right]
        if rank_left < rank_right:
            root_left, root_right = root_right, root_left
            rank_left, rank_right = rank_right, rank_left
        self.parent[root_right] = root_left
        if rank_left == rank_right:
            self.rank[root_left] += 1


def _is_active(feature: CellFeature) -> bool:
    """Return whether a stored cell contributes meaningful region evidence."""

    return bool(
        feature.has_value
        or feature.has_formula
        or feature.border_edges
        or feature.merged_range
        or feature.validation_refs
        or feature.has_comment
        or feature.has_hyperlink
    )


def _short_text(feature: CellFeature) -> bool:
    text = normalize_whitespace(feature.text or "")
    return bool(text) and len(text) <= 80


def _is_prose_annotation(feature: CellFeature) -> bool:
    """Identify isolated prose that must not bridge into a nearby grid."""

    text = normalize_whitespace(feature.text or "")
    return bool(text) and (
        len(text) >= 80 or bool(_PROSE_BANNER_RE.search(text))
    )


def _can_bridge(
    left: CellFeature,
    right: CellFeature,
    *,
    axis: str,
    gap: int,
) -> bool:
    """Conservatively bridge a bounded visual gutter between active cells."""

    if gap <= 0:
        return True
    if left.merged_range and left.merged_range == right.merged_range:
        return True
    # Instructions are commonly placed one blank column to the right of a POPS
    # grid.  Their proximity alone must not enlarge the table rectangle.
    if _is_prose_annotation(left) or _is_prose_annotation(right):
        return False
    if left.style_hash and left.style_hash == right.style_hash:
        return True
    if left.border_edges and right.border_edges:
        return True
    if left.validation_refs and right.validation_refs:
        return True
    if left.has_formula and right.has_formula:
        return True
    if left.numeric and right.numeric:
        return True

    # A short row label followed by a period value commonly has a one-cell
    # visual gutter.  For vertical bridging, require stronger style/formula
    # continuity so two stacked tables separated by a blank row are less likely
    # to be fused.
    if axis == "horizontal":
        left_value = left.numeric or left.has_formula
        right_value = right.numeric or right.has_formula
        # Without border/style continuity, bridge only a label-to-value pair.
        # Two text cells separated by a gutter may be unrelated (for example a
        # period header and a side instruction), so text-to-text proximity is
        # intentionally insufficient.
        return (
            left_value != right_value
            and (left.has_value or left.has_formula)
            and (right.has_value or right.has_formula)
            and (_short_text(left) or _short_text(right))
        )
    return False


def _sparse_components(
    index: _FeatureIndex,
    config: ExtractionConfig,
    merged_ranges: Sequence[Bounds],
) -> list[Bounds]:
    """Find connected structural regions without constructing a dense mask."""

    active = {
        coord: feature
        for coord, feature in index.by_coord.items()
        if _is_active(feature)
    }
    if not active:
        return []

    union_find = _UnionFind(active)
    rows: dict[int, list[CellFeature]] = defaultdict(list)
    columns: dict[int, list[CellFeature]] = defaultdict(list)
    for feature in active.values():
        rows[feature.row].append(feature)
        columns[feature.col].append(feature)

    for row in sorted(rows):
        row_features = sorted(rows[row], key=lambda item: item.col)
        for left, right in zip(row_features, row_features[1:]):
            gap = right.col - left.col - 1
            if gap <= config.horizontal_gap and _can_bridge(
                left, right, axis="horizontal", gap=gap
            ):
                union_find.union(
                    (left.row, left.col), (right.row, right.col)
                )

    for column in sorted(columns):
        column_features = sorted(columns[column], key=lambda item: item.row)
        for upper, lower in zip(column_features, column_features[1:]):
            gap = lower.row - upper.row - 1
            if gap <= config.vertical_gap and _can_bridge(
                upper, lower, axis="vertical", gap=gap
            ):
                union_find.union(
                    (upper.row, upper.col), (lower.row, lower.col)
                )

    # Material cells inside the same merge are one logical node even if the
    # scanner stores only a subset of the styled merged cells.
    for merge in merged_ranges:
        members = [
            (feature.row, feature.col)
            for feature in index.in_bounds(merge)
            if (feature.row, feature.col) in active
        ]
        if len(members) > 1:
            anchor = members[0]
            for member in members[1:]:
                union_find.union(anchor, member)

    groups: dict[tuple[int, int], list[tuple[int, int]]] = defaultdict(list)
    for coord in sorted(active):
        groups[union_find.find(coord)].append(coord)

    bounds: list[Bounds] = []
    for coords in groups.values():
        rows_in_group = [coord[0] for coord in coords]
        cols_in_group = [coord[1] for coord in coords]
        bounds.append(
            Bounds(
                min(rows_in_group),
                min(cols_in_group),
                max(rows_in_group),
                max(cols_in_group),
            )
        )
    return sorted(bounds)


def _strip_sheet_prefix(ref: str) -> str:
    value = ref.strip()
    if "!" in value:
        value = value.rsplit("!", 1)[1]
    return value.replace("$", "").strip().strip("'")


def _coerce_bounds(value: object) -> Bounds:
    """Normalize common range representations to :class:`Bounds`."""

    if isinstance(value, Bounds):
        return value
    if isinstance(value, TableCandidate):
        return value.bounds
    if isinstance(value, str):
        return Bounds.from_a1(_strip_sheet_prefix(value))
    if isinstance(value, Mapping):
        for key in ("bounds", "range", "ref", "sqref"):
            if key in value:
                return _coerce_bounds(value[key])
    if isinstance(value, Sequence) and not isinstance(value, (str, bytes)):
        if len(value) == 4 and all(isinstance(item, int) for item in value):
            min_row, min_col, max_row, max_col = value
            return Bounds(min_row, min_col, max_row, max_col)
    if all(
        hasattr(value, attribute)
        for attribute in ("min_row", "min_col", "max_row", "max_col")
    ):
        return Bounds(
            int(getattr(value, "min_row")),
            int(getattr(value, "min_col")),
            int(getattr(value, "max_row")),
            int(getattr(value, "max_col")),
        )
    if hasattr(value, "ref"):
        return _coerce_bounds(getattr(value, "ref"))
    raise TypeError(f"Unsupported range representation: {value!r}")


def _looks_like_bounds_tuple(value: object) -> bool:
    return bool(
        isinstance(value, Sequence)
        and not isinstance(value, (str, bytes))
        and len(value) == 4
        and all(isinstance(item, int) for item in value)
    )


def _iter_named_inputs(
    values: object,
) -> Iterable[tuple[str | None, object]]:
    if values is None:
        return
    if isinstance(values, Mapping):
        for name in sorted(values, key=lambda item: str(item).casefold()):
            value = values[name]
            if (
                isinstance(value, Sequence)
                and not isinstance(value, (str, bytes, TableCandidate))
                and not _looks_like_bounds_tuple(value)
            ):
                for part in value:
                    yield str(name), part
            else:
                yield str(name), value
        return
    if isinstance(values, (str, Bounds, TableCandidate)) or _looks_like_bounds_tuple(
        values
    ):
        yield None, values
        return
    if isinstance(values, Iterable):
        for value in values:
            yield None, value
        return
    yield None, values


def _clone_candidate(candidate: TableCandidate) -> TableCandidate:
    return TableCandidate(
        bounds=candidate.bounds,
        methods=set(candidate.methods),
        confidence=float(candidate.confidence),
        features=dict(candidate.features),
        reasons=list(candidate.reasons),
        source_name=candidate.source_name,
        forced=bool(candidate.forced),
        uncertain=bool(candidate.uncertain),
    )


def _range_candidates(
    ranges: object,
    *,
    method: str,
    forced: bool,
) -> list[TableCandidate]:
    result: list[TableCandidate] = []
    for supplied_name, value in _iter_named_inputs(ranges):
        if isinstance(value, TableCandidate):
            candidate = _clone_candidate(value)
            candidate.methods.add(method)
            candidate.forced = candidate.forced or forced
            if supplied_name and not candidate.source_name:
                candidate.source_name = supplied_name
            result.append(candidate)
            continue

        name = supplied_name
        if isinstance(value, Mapping):
            raw_name = value.get("name") or value.get("source_name")
            if raw_name is not None:
                name = str(raw_name)
        try:
            bounds = _coerce_bounds(value)
        except (TypeError, ValueError):
            # Defined names can be formulas, constants, or multi-area unions;
            # only rectangular local ranges belong in table detection.
            continue
        result.append(
            TableCandidate(
                bounds=bounds,
                methods={method},
                source_name=name,
                forced=forced,
            )
        )
    return result


def _normalize_merged_ranges(values: object) -> list[Bounds]:
    result: dict[str, Bounds] = {}
    for _, value in _iter_named_inputs(values):
        try:
            bounds = _coerce_bounds(value)
        except (TypeError, ValueError):
            continue
        result[bounds.ref] = bounds
    return sorted(result.values())


def _snap_to_merges(
    bounds: Bounds,
    merged_ranges: Sequence[Bounds],
    max_area: int,
) -> Bounds:
    """Expand until every intersected merge is fully contained."""

    snapped = bounds
    changed = True
    while changed:
        changed = False
        for merge in merged_ranges:
            if not snapped.intersects(merge) or snapped.contains_bounds(merge):
                continue
            expanded = snapped.union(merge)
            if expanded.area <= max_area:
                snapped = expanded
                changed = True
    return snapped


def _mean(values: Iterable[float]) -> float:
    material = list(values)
    return sum(material) / len(material) if material else 0.0


def _jaccard(left: set[int], right: set[int]) -> float:
    union = left | right
    return len(left & right) / len(union) if union else 1.0


def _candidate_features(
    bounds: Bounds,
    cells: Sequence[CellFeature],
    methods: set[str],
    config: ExtractionConfig,
) -> dict[str, float]:
    area = max(1, bounds.area)
    material = [cell for cell in cells if _is_active(cell)]
    content = [cell for cell in cells if cell.has_value or cell.has_formula]
    text_cells = [cell for cell in content if normalize_whitespace(cell.text or "")]
    numeric_cells = [cell for cell in content if cell.numeric or cell.has_formula]
    border_cells = [cell for cell in material if cell.border_edges]

    by_row: dict[int, list[CellFeature]] = defaultdict(list)
    by_col: dict[int, list[CellFeature]] = defaultdict(list)
    for cell in material:
        by_row[cell.row].append(cell)
        by_col[cell.col].append(cell)

    active_row_sets = {
        row: {cell.col for cell in row_cells}
        for row, row_cells in by_row.items()
        if row_cells
    }
    row_set_values = [active_row_sets[row] for row in sorted(active_row_sets)]
    regularity = _mean(
        _jaccard(left, right)
        for left, right in zip(row_set_values, row_set_values[1:])
    )
    if len(row_set_values) == 1:
        regularity = min(0.5, len(row_set_values[0]) / max(2, bounds.width))

    row_span_density: list[float] = []
    for row_cells in by_row.values():
        cols = [cell.col for cell in row_cells]
        span = max(cols) - min(cols) + 1
        row_span_density.append(len(set(cols)) / span)
    row_coverage = len(by_row) / max(1, bounds.height)
    col_coverage = len(by_col) / max(1, bounds.width)
    rectangularity = (
        math.sqrt(row_coverage * col_coverage) * _mean(row_span_density)
    )

    grid = 0.0
    if material:
        border_ratio = len(border_cells) / len(material)
        edge_ratio = sum(min(cell.border_edges, 4) for cell in material) / (
            4 * len(material)
        )
        grid = 0.55 * border_ratio + 0.45 * edge_ratio

    top_limit = min(bounds.max_row, bounds.min_row + 4)
    top_cells = [cell for cell in content if cell.row <= top_limit]
    top_text = [
        normalize_whitespace(cell.text or "")
        for cell in top_cells
        if normalize_whitespace(cell.text or "")
    ]
    token_text = [text for text in top_text if _HEADER_TOKEN_RE.search(text)]
    styled_top = [
        cell for cell in top_cells if cell.bold or cell.has_fill or cell.merged_range
    ]
    token_ratio = len(token_text) / max(1, len(top_text))
    styled_ratio = len(styled_top) / max(1, len(top_cells))
    numeric_below = any(cell.row > top_limit and cell.numeric for cell in content)
    header = min(
        1.0,
        0.60 * token_ratio
        + 0.25 * styled_ratio
        + (0.15 if top_text and numeric_below else 0.0),
    )

    rows_with_axes = 0
    rows_with_values = 0
    series_rows = 0
    for row_cells in by_row.values():
        texts = [
            cell
            for cell in row_cells
            if normalize_whitespace(cell.text or "") and not cell.numeric
        ]
        values = [cell for cell in row_cells if cell.numeric or cell.has_formula]
        if values:
            rows_with_values += 1
        if len(values) >= 2:
            series_rows += 1
        if texts and values and min(cell.col for cell in texts) < max(
            cell.col for cell in values
        ):
            rows_with_axes += 1
    axes = rows_with_axes / max(1, rows_with_values)
    series = series_rows / max(1, rows_with_values)

    perimeter = [
        cell
        for cell in material
        if cell.row in {bounds.min_row, bounds.max_row}
        or cell.col in {bounds.min_col, bounds.max_col}
    ]
    boundary = (
        sum(bool(cell.border_edges) for cell in perimeter) / len(perimeter)
        if perimeter
        else 0.0
    )

    section_rows = 0
    for row_cells in by_row.values():
        row_text = [cell for cell in row_cells if normalize_whitespace(cell.text or "")]
        row_values = [cell for cell in row_cells if cell.numeric or cell.has_formula]
        if (
            0 < len(row_text) <= 2
            and not row_values
            and any(cell.bold or cell.has_fill or cell.merged_range for cell in row_text)
        ):
            section_rows += 1
    section_structure = min(1.0, section_rows / max(1, len(by_row) * 0.15))

    structural_density = min(1.0, (len(material) / area) * 2.0)
    long_text = [
        text
        for text in (normalize_whitespace(cell.text or "") for cell in text_cells)
        if len(text) >= 80 or _PROSE_BANNER_RE.search(text)
    ]
    prose_only = (
        len(long_text) / max(1, len(text_cells))
        if len(numeric_cells) <= 1
        else 0.0
    )
    background_only = float(
        not content
        and not any(cell.validation_refs for cell in material)
        and not any(cell.has_comment or cell.has_hyperlink for cell in material)
    )
    huge_sparse = float(
        bounds.area > config.max_table_area
        or (bounds.area > 10_000 and len(material) / area < 0.02)
    )
    tiny = float(bounds.area < 4 or bounds.width == 1 or bounds.height == 1)

    return {
        "grid": round(grid, 6),
        "header": round(header, 6),
        "axes": round(axes, 6),
        "rectangularity": round(rectangularity, 6),
        "regularity": round(regularity, 6),
        "series": round(series, 6),
        "boundary": round(boundary, 6),
        "section_structure": round(section_structure, 6),
        "structural_density": round(structural_density, 6),
        "prose_only": round(prose_only, 6),
        "background_only": background_only,
        "huge_sparse": huge_sparse,
        "tiny": tiny,
        "material_cells": float(len(material)),
        "content_cells": float(len(content)),
        "formula_cells": float(sum(cell.has_formula for cell in cells)),
        "numeric_cells": float(len(numeric_cells)),
        "explicit": float("explicit-range" in methods),
        "named": float("named-range" in methods),
    }


def _score_candidate(
    candidate: TableCandidate,
    index: _FeatureIndex,
    config: ExtractionConfig,
) -> TableCandidate:
    cells = index.in_bounds(candidate.bounds)
    metrics = _candidate_features(candidate.bounds, cells, candidate.methods, config)
    score = (
        0.16 * metrics["grid"]
        + 0.14 * metrics["header"]
        + 0.12 * metrics["axes"]
        + 0.10 * metrics["rectangularity"]
        + 0.09 * metrics["regularity"]
        + 0.08 * metrics["series"]
        + 0.07 * metrics["boundary"]
        + 0.06 * metrics["section_structure"]
        + 0.08 * metrics["structural_density"]
        + 0.18 * metrics["explicit"]
        + 0.06 * metrics["named"]
        - 0.18 * metrics["prose_only"]
        - 0.18 * metrics["background_only"]
        - 0.12 * metrics["huge_sparse"]
        - 0.12 * metrics["tiny"]
    )
    score = max(0.0, min(1.0, score))
    if candidate.forced:
        score = max(score, 0.98)

    reasons: list[str] = []
    labels = {
        "grid": "continuous border/grid scaffold",
        "header": "header or period/scenario evidence",
        "axes": "label and value axes",
        "rectangularity": "rectangular structural occupancy",
        "regularity": "repeated row/column structure",
        "series": "repeated numeric/formula series",
        "section_structure": "section-band structure",
    }
    for key, label in labels.items():
        if metrics[key] >= 0.55:
            reasons.append(label)
    if metrics["prose_only"] >= 0.5:
        reasons.append("annotation-like prose dominates")
    if metrics["background_only"]:
        reasons.append("style/background evidence without content")
    if metrics["huge_sparse"]:
        reasons.append("oversized or very sparse range")
    if candidate.forced:
        reasons.append("forced native range")

    candidate.features = metrics
    candidate.confidence = round(score, 6)
    candidate.reasons = unique_preserving_order(candidate.reasons + reasons)
    return candidate


def _merge_candidate_metadata(
    winner: TableCandidate, duplicate: TableCandidate
) -> None:
    winner.methods.update(duplicate.methods)
    winner.confidence = max(winner.confidence, duplicate.confidence)
    winner.forced = winner.forced or duplicate.forced
    winner.uncertain = winner.uncertain and duplicate.uncertain
    winner.reasons = unique_preserving_order(
        winner.reasons + duplicate.reasons
    )
    for key, value in duplicate.features.items():
        winner.features[key] = max(winner.features.get(key, value), value)
    if not winner.source_name and duplicate.source_name:
        winner.source_name = duplicate.source_name


def _deduplicate(candidates: Sequence[TableCandidate]) -> list[TableCandidate]:
    """Remove peer duplicates while retaining every distinct forced range."""

    exact: dict[Bounds, TableCandidate] = {}
    for candidate in sorted(
        candidates,
        key=lambda item: (
            item.bounds,
            not item.forced,
            -item.confidence,
            item.source_name or "",
        ),
    ):
        existing = exact.get(candidate.bounds)
        if existing is None:
            exact[candidate.bounds] = candidate
        else:
            # Exact native/named/heuristic representations are one range; the
            # forced flag and all evidence are folded into the retained object.
            _merge_candidate_metadata(existing, candidate)

    ranked = sorted(
        exact.values(),
        key=lambda item: (
            not item.forced,
            -item.confidence,
            -item.bounds.area,
            item.bounds,
            item.source_name or "",
        ),
    )
    kept: list[TableCandidate] = []
    for candidate in ranked:
        duplicate_of: TableCandidate | None = None
        for existing in kept:
            if candidate.forced and existing.forced:
                continue
            intersection = candidate.bounds.intersection_area(existing.bounds)
            if not intersection:
                continue
            smaller_area = min(candidate.bounds.area, existing.bounds.area)
            containment = intersection / smaller_area
            area_ratio = max(candidate.bounds.area, existing.bounds.area) / smaller_area
            near_duplicate = candidate.bounds.iou(existing.bounds) >= 0.78
            near_same_container = containment >= 0.97 and area_ratio <= 1.35
            peer_overlap = candidate.bounds.iou(existing.bounds) >= 0.55

            if near_duplicate or near_same_container or peer_overlap:
                duplicate_of = existing
                break
            if containment >= 0.97:
                if existing.forced and existing.bounds.contains_bounds(
                    candidate.bounds
                ):
                    duplicate_of = existing
                    break
                # A much larger heuristic container around a forced native
                # range can be a valid compound table, so retain both.
                if not existing.forced and not candidate.forced:
                    duplicate_of = existing
                    break

        if duplicate_of is None:
            kept.append(candidate)
        else:
            _merge_candidate_metadata(duplicate_of, candidate)

    return kept


def detect_table_candidates(
    features: FeatureCollection,
    explicit_ranges: object,
    named_ranges: object,
    merged_ranges: object,
    config: ExtractionConfig,
) -> tuple[list[TableCandidate], list[TableCandidate]]:
    """Detect and rank table-like regions from sparse worksheet features.

    Parameters
    ----------
    features:
        Mapping or iterable of :class:`CellFeature` instances for one sheet.
    explicit_ranges:
        Native Excel table ranges.  Accepted representations include
        :class:`Bounds`, A1 strings, ``TableCandidate`` objects, four-integer
        ``(min_row, min_col, max_row, max_col)`` sequences, mappings of names to
        any of those, and openpyxl range-like objects.  Every explicit range is
        forced into the accepted result.
    named_ranges:
        Rectangular sheet-local defined ranges in the same representations.
        Formula, constant, and multi-area names are ignored.
    merged_ranges:
        Merged ranges used to connect and snap heuristic candidates.
    config:
        Extraction thresholds and bounded horizontal/vertical gap sizes.

    Returns
    -------
    (accepted, rejected):
        Accepted candidates include uncertain candidates whose confidence is
        between ``uncertain_table_threshold`` and ``table_score_threshold``.
        Such candidates have ``uncertain=True``.  The rejected list contains
        low-scoring regions suitable for annotation inference.  Both lists are
        deterministically ordered by worksheet position.
    """

    index = _FeatureIndex.build(features)
    merges = _normalize_merged_ranges(merged_ranges)

    candidates: list[TableCandidate] = []
    candidates.extend(
        _range_candidates(
            explicit_ranges, method="explicit-range", forced=True
        )
    )
    candidates.extend(
        _range_candidates(named_ranges, method="named-range", forced=False)
    )

    for bounds in _sparse_components(index, config, merges):
        snapped = _snap_to_merges(bounds, merges, config.max_table_area)
        candidates.append(
            TableCandidate(bounds=snapped, methods={"sparse-component"})
        )

    # Native/named ranges must also respect intersecting merges; forced native
    # ranges are allowed to exceed the heuristic area cap because the workbook
    # author explicitly declared them.
    for candidate in candidates:
        cap = max(config.max_table_area, candidate.bounds.area) if candidate.forced else config.max_table_area
        candidate.bounds = _snap_to_merges(candidate.bounds, merges, cap)
        _score_candidate(candidate, index, config)

    deduplicated = _deduplicate(candidates)
    accepted: list[TableCandidate] = []
    rejected: list[TableCandidate] = []
    for candidate in deduplicated:
        if candidate.forced:
            candidate.uncertain = False
            accepted.append(candidate)
        elif candidate.confidence >= config.uncertain_table_threshold:
            candidate.uncertain = (
                candidate.confidence < config.table_score_threshold
            )
            if candidate.uncertain:
                candidate.reasons = unique_preserving_order(
                    candidate.reasons
                    + ["below automatic table confidence threshold"]
                )
            accepted.append(candidate)
        else:
            candidate.uncertain = True
            rejected.append(candidate)

    order_key = lambda item: (
        item.bounds.min_row,
        item.bounds.min_col,
        item.bounds.max_row,
        item.bounds.max_col,
        item.source_name or "",
        tuple(sorted(item.methods)),
    )
    accepted.sort(key=order_key)
    rejected.sort(key=order_key)
    return accepted, rejected


__all__ = ["detect_table_candidates"]
