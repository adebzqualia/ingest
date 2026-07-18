"""Sheet catalog display and explicit/interactive selection."""

from __future__ import annotations

from fnmatch import fnmatchcase
import sys
from typing import Iterable, Sequence


class SheetSelectionError(ValueError):
    """Raised when a sheet selector is missing, ambiguous, or duplicated."""


def format_sheet_catalog(catalog: Sequence[dict[str, object]]) -> str:
    lines = ["#  State        Sheet name                         XML dimension  Formulas"]
    lines.append("-  -----------  ---------------------------------  -------------  --------")
    for index, sheet in enumerate(catalog, 1):
        name = str(sheet.get("name", ""))
        state = str(sheet.get("state", "visible"))
        dimension = str(sheet.get("stored_dimension") or "-")
        formula_count = str(sheet.get("formula_count", "-"))
        lines.append(
            f"{index:<2} {state:<11}  {name[:33]:<33}  "
            f"{dimension[:13]:<13}  {formula_count:>8}"
        )
    return "\n".join(lines)


def _add_unique(result: list[str], seen: set[str], name: str, source: str) -> None:
    if name in seen:
        raise SheetSelectionError(
            f"Sheet {name!r} was selected more than once (latest selector: {source!r})."
        )
    result.append(name)
    seen.add(name)


def resolve_sheet_selection(
    catalog: Sequence[dict[str, object]],
    *,
    sheet_names: Iterable[str] = (),
    sheet_globs: Iterable[str] = (),
    index_selectors: Iterable[str] = (),
    all_sheets: bool = False,
    visible_sheets: bool = False,
) -> list[str]:
    """Resolve selectors while preserving original workbook order.

    Exact names, globs, and one-based index/range selectors are deliberately
    separate so a numeric sheet name is never mistaken for an index.
    """

    ordered_names = [str(sheet["name"]) for sheet in catalog]
    state_by_name = {
        str(sheet["name"]): str(sheet.get("state", "visible")) for sheet in catalog
    }

    selection_modes = sum(
        bool(value)
        for value in (
            list(sheet_names),
            list(sheet_globs),
            list(index_selectors),
            all_sheets,
            visible_sheets,
        )
    )
    if selection_modes == 0:
        return []
    if all_sheets and visible_sheets:
        raise SheetSelectionError("Use either all sheets or visible sheets, not both.")

    requested: list[str] = []
    seen: set[str] = set()

    if all_sheets:
        return ordered_names
    if visible_sheets:
        return [name for name in ordered_names if state_by_name[name] == "visible"]

    for name in sheet_names:
        if name not in state_by_name:
            suggestions = [n for n in ordered_names if n.casefold() == name.casefold()]
            suffix = f" Did you mean {suggestions[0]!r}?" if len(suggestions) == 1 else ""
            raise SheetSelectionError(f"Workbook has no sheet named {name!r}.{suffix}")
        _add_unique(requested, seen, name, name)

    for pattern in sheet_globs:
        matches = [name for name in ordered_names if fnmatchcase(name, pattern)]
        if not matches:
            raise SheetSelectionError(f"Sheet glob {pattern!r} matched no sheets.")
        for name in matches:
            _add_unique(requested, seen, name, pattern)

    for selector in index_selectors:
        for token in (part.strip() for part in selector.split(",")):
            if not token:
                continue
            if "-" in token:
                pieces = token.split("-", 1)
                if not all(piece.strip().isdigit() for piece in pieces):
                    raise SheetSelectionError(f"Invalid sheet index range {token!r}.")
                start, end = (int(piece.strip()) for piece in pieces)
                if start > end:
                    raise SheetSelectionError(f"Descending sheet range {token!r} is invalid.")
                indexes = range(start, end + 1)
            elif token.isdigit():
                indexes = (int(token),)
            else:
                raise SheetSelectionError(f"Invalid sheet index selector {token!r}.")
            for index in indexes:
                if index < 1 or index > len(ordered_names):
                    raise SheetSelectionError(
                        f"Sheet index {index} is outside 1..{len(ordered_names)}."
                    )
                _add_unique(requested, seen, ordered_names[index - 1], token)

    requested_set = set(requested)
    return [name for name in ordered_names if name in requested_set]


def prompt_for_sheets(catalog: Sequence[dict[str, object]]) -> list[str]:
    """Prompt for a numbered multi-selection before workbook extraction."""

    if not sys.stdin.isatty():
        raise SheetSelectionError(
            "No sheets were selected and stdin is not interactive. Use --sheet, "
            "--sheet-index, --sheet-glob, --all-sheets, or --visible-sheets."
        )
    print(format_sheet_catalog(catalog))
    print("\nChoose sheets now. Enter indexes/ranges (for example 2,4-7), 'all', or 'visible'.")
    while True:
        try:
            answer = input("Sheets to extract: ").strip()
        except EOFError as exc:
            raise SheetSelectionError("Sheet selection was cancelled.") from exc
        try:
            if answer.casefold() == "all":
                return resolve_sheet_selection(catalog, all_sheets=True)
            if answer.casefold() == "visible":
                return resolve_sheet_selection(catalog, visible_sheets=True)
            selected = resolve_sheet_selection(catalog, index_selectors=[answer])
            if selected:
                return selected
            print("Select at least one sheet.")
        except SheetSelectionError as exc:
            print(f"Invalid selection: {exc}")

