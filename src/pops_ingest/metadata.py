"""Workbook/name metadata serialization without evaluating definitions."""

from __future__ import annotations

from typing import Any, Iterable

from openpyxl.utils import range_boundaries
from openpyxl.workbook import Workbook

from .utils import encode_typed_value


_EXCEL_MAX_ROW = 1_048_576
_EXCEL_MAX_COLUMN = 16_384


def _destination_validation_error(sheet_name: object, ref: object) -> str | None:
    """Return why a parsed name destination is unsafe as finite table geometry."""

    if not isinstance(sheet_name, str) or not sheet_name:
        return "destination sheet is empty or not text"
    if not isinstance(ref, str) or not ref:
        return "destination range is empty or not text"
    try:
        min_col, min_row, max_col, max_row = range_boundaries(ref)
    except (TypeError, ValueError) as exc:
        return f"invalid A1 range: {exc}"
    values = (min_col, min_row, max_col, max_row)
    if any(value is None for value in values):
        return "whole-row or whole-column destinations are not finite table ranges"
    if not all(isinstance(value, int) for value in values):
        return "destination boundaries are not integers"
    if not (
        1 <= min_col <= max_col <= _EXCEL_MAX_COLUMN
        and 1 <= min_row <= max_row <= _EXCEL_MAX_ROW
    ):
        return "destination is outside Excel worksheet bounds"
    return None


def _public_serializable_attributes(obj: object, names: Iterable[str]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for name in names:
        value = getattr(obj, name, None)
        if value is None:
            continue
        encoded, _ = encode_typed_value(value)
        result[name] = encoded
    return result


def workbook_metadata(workbook: Workbook) -> dict[str, Any]:
    properties = _public_serializable_attributes(
        workbook.properties,
        (
            "title",
            "subject",
            "creator",
            "keywords",
            "description",
            "lastModifiedBy",
            "revision",
            "created",
            "modified",
            "category",
            "contentStatus",
            "identifier",
            "language",
            "version",
        ),
    )
    calculation = _public_serializable_attributes(
        workbook.calculation,
        (
            "calcId",
            "calcMode",
            "fullCalcOnLoad",
            "refMode",
            "iterate",
            "iterateCount",
            "iterateDelta",
            "fullPrecision",
            "calcCompleted",
            "calcOnSave",
            "concurrentCalc",
            "forceFullCalc",
        ),
    )
    epoch = getattr(workbook, "epoch", None)
    epoch_name = "1904" if getattr(epoch, "year", None) == 1904 else "1900"
    security = _public_serializable_attributes(
        workbook.security,
        (
            "lockStructure",
            "lockWindows",
            "lockRevision",
            "workbookPassword",
            "revisionsPassword",
        ),
    )
    return {
        "properties": properties,
        "calculation": calculation,
        "date_system": epoch_name,
        "security": security,
        "iso_dates": bool(getattr(workbook, "iso_dates", False)),
        "vba_archive_loaded": getattr(workbook, "vba_archive", None) is not None,
    }


def _serialize_defined_name(name: str, defined: object, scope: str) -> dict[str, Any]:
    attr_text = getattr(defined, "attr_text", None)
    record: dict[str, Any] = {
        "name": name,
        "scope": scope,
        "definition": attr_text,
        "hidden": bool(getattr(defined, "hidden", False)),
        "function": bool(getattr(defined, "function", False)),
        "vb_procedure": bool(getattr(defined, "vbProcedure", False)),
        "xlm": bool(getattr(defined, "xlm", False)),
        "workbook_parameter": bool(getattr(defined, "workbookParameter", False)),
        "status": "raw",
        "destinations": [],
        "destination_parse_status": "not_attempted",
        "destination_parse_error": None,
    }
    definition_text = str(attr_text or "")
    if "#REF!" in definition_text.upper():
        record["status"] = "broken"
    elif any(token in definition_text.upper() for token in ("OFFSET(", "INDIRECT(")):
        record["status"] = "dynamic_unresolved"
    elif definition_text.startswith("[") or "]" in definition_text.split("!", 1)[0]:
        record["status"] = "external_unresolved"
    try:
        # `DefinedName.destinations` invokes openpyxl's formula tokenizer.
        # Real Excel files can contain stale/corrupt names such as
        # `=#REF!:#REF!`; the tokenizer raises instead of returning an empty
        # destination list. A bad name is workbook evidence, not a reason to
        # lose every selected sheet, so isolate this third-party parser edge.
        destinations = list(getattr(defined, "destinations"))
    except Exception as exc:  # noqa: BLE001 - untrusted workbook parser boundary
        destinations = []
        record["destination_parse_status"] = "failed"
        record["destination_parse_error"] = {
            "type": type(exc).__name__,
            "message": str(exc),
        }
        if record["status"] == "raw":
            record["status"] = "invalid_unparsed"
    else:
        invalid_destinations: list[dict[str, Any]] = []
        for destination in destinations:
            try:
                sheet_name, ref = destination
            except (TypeError, ValueError):
                invalid_destinations.append(
                    {
                        "raw": repr(destination),
                        "reason": "destination is not a two-item (sheet, range) pair",
                    }
                )
                continue
            reason = _destination_validation_error(sheet_name, ref)
            if reason:
                invalid_destinations.append(
                    {"sheet": sheet_name, "range": ref, "reason": reason}
                )
                continue
            record["destinations"].append({"sheet": sheet_name, "range": ref})
        if invalid_destinations:
            record["destination_parse_status"] = "invalid_destinations"
            record["destination_parse_error"] = {
                "type": "InvalidDestination",
                "message": "One or more destinations are not finite local A1 ranges.",
                "destinations": invalid_destinations,
            }
            if record["status"] == "raw":
                record["status"] = "invalid_unparsed"
        else:
            record["destination_parse_status"] = (
                "parsed" if destinations else "no_destinations"
            )
        # Never let a partially parseable name containing #REF!, an external
        # link, or a dynamic expression become a local static detector input.
        if (
            record["destinations"]
            and not invalid_destinations
            and record["status"] == "raw"
        ):
            record["status"] = "static"
    if name.startswith("_xlnm."):
        record["built_in"] = True
    return record


def defined_name_diagnostics(
    defined_names: Iterable[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Build non-fatal warnings for names that cannot be used as static ranges."""

    diagnostics: list[dict[str, Any]] = []
    for item in defined_names:
        status = str(item.get("status", "raw"))
        parse_failed = item.get("destination_parse_status") in {
            "failed",
            "invalid_destinations",
        }
        if status != "broken" and not parse_failed:
            continue
        name = str(item.get("name", "<unnamed>"))
        scope = str(item.get("scope", "workbook"))
        sheet = scope.removeprefix("sheet:") if scope.startswith("sheet:") else None
        code = (
            "DEFINED_NAME_BROKEN_REFERENCE"
            if status == "broken"
            else "DEFINED_NAME_DESTINATION_PARSE_FAILED"
        )
        message = (
            f"Defined name {name!r} contains a broken reference"
            if status == "broken"
            else f"Defined name {name!r} could not be parsed as a cell range"
        )
        diagnostics.append(
            {
                "code": code,
                "message": message
                + "; its raw definition was preserved and it was ignored for table detection.",
                "severity": "warning",
                "sheet": sheet,
                "defined_name": name,
                "details": {
                    "name": name,
                    "scope": scope,
                    "definition": item.get("definition"),
                    "status": status,
                    "destination_parse_error": item.get("destination_parse_error"),
                },
            }
        )
    return diagnostics


def defined_names_metadata(workbook: Workbook) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for name, defined in sorted(workbook.defined_names.items(), key=lambda item: item[0]):
        result.append(_serialize_defined_name(name, defined, "workbook"))
    for sheet in workbook.worksheets:
        local_names = getattr(sheet, "defined_names", {})
        for name, defined in sorted(local_names.items(), key=lambda item: item[0]):
            result.append(_serialize_defined_name(name, defined, f"sheet:{sheet.title}"))
    return result


def static_named_ranges_for_sheet(
    defined_names: list[dict[str, Any]],
    sheet_name: str,
    *,
    max_area: int | None = None,
) -> list[dict[str, Any]]:
    ranges: list[dict[str, Any]] = []
    for item in defined_names:
        if item.get("status") != "static" or item.get("built_in"):
            continue
        for destination in item.get("destinations", []):
            ref = destination.get("range")
            if destination.get("sheet") != sheet_name:
                continue
            if _destination_validation_error(sheet_name, ref):
                continue
            min_col, min_row, max_col, max_row = range_boundaries(ref)
            area = (max_col - min_col + 1) * (max_row - min_row + 1)
            if max_area is not None and area > max_area:
                continue
            ranges.append(
                {
                    "name": item["name"],
                    "ref": ref,
                    "scope": item["scope"],
                }
            )
    return ranges
