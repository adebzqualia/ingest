"""Workbook/name metadata serialization without evaluating definitions."""

from __future__ import annotations

from typing import Any, Iterable

from openpyxl.workbook import Workbook

from .utils import encode_typed_value


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
    }
    definition_text = str(attr_text or "")
    if "#REF!" in definition_text.upper():
        record["status"] = "broken"
    elif any(token in definition_text.upper() for token in ("OFFSET(", "INDIRECT(")):
        record["status"] = "dynamic_unresolved"
    elif definition_text.startswith("[") or "]" in definition_text.split("!", 1)[0]:
        record["status"] = "external_unresolved"
    try:
        destinations = list(getattr(defined, "destinations"))
    except (AttributeError, TypeError, ValueError):
        destinations = []
    for sheet_name, ref in destinations:
        record["destinations"].append({"sheet": sheet_name, "range": ref})
    if destinations:
        record["status"] = "static"
    if name.startswith("_xlnm."):
        record["built_in"] = True
    return record


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
    defined_names: list[dict[str, Any]], sheet_name: str
) -> list[dict[str, Any]]:
    ranges: list[dict[str, Any]] = []
    for item in defined_names:
        if item.get("status") != "static":
            continue
        for destination in item.get("destinations", []):
            if destination.get("sheet") == sheet_name:
                ranges.append(
                    {
                        "name": item["name"],
                        "ref": destination["range"],
                        "scope": item["scope"],
                    }
                )
    return ranges
