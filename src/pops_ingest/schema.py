"""Versioned JSON Schema shipped with every extraction bundle."""

from __future__ import annotations


BUNDLE_SCHEMA: dict[str, object] = {
    "$schema": "https://json-schema.org/draft/2020-12/schema",
    "$id": "urn:pops-ingest:schema:1.0.0",
    "title": "POPS ingestion manifest",
    "type": "object",
    "required": [
        "schema_version",
        "source",
        "selection",
        "engine",
        "workbook",
        "sheets",
        "fingerprints",
        "warnings",
    ],
    "properties": {
        "schema_version": {"const": "1.0.0"},
        "extracted_at": {"type": "string", "format": "date-time"},
        "source": {
            "type": "object",
            "required": ["filename", "size_bytes", "sha256", "format"],
            "properties": {
                "path": {"type": "string"},
                "filename": {"type": "string"},
                "size_bytes": {"type": "integer", "minimum": 0},
                "sha256": {"type": "string", "pattern": "^[0-9a-f]{64}$"},
                "format": {"enum": ["xlsx", "xlsm", "xltx", "xltm", "xlam"]},
            },
            "additionalProperties": True,
        },
        "selection": {
            "type": "object",
            "required": ["selected_sheets", "unselected_sheets"],
            "properties": {
                "selected_sheets": {"type": "array", "items": {"type": "string"}},
                "unselected_sheets": {"type": "array", "items": {"type": "string"}},
            },
            "additionalProperties": True,
        },
        "engine": {"type": "object"},
        "workbook": {"type": "object"},
        "sheets": {
            "type": "array",
            "items": {"$ref": "#/$defs/sheet_manifest"},
        },
        "fingerprints": {"type": "object"},
        "warnings": {"type": "array", "items": {"$ref": "#/$defs/warning"}},
    },
    "additionalProperties": True,
    "$defs": {
        "warning": {
            "type": "object",
            "required": ["code", "message"],
            "properties": {
                "code": {"type": "string"},
                "message": {"type": "string"},
                "severity": {"enum": ["info", "warning", "error"]},
                "sheet": {"type": ["string", "null"]},
                "coordinate": {"type": ["string", "null"]},
            },
            "additionalProperties": True,
        },
        "sheet_manifest": {
            "type": "object",
            "required": ["name", "state", "relative_path", "counts", "fingerprints", "tables"],
            "properties": {
                "name": {"type": "string"},
                "state": {"enum": ["visible", "hidden", "veryHidden"]},
                "relative_path": {"type": "string"},
                "counts": {"type": "object"},
                "fingerprints": {"type": "object"},
                "tables": {"type": "array"},
            },
            "additionalProperties": True,
        },
        "cell": {
            "type": "object",
            "required": [
                "cell_id",
                "sheet",
                "coordinate",
                "row",
                "column",
                "literal_value",
                "formula",
                "cached_value",
                "cached_value_status",
                "style_ref",
                "source",
            ],
            "properties": {
                "cell_id": {"type": "string", "pattern": "^cell_"},
                "sheet": {"type": "string"},
                "coordinate": {"type": "string"},
                "row": {"type": "integer", "minimum": 1},
                "column": {"type": "integer", "minimum": 1},
                "formula": {"type": ["string", "null"]},
                "cached_value_status": {
                    "enum": ["present", "missing", "not_requested", "not_applicable", "not_data_only"]
                },
                "style_ref": {"type": ["string", "null"]},
                "source": {"type": "object"},
            },
            "additionalProperties": True,
        },
        "logical_table": {
            "type": "object",
            "required": [
                "table_id",
                "template_key",
                "sheet",
                "range",
                "confidence",
                "columns",
                "rows",
            ],
            "properties": {
                "table_id": {"type": "string", "pattern": "^tbl_"},
                "sheet": {"type": "string"},
                "range": {"type": "string"},
                "confidence": {"type": "number", "minimum": 0, "maximum": 1},
                "columns": {"type": "array"},
                "rows": {"type": "array"},
            },
            "additionalProperties": True,
        },
        "long_record": {
            "type": "object",
            "required": [
                "table_id",
                "cell_id",
                "source_sheet",
                "source_cell",
                "value_state",
            ],
            "properties": {
                "table_id": {"type": "string", "pattern": "^tbl_"},
                "cell_id": {"type": "string", "pattern": "^cell_"},
                "source_sheet": {"type": "string"},
                "source_cell": {"type": "string"},
                "value_state": {"type": "string"},
            },
            "additionalProperties": True,
        },
    },
    "x-jsonl-item-schemas": {
        "sheets/*/cells.jsonl": "#/$defs/cell",
        "sheets/*/tables/*/table.json": "#/$defs/logical_table",
        "records_long.jsonl": "#/$defs/long_record",
    },
}
