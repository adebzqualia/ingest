"""Read-only, safety-bounded indexing of an OOXML workbook package.

This module deliberately works below :mod:`openpyxl`.  The high-level reader is
excellent for normal cell access, but it does not expose every package part or
every fragment of formula/drawing XML needed to build a durable reference
fingerprint.  ``OOXMLIndex`` therefore records the package and selected raw XML
facts without evaluating formulas, executing macros, refreshing connections, or
following external relationships.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from hashlib import sha256
import math
import posixpath
from pathlib import Path, PurePosixPath, PureWindowsPath
import re
import stat
from typing import Any, BinaryIO
from urllib.parse import unquote, urlsplit
import xml.etree.ElementTree as ET
import zipfile

from .config import ExtractionConfig
from .models import WarningRecord
from .utils import sha256_file


_SUPPORTED_EXTENSIONS = frozenset({".xlsx", ".xlsm", ".xltx", ".xltm", ".xlam"})
_OFFICE_REL_SUFFIX = "/officeDocument"
_WORKSHEET_REL_SUFFIXES = (
    "/worksheet",
    "/chartsheet",
    "/dialogsheet",
    "/macrosheet",
)
_DRAWING_REL_SUFFIXES = ("/drawing", "/vmlDrawing")
_DRAWING_ELEMENT_NAMES = frozenset(
    {"sp", "pic", "graphicFrame", "cxnSp", "grpSp", "contentPart", "shape"}
)
_ANCHOR_ELEMENT_NAMES = frozenset({"twoCellAnchor", "oneCellAnchor", "absoluteAnchor"})
_EXTERNAL_URI_RE = re.compile(r"^[A-Za-z][A-Za-z0-9+.-]*:")


class OOXMLIndexError(ValueError):
    """Base class for an OOXML package that cannot safely be indexed."""


class OOXMLSecurityError(OOXMLIndexError):
    """Raised when package preflight violates a configured safety bound."""


class OOXMLFormatError(OOXMLIndexError):
    """Raised when required OOXML package structures are missing or malformed."""


@dataclass(slots=True)
class OOXMLIndex:
    """A fully materialized, read-only inventory of an OOXML workbook.

    Use :meth:`open`; callers should not instantiate this class directly.  No
    ``ZipFile`` handle is retained after construction, so an index never holds a
    lock on the source workbook.
    """

    path: Path
    config: ExtractionConfig
    package_manifest: dict[str, Any] = field(default_factory=dict)
    sheet_catalog: list[dict[str, Any]] = field(default_factory=list)
    formula_records: dict[str, dict[str, dict[str, Any]]] = field(default_factory=dict)
    drawings_by_sheet: dict[str, list[dict[str, Any]]] = field(default_factory=dict)
    external_relationships: list[dict[str, Any]] = field(default_factory=list)
    warnings: list[WarningRecord] = field(default_factory=list)

    @classmethod
    def open(cls, path: Path, config: ExtractionConfig) -> "OOXMLIndex":
        """Safely inspect *path* and return an in-memory package index.

        Safety limits are checked from the central directory before any package
        member is decompressed.  The workbook is opened read-only and is never
        rewritten.
        """

        source = Path(path)
        index = cls(path=source, config=config)
        index._build()
        return index

    def formula_record(self, sheet_name: str, coordinate: str) -> dict[str, Any] | None:
        """Return raw formula metadata for ``sheet_name!coordinate``, if present.

        Coordinates are accepted case-insensitively.  The returned mapping is the
        stored record (not a copy), matching normal dictionary lookup semantics.
        """

        records = self.formula_records.get(sheet_name)
        if records is None:
            return None
        return records.get(str(coordinate).upper())

    # ------------------------------------------------------------------ build

    def _build(self) -> None:
        resolved = self.path.expanduser().resolve(strict=False)
        self.path = resolved
        self._preflight_path()

        try:
            with zipfile.ZipFile(self.path, mode="r") as package:
                infos, preflight = self._preflight_zip(package)
                content_types = self._read_content_types(package)
                parts, parts_by_name = self._inventory_parts(
                    package, infos, content_types
                )
                relationships = self._read_all_relationships(package, parts_by_name)
                self.external_relationships = [
                    dict(item) for item in relationships if item["external"]
                ]

                workbook_part = self._find_workbook_part(
                    relationships, parts_by_name
                )
                self.sheet_catalog = self._read_sheet_catalog(
                    package, workbook_part, relationships, parts_by_name
                )

                relationship_lookup = {
                    (item["source_part"], item["id"]): item
                    for item in relationships
                }
                for sheet in self.sheet_catalog:
                    self._index_sheet(package, sheet, relationship_lookup)
                self._index_drawings(package, relationship_lookup, parts_by_name)

                features, feature_parts = self._detect_features(parts)
                self.package_manifest = {
                    "source_path": str(self.path),
                    "extension": self.path.suffix.casefold(),
                    "file_size_bytes": self.path.stat().st_size,
                    "sha256": sha256_file(self.path),
                    "config_fingerprint": self.config.fingerprint,
                    "workbook_part": workbook_part,
                    **preflight,
                    "content_types": content_types,
                    "features": features,
                    "feature_parts": feature_parts,
                    "has_macros": features["macros"],
                    "has_queries": features["queries"],
                    "has_connections": features["connections"],
                    "has_threaded_comments": features["threaded_comments"],
                    "parts": parts,
                    "parts_by_name": parts_by_name,
                    "relationship_count": len(relationships),
                    "external_relationship_count": len(self.external_relationships),
                }

                self._add_feature_warnings(features, feature_parts)
                if self.external_relationships:
                    self._warn(
                        "external_relationships_present",
                        f"Package contains {len(self.external_relationships)} external "
                        "relationship(s); targets were inventoried but not followed.",
                        details={"count": len(self.external_relationships)},
                    )
        except zipfile.BadZipFile as exc:
            raise OOXMLFormatError(f"Invalid or corrupt OOXML ZIP package: {self.path}") from exc

    # --------------------------------------------------------------- preflight

    def _preflight_path(self) -> None:
        suffix = self.path.suffix.casefold()
        if suffix not in _SUPPORTED_EXTENSIONS:
            supported = ", ".join(sorted(_SUPPORTED_EXTENSIONS))
            raise OOXMLFormatError(
                f"Unsupported workbook extension {suffix or '<none>'!r}; expected one of "
                f"{supported}. Legacy .xls and binary .xlsb files are not OOXML workbooks."
            )
        if not self.path.exists():
            raise OOXMLFormatError(f"Workbook does not exist: {self.path}")
        if not self.path.is_file():
            raise OOXMLFormatError(f"Workbook path is not a regular file: {self.path}")
        size = self.path.stat().st_size
        if size > self.config.max_file_bytes:
            raise OOXMLSecurityError(
                f"Workbook is {size} bytes, above max_file_bytes="
                f"{self.config.max_file_bytes}."
            )
        if size < 4 or not zipfile.is_zipfile(self.path):
            raise OOXMLFormatError(f"File is not a valid ZIP-based OOXML package: {self.path}")

    def _preflight_zip(
        self, package: zipfile.ZipFile
    ) -> tuple[list[zipfile.ZipInfo], dict[str, Any]]:
        infos = package.infolist()
        if len(infos) > self.config.max_zip_entries:
            raise OOXMLSecurityError(
                f"Package has {len(infos)} entries, above max_zip_entries="
                f"{self.config.max_zip_entries}."
            )

        seen_exact: set[str] = set()
        seen_normalized: set[str] = set()
        total_uncompressed = 0
        total_compressed = 0
        maximum_ratio = 1.0

        for info in infos:
            name = info.filename
            normalized = _validate_member_name(name)
            if name in seen_exact or normalized.casefold() in seen_normalized:
                raise OOXMLSecurityError(
                    f"Package contains a duplicate or case-colliding member: {name!r}."
                )
            seen_exact.add(name)
            seen_normalized.add(normalized.casefold())

            if info.flag_bits & 0x1:
                raise OOXMLSecurityError(
                    f"Encrypted ZIP member is not supported: {name!r}."
                )
            mode = (info.external_attr >> 16) & 0xFFFF
            if mode and stat.S_ISLNK(mode):
                raise OOXMLSecurityError(
                    f"Symbolic-link ZIP member is not allowed: {name!r}."
                )

            total_uncompressed += info.file_size
            total_compressed += info.compress_size
            if total_uncompressed > self.config.max_uncompressed_bytes:
                raise OOXMLSecurityError(
                    "Package uncompressed size exceeds max_uncompressed_bytes="
                    f"{self.config.max_uncompressed_bytes}."
                )

            ratio = _compression_ratio(info)
            maximum_ratio = max(maximum_ratio, ratio)
            if ratio > self.config.max_entry_compression_ratio:
                raise OOXMLSecurityError(
                    f"ZIP member {name!r} has compression ratio {ratio:.1f}, above "
                    f"max_entry_compression_ratio="
                    f"{self.config.max_entry_compression_ratio}."
                )

        required = {"[Content_Types].xml", "_rels/.rels"}
        available = {_normalize_member_name(info.filename) for info in infos}
        missing = sorted(required - available)
        if missing:
            raise OOXMLFormatError(
                "Missing required OOXML package part(s): " + ", ".join(missing)
            )

        return infos, {
            "zip_entry_count": len(infos),
            "zip_file_part_count": sum(not info.is_dir() for info in infos),
            "total_uncompressed_bytes": total_uncompressed,
            "total_compressed_bytes": total_compressed,
            "maximum_entry_compression_ratio": maximum_ratio,
        }

    # ------------------------------------------------------- package inventory

    def _read_content_types(self, package: zipfile.ZipFile) -> dict[str, Any]:
        defaults: dict[str, str] = {}
        overrides: dict[str, str] = {}
        try:
            with package.open("[Content_Types].xml", "r") as stream:
                _reject_doctype(stream, "[Content_Types].xml")
            with package.open("[Content_Types].xml", "r") as stream:
                for _event, element in ET.iterparse(stream, events=("end",)):
                    local = _local_name(element.tag)
                    if local == "Default":
                        extension = element.attrib.get("Extension", "").casefold()
                        if extension:
                            defaults[extension] = element.attrib.get("ContentType", "")
                    elif local == "Override":
                        part = _normalize_member_name(
                            element.attrib.get("PartName", "").lstrip("/")
                        )
                        if part:
                            overrides[part] = element.attrib.get("ContentType", "")
                    element.clear()
        except (ET.ParseError, KeyError, RuntimeError) as exc:
            raise OOXMLFormatError("Unable to parse [Content_Types].xml") from exc
        return {
            "defaults": dict(sorted(defaults.items())),
            "overrides": dict(sorted(overrides.items())),
        }

    def _inventory_parts(
        self,
        package: zipfile.ZipFile,
        infos: list[zipfile.ZipInfo],
        content_types: dict[str, Any],
    ) -> tuple[list[dict[str, Any]], dict[str, dict[str, Any]]]:
        defaults: dict[str, str] = content_types["defaults"]
        overrides: dict[str, str] = content_types["overrides"]
        parts: list[dict[str, Any]] = []
        parts_by_name: dict[str, dict[str, Any]] = {}

        for entry_index, info in enumerate(infos):
            if info.is_dir():
                continue
            normalized = _normalize_member_name(info.filename)
            extension = PurePosixPath(normalized).suffix.lstrip(".").casefold()
            content_type = overrides.get(normalized) or defaults.get(extension)
            digest = sha256()
            try:
                with package.open(info, "r") as stream:
                    for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                        digest.update(chunk)
            except (zipfile.BadZipFile, RuntimeError, OSError) as exc:
                raise OOXMLFormatError(
                    f"Unable to read ZIP member {info.filename!r}."
                ) from exc

            part = {
                "name": normalized,
                "zip_name": info.filename,
                "entry_index": entry_index,
                "content_type": content_type,
                "size_bytes": info.file_size,
                "compressed_size_bytes": info.compress_size,
                "compression_ratio": _compression_ratio(info),
                "crc32": f"{info.CRC:08x}",
                "sha256": digest.hexdigest(),
            }
            parts.append(part)
            parts_by_name[normalized] = part
        return parts, parts_by_name

    def _read_all_relationships(
        self, package: zipfile.ZipFile, parts_by_name: dict[str, dict[str, Any]]
    ) -> list[dict[str, Any]]:
        relationships: list[dict[str, Any]] = []
        rel_parts = sorted(
            name
            for name in parts_by_name
            if name == "_rels/.rels"
            or ("/_rels/" in name and name.casefold().endswith(".rels"))
        )
        for rels_part in rel_parts:
            source_part = _source_part_from_rels_part(rels_part)
            seen_ids: set[str] = set()
            try:
                with package.open(parts_by_name[rels_part]["zip_name"], "r") as stream:
                    _reject_doctype(stream, rels_part)
                with package.open(parts_by_name[rels_part]["zip_name"], "r") as stream:
                    for _event, element in ET.iterparse(stream, events=("end",)):
                        if _local_name(element.tag) != "Relationship":
                            element.clear()
                            continue
                        rel_id = element.attrib.get("Id", "")
                        target = element.attrib.get("Target", "")
                        rel_type = element.attrib.get("Type", "")
                        target_mode = element.attrib.get("TargetMode")
                        external = _is_external_target(target, target_mode)
                        resolved: str | None = None
                        resolution_error: str | None = None
                        if not external:
                            try:
                                resolved = _resolve_relationship_target(
                                    source_part, target
                                )
                            except OOXMLFormatError as exc:
                                resolution_error = str(exc)

                        record = {
                            "source_part": source_part,
                            "rels_part": rels_part,
                            "id": rel_id,
                            "type": rel_type,
                            "target": target,
                            "target_mode": target_mode,
                            "external": external,
                            "resolved_target": resolved,
                        }
                        if resolution_error:
                            record["resolution_error"] = resolution_error
                            self._warn(
                                "invalid_relationship_target",
                                f"Invalid internal relationship target in {rels_part}: "
                                f"{target!r}.",
                                details={"relationship_id": rel_id, "error": resolution_error},
                            )
                        if rel_id in seen_ids:
                            self._warn(
                                "duplicate_relationship_id",
                                f"Relationship part {rels_part} repeats Id {rel_id!r}.",
                                details={"rels_part": rels_part, "relationship_id": rel_id},
                            )
                        seen_ids.add(rel_id)
                        relationships.append(record)
                        element.clear()
            except (ET.ParseError, KeyError, RuntimeError) as exc:
                self._warn(
                    "relationship_parse_error",
                    f"Could not parse relationship part {rels_part!r}.",
                    severity="error",
                    details={"error": str(exc)},
                )
        return relationships

    def _find_workbook_part(
        self,
        relationships: list[dict[str, Any]],
        parts_by_name: dict[str, dict[str, Any]],
    ) -> str:
        candidates = [
            item
            for item in relationships
            if item["source_part"] == ""
            and str(item["type"]).endswith(_OFFICE_REL_SUFFIX)
            and not item["external"]
            and item["resolved_target"]
        ]
        if candidates:
            workbook_part = str(candidates[0]["resolved_target"])
        elif "xl/workbook.xml" in parts_by_name:
            workbook_part = "xl/workbook.xml"
            self._warn(
                "workbook_relationship_missing",
                "Root officeDocument relationship is missing; using xl/workbook.xml.",
            )
        else:
            raise OOXMLFormatError(
                "Root relationships do not identify a workbook part."
            )
        if workbook_part not in parts_by_name:
            raise OOXMLFormatError(
                f"Workbook part referenced by package is missing: {workbook_part!r}."
            )
        return workbook_part

    # ------------------------------------------------------------- worksheets

    def _read_sheet_catalog(
        self,
        package: zipfile.ZipFile,
        workbook_part: str,
        relationships: list[dict[str, Any]],
        parts_by_name: dict[str, dict[str, Any]],
    ) -> list[dict[str, Any]]:
        rel_lookup = {
            item["id"]: item
            for item in relationships
            if item["source_part"] == workbook_part
        }
        sheets: list[dict[str, Any]] = []
        seen_names: set[str] = set()
        zip_name = parts_by_name[workbook_part]["zip_name"]
        try:
            with package.open(zip_name, "r") as stream:
                _reject_doctype(stream, workbook_part)
            with package.open(zip_name, "r") as stream:
                for _event, element in ET.iterparse(stream, events=("end",)):
                    if _local_name(element.tag) != "sheet":
                        element.clear()
                        continue
                    name = element.attrib.get("name", "")
                    relationship_id = _attribute_by_local_name(element, "id")
                    relation = rel_lookup.get(relationship_id or "")
                    part = relation.get("resolved_target") if relation else None
                    state = element.attrib.get("state", "visible")
                    sheet = {
                        "index": len(sheets),
                        "position": len(sheets) + 1,
                        "name": name,
                        "sheet_id": element.attrib.get("sheetId"),
                        "state": state,
                        "relationship_id": relationship_id,
                        "relationship_type": relation.get("type") if relation else None,
                        "target": relation.get("target") if relation else None,
                        "part": part,
                        "part_present": bool(part and part in parts_by_name),
                        "stored_dimension": None,
                        "stored_cell_count": 0,
                        "formula_count": 0,
                        "formula_cached_value_count": 0,
                        "formula_missing_cached_value_count": 0,
                        "drawing_relationships": [],
                    }
                    if name.casefold() in seen_names:
                        self._warn(
                            "duplicate_sheet_name",
                            f"Workbook repeats sheet name {name!r}.",
                            sheet=name,
                        )
                    seen_names.add(name.casefold())
                    if relation is None:
                        self._warn(
                            "sheet_relationship_missing",
                            f"Sheet {name!r} has no matching workbook relationship.",
                            sheet=name,
                            details={"relationship_id": relationship_id},
                        )
                    elif relation["external"]:
                        self._warn(
                            "external_sheet_relationship",
                            f"Sheet {name!r} unexpectedly has an external relationship.",
                            sheet=name,
                        )
                    elif not sheet["part_present"]:
                        self._warn(
                            "sheet_part_missing",
                            f"Worksheet part for {name!r} is missing: {part!r}.",
                            sheet=name,
                        )
                    sheets.append(sheet)
                    element.clear()
        except (ET.ParseError, KeyError, RuntimeError) as exc:
            raise OOXMLFormatError(
                f"Unable to parse workbook part {workbook_part!r}."
            ) from exc
        return sheets

    def _index_sheet(
        self,
        package: zipfile.ZipFile,
        sheet: dict[str, Any],
        relationship_lookup: dict[tuple[str, str], dict[str, Any]],
    ) -> None:
        name = str(sheet["name"])
        self.formula_records.setdefault(name, {})
        self.drawings_by_sheet.setdefault(name, [])
        part = sheet.get("part")
        if not part or not sheet.get("part_present"):
            return
        rel_type = str(sheet.get("relationship_type") or "")
        if rel_type and not rel_type.endswith(_WORKSHEET_REL_SUFFIXES):
            return

        records = self.formula_records[name]
        drawing_refs: list[dict[str, str | None]] = []
        stored_dimension: str | None = None
        cell_count = 0
        formula_count = 0
        cache_count = 0
        missing_cache_count = 0
        try:
            with package.open(part, "r") as stream:
                _reject_doctype(stream, part)
            with package.open(part, "r") as stream:
                for _event, element in ET.iterparse(stream, events=("end",)):
                    local = _local_name(element.tag)
                    if local == "dimension" and stored_dimension is None:
                        stored_dimension = element.attrib.get("ref")
                    elif local == "c":
                        cell_count += 1
                        formula_element = _first_child(element, "f")
                        if formula_element is not None:
                            formula_count += 1
                            value_element = _first_child(element, "v")
                            has_cache = value_element is not None
                            if has_cache:
                                cache_count += 1
                            else:
                                missing_cache_count += 1
                            raw_coordinate = element.attrib.get("r")
                            if not raw_coordinate:
                                self._warn(
                                    "formula_coordinate_missing",
                                    "Formula cell has no stored coordinate and was skipped.",
                                    sheet=name,
                                    details={"sheet_part": part},
                                )
                            else:
                                coordinate = raw_coordinate.upper()
                                formula_attrs_raw = dict(formula_element.attrib)
                                formula_attrs = {
                                    _local_name(key): value
                                    for key, value in formula_attrs_raw.items()
                                }
                                record = {
                                    "sheet": name,
                                    "coordinate": coordinate,
                                    "source_coordinate_raw": raw_coordinate,
                                    "formula_text": formula_element.text,
                                    "formula_attributes": formula_attrs,
                                    "formula_attributes_raw": formula_attrs_raw,
                                    "cached_value_present": has_cache,
                                    "cached_value_raw": (
                                        value_element.text
                                        if has_cache and self.config.include_cached_values
                                        else None
                                    ),
                                    "cached_value_included": bool(
                                        has_cache and self.config.include_cached_values
                                    ),
                                    "cell_type": element.attrib.get("t"),
                                    "style_index": element.attrib.get("s"),
                                    "cell_attributes": {
                                        _local_name(key): value
                                        for key, value in element.attrib.items()
                                    },
                                    "cell_attributes_raw": dict(element.attrib),
                                    "sheet_part": part,
                                }
                                if coordinate in records:
                                    self._warn(
                                        "duplicate_formula_coordinate",
                                        f"Worksheet XML repeats formula coordinate {coordinate}.",
                                        sheet=name,
                                        coordinate=coordinate,
                                    )
                                records[coordinate] = record
                        element.clear()
                    elif local in {"drawing", "legacyDrawing", "legacyDrawingHF"}:
                        relationship_id = _attribute_by_local_name(element, "id")
                        relation = relationship_lookup.get((part, relationship_id or ""))
                        drawing_refs.append(
                            {
                                "element_type": local,
                                "relationship_id": relationship_id,
                                "relationship_type": (
                                    relation.get("type") if relation else None
                                ),
                                "target": relation.get("target") if relation else None,
                                "part": (
                                    relation.get("resolved_target") if relation else None
                                ),
                                "external": (
                                    bool(relation.get("external")) if relation else False
                                ),
                            }
                        )
                        element.clear()
                    elif local in {"row", "sheetData", "worksheet"}:
                        element.clear()
        except (ET.ParseError, KeyError, RuntimeError) as exc:
            self._warn(
                "sheet_xml_parse_error",
                f"Could not fully parse worksheet XML for {name!r}.",
                severity="error",
                sheet=name,
                details={"sheet_part": part, "error": str(exc)},
            )

        sheet["stored_dimension"] = stored_dimension
        sheet["stored_cell_count"] = cell_count
        sheet["formula_count"] = formula_count
        sheet["formula_cached_value_count"] = cache_count
        sheet["formula_missing_cached_value_count"] = missing_cache_count
        sheet["drawing_relationships"] = drawing_refs
        if missing_cache_count:
            self._warn(
                "formula_cached_values_missing",
                f"Sheet {name!r} has {missing_cache_count} formula cell(s) with no "
                "stored <v> cache.",
                sheet=name,
                details={
                    "formula_count": formula_count,
                    "missing_cached_value_count": missing_cache_count,
                },
            )

    # --------------------------------------------------------------- drawings

    def _index_drawings(
        self,
        package: zipfile.ZipFile,
        relationship_lookup: dict[tuple[str, str], dict[str, Any]],
        parts_by_name: dict[str, dict[str, Any]],
    ) -> None:
        parsed_cache: dict[str, list[dict[str, Any]]] = {}
        for sheet in self.sheet_catalog:
            sheet_name = str(sheet["name"])
            destination = self.drawings_by_sheet.setdefault(sheet_name, [])
            sheet_part = sheet.get("part")

            refs_by_id = {
                item.get("relationship_id"): item
                for item in sheet.get("drawing_relationships", [])
                if item.get("relationship_id")
            }
            # Include drawing relationships even if malformed sheet XML omitted
            # the corresponding element; this is useful forensic evidence.
            if sheet_part:
                for (source, rel_id), relation in relationship_lookup.items():
                    if source != sheet_part:
                        continue
                    if not str(relation.get("type") or "").endswith(
                        _DRAWING_REL_SUFFIXES
                    ):
                        continue
                    refs_by_id.setdefault(
                        rel_id,
                        {
                            "element_type": "relationship_only",
                            "relationship_id": rel_id,
                            "relationship_type": relation.get("type"),
                            "target": relation.get("target"),
                            "part": relation.get("resolved_target"),
                            "external": bool(relation.get("external")),
                        },
                    )

            for rel_id, reference in refs_by_id.items():
                relation = relationship_lookup.get((sheet_part, rel_id)) if sheet_part else None
                drawing_part = reference.get("part")
                base = {
                    "sheet": sheet_name,
                    "sheet_part": sheet_part,
                    "relationship_id": rel_id,
                    "relationship_type": reference.get("relationship_type"),
                    "relationship_target": reference.get("target"),
                    "sheet_element_type": reference.get("element_type"),
                    "drawing_part": drawing_part,
                }
                if reference.get("external") or (relation and relation.get("external")):
                    destination.append(
                        {
                            **base,
                            "type": "external_drawing_relationship",
                            "anchor_type": None,
                            "anchor": None,
                            "name": None,
                            "description": None,
                            "text": None,
                        }
                    )
                    continue
                if not drawing_part or drawing_part not in parts_by_name:
                    self._warn(
                        "drawing_part_missing",
                        f"Drawing relationship {rel_id!r} on sheet {sheet_name!r} "
                        f"does not resolve to a package part.",
                        sheet=sheet_name,
                        details={"drawing_part": drawing_part},
                    )
                    continue
                if drawing_part not in parsed_cache:
                    try:
                        if str(drawing_part).casefold().endswith(".vml"):
                            parsed_cache[drawing_part] = _parse_vml_drawing(
                                package, parts_by_name[drawing_part]["zip_name"]
                            )
                        else:
                            parsed_cache[drawing_part] = _parse_drawingml(
                                package, parts_by_name[drawing_part]["zip_name"]
                            )
                    except (ET.ParseError, KeyError, RuntimeError) as exc:
                        parsed_cache[drawing_part] = []
                        self._warn(
                            "drawing_parse_error",
                            f"Could not parse drawing part {drawing_part!r}.",
                            sheet=sheet_name,
                            details={"error": str(exc)},
                        )
                objects = parsed_cache[drawing_part]
                if not objects:
                    destination.append(
                        {
                            **base,
                            "type": "drawing_part",
                            "anchor_type": None,
                            "anchor": None,
                            "name": None,
                            "description": None,
                            "text": None,
                        }
                    )
                else:
                    for drawing_object in objects:
                        destination.append({**base, **drawing_object})

    # --------------------------------------------------------------- features

    @staticmethod
    def _detect_features(
        parts: list[dict[str, Any]],
    ) -> tuple[dict[str, bool], dict[str, list[str]]]:
        categories: dict[str, list[str]] = {
            "macros": [],
            "queries": [],
            "connections": [],
            "threaded_comments": [],
        }
        for part in parts:
            name = str(part["name"])
            lowered_name = name.casefold()
            lowered_type = str(part.get("content_type") or "").casefold()
            if (
                "vbaproject" in lowered_name
                or "vbaProject".casefold() in lowered_type
                or "macroenabled" in lowered_type
            ):
                categories["macros"].append(name)
            if (
                "/querytables/" in f"/{lowered_name}"
                or "/queries/" in f"/{lowered_name}"
                or "querytable" in lowered_type
                or "query" in lowered_type and "query" in lowered_name
            ):
                categories["queries"].append(name)
            if lowered_name.endswith("/connections.xml") or "connections" in lowered_type:
                categories["connections"].append(name)
            if (
                "/threadedcomments/" in f"/{lowered_name}"
                or "threadedcomment" in lowered_type
            ):
                categories["threaded_comments"].append(name)
        feature_parts = {
            key: sorted(set(value)) for key, value in categories.items()
        }
        return {key: bool(value) for key, value in feature_parts.items()}, feature_parts

    def _add_feature_warnings(
        self, features: dict[str, bool], feature_parts: dict[str, list[str]]
    ) -> None:
        if features["macros"]:
            self._warn(
                "macros_present_not_executed",
                "Macro-related package parts are present; they were checksummed but "
                "never executed.",
                severity="info",
                details={"parts": feature_parts["macros"]},
            )
        if features["queries"] or features["connections"]:
            self._warn(
                "data_connections_not_refreshed",
                "Query/connection package parts are present; no connection was opened "
                "or refreshed.",
                severity="info",
                details={
                    "query_parts": feature_parts["queries"],
                    "connection_parts": feature_parts["connections"],
                },
            )

    def _warn(
        self,
        code: str,
        message: str,
        *,
        severity: str = "warning",
        sheet: str | None = None,
        coordinate: str | None = None,
        details: dict[str, Any] | None = None,
    ) -> None:
        self.warnings.append(
            WarningRecord(
                code=code,
                message=message,
                severity=severity,
                sheet=sheet,
                coordinate=coordinate,
                details=details or {},
            )
        )


# ------------------------------------------------------------------ utilities


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1].rsplit(":", 1)[-1]


def _attribute_by_local_name(element: ET.Element, local_name: str) -> str | None:
    for key, value in element.attrib.items():
        if _local_name(key) == local_name:
            return value
    return None


def _first_child(element: ET.Element, local_name: str) -> ET.Element | None:
    for child in element:
        if _local_name(child.tag) == local_name:
            return child
    return None


def _normalize_member_name(name: str) -> str:
    return name.replace("\\", "/").lstrip("/")


def _validate_member_name(name: str) -> str:
    if not name or "\x00" in name:
        raise OOXMLSecurityError("ZIP package contains an empty or NUL member name.")
    if name.startswith(("/", "\\")) or PureWindowsPath(name).drive:
        raise OOXMLSecurityError(f"ZIP member has an absolute path: {name!r}.")
    normalized = name.replace("\\", "/")
    raw_parts = normalized.split("/")
    if normalized.endswith("/"):
        raw_parts = raw_parts[:-1]
    if not raw_parts or any(part in {"", ".", ".."} for part in raw_parts):
        raise OOXMLSecurityError(f"ZIP member has an unsafe path: {name!r}.")
    return normalized.rstrip("/") if normalized.endswith("/") else normalized


def _compression_ratio(info: zipfile.ZipInfo) -> float:
    if info.file_size == 0:
        return 1.0
    if info.compress_size == 0:
        return math.inf
    return info.file_size / info.compress_size


def _source_part_from_rels_part(rels_part: str) -> str:
    normalized = _normalize_member_name(rels_part)
    if normalized == "_rels/.rels":
        return ""
    path = PurePosixPath(normalized)
    if path.parent.name != "_rels" or not path.name.endswith(".rels"):
        raise OOXMLFormatError(f"Invalid relationships part name: {rels_part!r}.")
    source_name = path.name[: -len(".rels")]
    return str(path.parent.parent / source_name)


def _is_external_target(target: str, target_mode: str | None) -> bool:
    if str(target_mode or "").casefold() == "external":
        return True
    stripped = target.strip()
    if stripped.startswith(("\\\\", "//")):
        return True
    return bool(_EXTERNAL_URI_RE.match(stripped))


def _resolve_relationship_target(source_part: str, target: str) -> str:
    parsed = urlsplit(target)
    target_path = unquote(parsed.path).replace("\\", "/")
    if not target_path:
        raise OOXMLFormatError(f"Relationship target has no package path: {target!r}.")
    if target_path.startswith("/"):
        candidate = posixpath.normpath(target_path.lstrip("/"))
    else:
        base = posixpath.dirname(source_part) if source_part else ""
        candidate = posixpath.normpath(posixpath.join(base, target_path))
    if candidate in {"", ".", ".."} or candidate.startswith("../"):
        raise OOXMLFormatError(
            f"Relationship target escapes the package root: {target!r}."
        )
    return _normalize_member_name(candidate)


def _reject_doctype(stream: BinaryIO, part_name: str) -> None:
    """Reject DTD-bearing XML without retaining a large part in memory."""

    carry = b""
    while chunk := stream.read(64 * 1024):
        sample = (carry + chunk).upper()
        if b"<!DOCTYPE" in sample or b"<!ENTITY" in sample:
            raise OOXMLSecurityError(
                f"DTD/entity declarations are not allowed in OOXML part {part_name!r}."
            )
        carry = sample[-16:]


def _parse_drawingml(
    package: zipfile.ZipFile, zip_name: str
) -> list[dict[str, Any]]:
    with package.open(zip_name, "r") as stream:
        _reject_doctype(stream, zip_name)
    objects: list[dict[str, Any]] = []
    with package.open(zip_name, "r") as stream:
        for _event, element in ET.iterparse(stream, events=("end",)):
            anchor_type = _local_name(element.tag)
            if anchor_type not in _ANCHOR_ELEMENT_NAMES:
                continue
            object_element = next(
                (
                    child
                    for child in element
                    if _local_name(child.tag) in _DRAWING_ELEMENT_NAMES
                ),
                None,
            )
            object_type = _local_name(object_element.tag) if object_element is not None else "unknown"
            non_visual = _first_descendant(element, "cNvPr")
            preset_geometry = _first_descendant(element, "prstGeom")
            text = _drawing_text(element)
            relationship_ids: list[str] = []
            for descendant in element.iter():
                for key, value in descendant.attrib.items():
                    if _local_name(key) in {"embed", "link"}:
                        relationship_ids.append(value)
            record = {
                "type": object_type,
                "anchor_type": anchor_type,
                "anchor": _drawing_anchor(element, anchor_type),
                "object_id": non_visual.attrib.get("id") if non_visual is not None else None,
                "name": non_visual.attrib.get("name") if non_visual is not None else None,
                "description": (
                    non_visual.attrib.get("descr") if non_visual is not None else None
                ),
                "title": non_visual.attrib.get("title") if non_visual is not None else None,
                "hidden": non_visual.attrib.get("hidden") if non_visual is not None else None,
                "shape_type": (
                    preset_geometry.attrib.get("prst")
                    if preset_geometry is not None
                    else None
                ),
                "text": text or None,
                "embedded_relationship_ids": list(dict.fromkeys(relationship_ids)),
            }
            objects.append(record)
            element.clear()
    return objects


def _first_descendant(element: ET.Element, local_name: str) -> ET.Element | None:
    for descendant in element.iter():
        if _local_name(descendant.tag) == local_name:
            return descendant
    return None


def _drawing_text(element: ET.Element) -> str:
    paragraphs: list[str] = []
    for descendant in element.iter():
        if _local_name(descendant.tag) != "p":
            continue
        runs = [
            child.text or ""
            for child in descendant.iter()
            if _local_name(child.tag) == "t"
        ]
        if runs:
            paragraphs.append("".join(runs))
    if paragraphs:
        return "\n".join(paragraphs)
    texts = [
        descendant.text or ""
        for descendant in element.iter()
        if _local_name(descendant.tag) == "t"
    ]
    return "".join(texts)


def _drawing_anchor(element: ET.Element, anchor_type: str) -> dict[str, Any]:
    result: dict[str, Any] = {}
    if anchor_type in {"twoCellAnchor", "oneCellAnchor"}:
        start = _first_child(element, "from")
        if start is not None:
            result["from"] = _marker(start)
    if anchor_type == "twoCellAnchor":
        end = _first_child(element, "to")
        if end is not None:
            result["to"] = _marker(end)
    extent = _first_child(element, "ext")
    if extent is not None:
        result["extent"] = {
            "cx": _integer_or_text(extent.attrib.get("cx")),
            "cy": _integer_or_text(extent.attrib.get("cy")),
        }
    position = _first_child(element, "pos")
    if position is not None:
        result["position"] = {
            "x": _integer_or_text(position.attrib.get("x")),
            "y": _integer_or_text(position.attrib.get("y")),
        }
    return result


def _marker(element: ET.Element) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key in ("col", "colOff", "row", "rowOff"):
        child = _first_child(element, key)
        result[key] = _integer_or_text(child.text if child is not None else None)
    return result


def _integer_or_text(value: str | None) -> int | str | None:
    if value is None:
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return value


def _parse_vml_drawing(
    package: zipfile.ZipFile, zip_name: str
) -> list[dict[str, Any]]:
    with package.open(zip_name, "r") as stream:
        _reject_doctype(stream, zip_name)
    objects: list[dict[str, Any]] = []
    with package.open(zip_name, "r") as stream:
        tree = ET.parse(stream)
    for element in tree.getroot().iter():
        if _local_name(element.tag) not in {"shape", "group"}:
            continue
        text_parts = [
            child.text or ""
            for child in element.iter()
            if child.text and _local_name(child.tag) not in {"Anchor", "Row", "Column"}
        ]
        anchor_element = _first_descendant(element, "Anchor")
        anchor = _vml_anchor(anchor_element.text if anchor_element is not None else None)
        objects.append(
            {
                "type": _local_name(element.tag),
                "anchor_type": "vml" if anchor else None,
                "anchor": anchor,
                "object_id": element.attrib.get("id"),
                "name": element.attrib.get("id"),
                "description": element.attrib.get("alt"),
                "title": element.attrib.get("title"),
                "hidden": None,
                "shape_type": element.attrib.get("type"),
                "text": " ".join(" ".join(text_parts).split()) or None,
                "embedded_relationship_ids": [],
                "style": element.attrib.get("style"),
            }
        )
    return objects


def _vml_anchor(text: str | None) -> dict[str, Any] | None:
    if not text:
        return None
    pieces = [piece.strip() for piece in text.split(",")]
    if len(pieces) != 8:
        return {"raw": text}
    values = [_integer_or_text(piece) for piece in pieces]
    return {
        "from": {
            "col": values[0],
            "colOff": values[1],
            "row": values[2],
            "rowOff": values[3],
        },
        "to": {
            "col": values[4],
            "colOff": values[5],
            "row": values[6],
            "rowOff": values[7],
        },
        "raw": text,
    }
