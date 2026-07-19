# Extraction schema

Schema version `1.0.0` is intentionally split into raw evidence and interpreted observations.

## `manifest.json`

The manifest records source provenance, selected/unselected sheets, extraction configuration, package inventory, workbook metadata, names, external relationships, per-sheet artifact paths, warnings, and two top-level hashes:

- `structure_sha256`: layout + formula + style fingerprints;
- `content_sha256`: literal/cached content only.

Country inputs are expected to change content while keeping the reference structure stable.

Defined names retain their raw `definition`, scope, parsed destinations, and destination parsing status. Broken or malformed names (for example, `=#REF!:#REF!`) are preserved with `status: "broken"` or `status: "invalid_unparsed"`; they generate a warning and are excluded from table detection instead of aborting extraction.

## `cells.jsonl`

Each line is one coordinate-aware cell record. Important fields are:

```json
{
  "cell_id": "cell_…",
  "sheet": "OBS KPI",
  "coordinate": "N18",
  "row": 18,
  "column": 14,
  "literal_value": null,
  "literal_value_type": "blank",
  "formula": "='CVC'!N27",
  "formula_signature": "='CVC'!N27",
  "cached_value": -8.9,
  "cached_value_status": "present",
  "number_format": "0.0",
  "style_ref": "style:…",
  "merge_ref": null,
  "visibility": {"row_hidden": false, "column_hidden": true},
  "source": {"sheet_xml": "xl/worksheets/sheet7.xml", "coordinate": "N18"}
}
```

Formula caches are not claimed to be recalculated or current.

`formula_tokenization_status` and `formula_tokenization_error` distinguish a preserved formula that could not be parsed for dependencies or a relative signature. A malformed/error expression never prevents the remaining cells or sheets from being extracted.

## `table.json`

A logical table contains its exact physical range, stable ID, template key, detection methods/confidence, headers, row roles, sections, groups, and fingerprints. Header tokens keep source coordinates.

Column descriptors add optional interpretations such as `scenario`, `year`, `comparison_kind`, `unit`, and `measure`; raw `header_path` is authoritative when an interpretation is uncertain.

## `records_long.jsonl`

Each record is an interpreted table observation with direct cell provenance:

```json
{
  "table_id": "tbl_…",
  "cell_id": "cell_…",
  "section_path": ["COST & PRODUCTIVITY"],
  "row_label_path": ["Customer Services VC OPEX"],
  "column_header_path": ["PLAN", "2028"],
  "scenario": "plan",
  "year": 2028,
  "value": -8.9,
  "value_state": "number",
  "formula": "='CVC'!N27",
  "source_sheet": "CUSTOMER SERVICES",
  "source_cell": "N28"
}
```

`value_state` distinguishes at least `blank`, `formula_cache_missing`, `zero`, `number`, `dash`, `ns`, `placeholder`, `excel_error`, `boolean`, and `text`.

## Clean table artifacts

Every detected region is post-processed into `clean_table.json`. The raw range remains unchanged for audit and future anomaly detection, while the clean artifact:

- removes presentation banners, notes, instructions, footnotes, header rows, section-only rows, and empty spacer rows;
- removes columns that are empty throughout the retained business rows;
- flattens contributing merged/multi-row headers into unique, readable column labels;
- carries section/group context onto each retained row instead of keeping decorative separator rows;
- retains source coordinates, typed literal values, exact formulas, cached values, and number formats for every clean cell.

Regions with no usable business rows or columns receive `status: "empty"` and are excluded from the default report navigation. They remain available in the raw table and cell artifacts.

For `status: "ready"`, each table directory also contains:

- `clean_table.csv`: readable values, using the exact formula when no cached result exists;
- `clean_formulas.csv`: the same clean shape with formulas visible where present;
- `clean_table.json`: authoritative clean rows, columns, provenance, and removal decisions.

At bundle root, `clean_tables_index.csv` lists only usable clean tables and `clean_records.jsonl` provides one structured object per clean row.

## Styles

`styles.json` is keyed by a hash of semantic font/fill/border/alignment/protection/number-format properties. Workbook-local `style_id` is retained on cells for provenance but is not used as the comparison identity, because style IDs can change after a harmless Excel re-save.
