# POPS Excel ingestion — phase 1

This project ingests a POPS reference/template workbook without modifying it, lets the user choose sheets **before** full extraction, detects messy multi-section tables, and exports two complementary representations:

- a coordinate-faithful audit layer for later template-vs-country anomaly detection;
- interpreted, analysis-ready table records with source-cell provenance.

It is designed for the layouts in the supplied examples: merged multi-row scenario headers, budget/forecast/plan columns, right-side variance blocks, vertical hierarchy labels, stacked sections, side-by-side OPEX/FTE panels, formulas, hidden/grouped rows and columns, native Excel Tables, intentional blank input grids, and notes around the tables.

## Quick start

Python 3.11 or newer is required. Docker is not used.

```powershell
cd C:\Users\adnan\Documents\POPS1
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e .
```

Run one command and choose sheets from the numbered prompt:

```powershell
pops-ingest "C:\path\to\POPS_reference.xlsx"
```

The equivalent module command is:

```powershell
python -m pops_ingest "C:\path\to\POPS_reference.xlsx"
```

List the sheet inventory without extracting:

```powershell
pops-ingest "C:\path\to\POPS_reference.xlsx" --list-sheets
```

Choose exact sheets non-interactively:

```powershell
pops-ingest "C:\path\to\POPS_reference.xlsx" `
  --sheet "ID Card" `
  --sheet "OBS KPI" `
  --sheet "MBS (OPEX)" `
  --output ".\output\reference_ingestion"
```

Other selection modes are `--sheet-index "2,4-7"`, `--sheet-glob "FTE*"`, `--visible-sheets`, and `--all-sheets`. Hidden and `veryHidden` sheets are shown in the inventory and may be selected explicitly. In a non-interactive process, omitting sheet selection is an error rather than silently processing everything.

## Result bundle

An extraction succeeds atomically: files are first written to a private temporary directory and the final directory appears only when the complete run succeeds. Existing output directories are never overwritten.

```text
reference_ingestion/
├── report.html                 self-contained human review UI
├── manifest.json               workbook inventory, selection, hashes, warnings
├── schema.json                 JSON Schema plus JSONL item definitions
├── styles.json                 deduplicated semantic style definitions
├── tables_index.csv            one row per detected raw region
├── clean_tables_index.csv      one row per usable clean table
├── clean_records.jsonl         structured clean business rows
├── records_long.jsonl          authoritative tidy observations
├── records_long.csv            review-safe secondary export
└── sheets/
    └── 03_obs-kpi/
        ├── sheet.json          sheet structure and controls
        ├── cells.jsonl         sparse/lossless cells plus detected table grids
        ├── unassigned_cells.jsonl
        └── tables/
            └── 01_tbl_.../
                ├── table.json
                ├── clean_table.json       clean schema + cell provenance
                ├── clean_table.csv        readable values/formula fallback
                ├── clean_formulas.csv     clean shape with formulas visible
                ├── records_long.jsonl
                ├── records_long.csv
                └── values_wide.csv
```

Open `report.html` directly in a browser. It opens the first usable clean table, with presentation banners, notes, blank spacer rows, and body-empty columns removed. Multi-row Excel headers are flattened into readable business column names. A Raw toggle keeps the original detected rectangle available for audit, alongside value, exact-formula, and cached-result modes.

Post-processing never deletes evidence from the extraction bundle. Every removal decision is listed in `clean_table.json`; false regions with no business rows are marked `empty` and hidden from the report's default navigation. This separation gives downstream analysis clean tables while preserving the coordinate-faithful layer needed for later country-template anomaly detection.

JSON/JSONL is authoritative. CSV formula-like text is prefixed with an apostrophe to prevent spreadsheet-formula injection if a reviewer opens it in Excel; exact formula text remains unchanged in JSONL.

## What is preserved

The workbook manifest inventories every sheet, including unselected sheets. Selected sheets receive full extraction of:

- literal values and types;
- exact resolved formulas, raw OOXML formula attributes, relative signatures, and best-effort dependencies;
- cached formula results as a **separate, explicitly qualified** field;
- merge ranges/anchors, borders, fills, fonts, alignment, indentation, number formats, and protection;
- hidden, grouped, collapsed rows/columns and their sizes;
- comments, hyperlinks, validations, conditional-format ranges, native Excel Tables, filters, names, print/freeze settings;
- drawings and text-box content where OOXML exposes it;
- source coordinate and stable `cell_id` for every semantic observation.

Blank, numeric zero, dash, `ns`, placeholder, Excel error, and missing formula cache are kept as different states. Excel's visible `#####` overflow is not stored in the workbook and therefore never replaces the underlying value.

The source workbook is never saved through `openpyxl`; macros are never executed, external links are never followed, queries are never refreshed, and formulas are never recalculated.

## How table detection works

The extractor does not call `pandas.read_excel()` and hope that one rectangle starts at row 1. It uses an ensemble:

1. authoritative native Excel Table ranges;
2. static named ranges that are table-like;
3. sparse border/style/content components, including bordered blank input cells;
4. bounded gap bridging for visual scenario separators;
5. merged titles and multi-period header evidence;
6. row/column type regularity and numeric/formula series;
7. note/prose penalties and overlap deduplication.

Inside each accepted range it infers multi-row header paths, column groups, scenario/year/comparison dimensions, row-label hierarchy, section bands, details, totals, subtotals, placeholders, and notes. Raw labels and coordinates always remain available beside normalized fields. Low-confidence regions above `--uncertain-threshold` are retained and highlighted for review instead of being silently discarded.

The default thresholds are intentionally configurable because the real POPS template—not screenshots—is the final calibration source:

```powershell
pops-ingest template.xlsx --sheet "OBS KPI" `
  --table-threshold 0.55 --uncertain-threshold 0.32
```

## Formula cache semantics

`openpyxl` does not calculate Excel formulas. A second `data_only=True` view reads only the last result saved by Excel, which may be stale or absent. The output therefore uses:

- `formula`: exact formula expression;
- `cached_value`: last saved Excel result, if present;
- `cached_value_status`: `present`, `missing`, or `not_requested`;
- `value`: the cache only when its status is `present`.

Missing caches generate `FORMULA_CACHE_MISSING`; the extractor never invents a zero or computes a guessed result. Use `--no-cached-values` when memory matters and cached results are not needed.

## Supported inputs and safety

Supported: `.xlsx`, `.xlsm`, `.xltx`, `.xltm` OOXML packages. Macro parts are inventoried but not run.

Legacy `.xls`, binary `.xlsb`, encrypted/password-protected, malformed, and non-ZIP files fail with an actionable error. Before loading, the tool checks file size, ZIP entry count, total expanded size, and extreme compression ratios. Per-sheet material-cell and table-area limits prevent an inflated worksheet dimension or full-column rule from forcing a dense million-row scan.

Stale Excel defined names such as `=#REF!:#REF!` are treated differently from a malformed workbook package: their raw definitions and parse diagnostics are retained in `manifest.json`, a warning is emitted, and extraction continues. Such names are never used as table ranges.

## Verification

The test suite builds POPS-like workbooks containing merged hierarchy headers, stacked/side-by-side tables, blank bordered grids, formulas without caches, native Excel Tables, hidden/grouped dimensions, validations, comments, hyperlinks, named ranges, notes, and formula-injection strings.

```powershell
python -m unittest discover -s tests -v
```

The end-to-end checks assert sheet selection, formula separation, merge/style fidelity, stable IDs, source provenance, atomic outputs, HTML generation, and deterministic structural hashes.

## Scope boundary

This is phase 1 only: reference ingestion, structure capture, table detection, and standardization. It does not yet compare country returns, flag changed rows/formulas, validate business totals, repair workbooks, consolidate countries, or calculate global KPIs. The independent layout/schema/formula/style/dependency/content fingerprints and source-cell records are deliberately produced now so those later controls can be built without redesigning ingestion.
