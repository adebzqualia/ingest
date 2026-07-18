"""Command-line interface for POPS workbook ingestion."""

from __future__ import annotations

import argparse
from pathlib import Path
import sys
import traceback

from . import __version__
from .config import ExtractionConfig
from .extractor import ExtractionError, extract_workbook
from .ooxml import OOXMLIndex
from .output import OutputExistsError, default_output_path
from .selection import (
    SheetSelectionError,
    format_sheet_catalog,
    prompt_for_sheets,
    resolve_sheet_selection,
)


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pops-ingest",
        description=(
            "Extract messy POPS Excel worksheets into a lossless cell layer, "
            "structured tables, review-safe CSV, and a self-contained HTML report."
        ),
    )
    parser.add_argument("workbook", type=Path, help="Reference/template .xlsx or .xlsm file")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    parser.add_argument(
        "--list-sheets",
        action="store_true",
        help="Inventory sheets and exit without extracting",
    )
    parser.add_argument(
        "--sheet",
        action="append",
        default=[],
        metavar="NAME",
        help="Exact sheet name; repeat for multiple sheets",
    )
    parser.add_argument(
        "--sheet-glob",
        action="append",
        default=[],
        metavar="PATTERN",
        help="Case-sensitive sheet-name glob such as 'FTE*'; repeatable",
    )
    parser.add_argument(
        "--sheet-index",
        action="append",
        default=[],
        metavar="1,3-5",
        help="One-based sheet indexes/ranges; repeatable",
    )
    selection = parser.add_mutually_exclusive_group()
    selection.add_argument("--all-sheets", action="store_true", help="Extract every sheet")
    selection.add_argument(
        "--visible-sheets", action="store_true", help="Extract every visible sheet"
    )
    parser.add_argument(
        "--output",
        type=Path,
        help="New output directory (existing directories are never overwritten)",
    )
    parser.add_argument(
        "--table-threshold",
        type=float,
        default=0.50,
        metavar="0..1",
        help="Automatic table acceptance threshold (default: 0.50)",
    )
    parser.add_argument(
        "--uncertain-threshold",
        type=float,
        default=0.35,
        metavar="0..1",
        help="Lowest score still included for human review (default: 0.35)",
    )
    parser.add_argument(
        "--no-cached-values",
        action="store_true",
        help="Skip the second data-only workbook load (formula results will be unavailable)",
    )
    parser.add_argument(
        "--max-file-mb",
        type=int,
        default=300,
        help="Reject larger source files before parsing (default: 300)",
    )
    parser.add_argument(
        "--max-material-cells",
        type=int,
        default=1_500_000,
        help="Per-sheet exported cell safety limit",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Print a traceback for controlled failures",
    )
    return parser


def _config_from_args(args: argparse.Namespace) -> ExtractionConfig:
    if not 0 <= args.uncertain_threshold <= args.table_threshold <= 1:
        raise ValueError(
            "Thresholds must satisfy 0 <= uncertain-threshold <= table-threshold <= 1."
        )
    if args.max_file_mb < 1 or args.max_material_cells < 1:
        raise ValueError("Safety limits must be positive.")
    return ExtractionConfig(
        table_score_threshold=args.table_threshold,
        uncertain_table_threshold=args.uncertain_threshold,
        max_file_bytes=args.max_file_mb * 1024 * 1024,
        max_material_cells_per_sheet=args.max_material_cells,
        include_cached_values=not args.no_cached_values,
    )


def _selection_method(args: argparse.Namespace, prompted: bool) -> str:
    if prompted:
        return "interactive"
    if args.all_sheets:
        return "all_sheets"
    if args.visible_sheets:
        return "visible_sheets"
    return "explicit_cli"


def run(args: argparse.Namespace) -> int:
    config = _config_from_args(args)
    workbook = args.workbook.expanduser().resolve()
    index = OOXMLIndex.open(workbook, config)
    if args.list_sheets:
        print(format_sheet_catalog(index.sheet_catalog))
        return 0
    selected = resolve_sheet_selection(
        index.sheet_catalog,
        sheet_names=args.sheet,
        sheet_globs=args.sheet_glob,
        index_selectors=args.sheet_index,
        all_sheets=args.all_sheets,
        visible_sheets=args.visible_sheets,
    )
    prompted = False
    if not selected:
        selected = prompt_for_sheets(index.sheet_catalog)
        prompted = True
    output = args.output or default_output_path(workbook)
    print("Selected sheets: " + ", ".join(selected))
    print(f"Extracting {workbook.name} …")
    final = extract_workbook(
        workbook,
        selected,
        output,
        config=config,
        ooxml_index=index,
        selection_method=_selection_method(args, prompted),
    )
    print(f"Extraction complete: {final}")
    print(f"Review report:       {final / 'report.html'}")
    print(f"Machine manifest:    {final / 'manifest.json'}")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = _parser()
    args = parser.parse_args(argv)
    try:
        return run(args)
    except (
        ExtractionError,
        OutputExistsError,
        SheetSelectionError,
        ValueError,
        FileNotFoundError,
        PermissionError,
    ) as exc:
        print(f"error: {exc}", file=sys.stderr)
        if args.debug:
            traceback.print_exc()
        return 2
    except KeyboardInterrupt:
        print("\nCancelled; no output bundle was committed.", file=sys.stderr)
        return 130


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
