from __future__ import annotations

import json
import csv
from pathlib import Path
import tempfile
import unittest

from openpyxl import load_workbook
from openpyxl.workbook.defined_name import DefinedName

from pops_ingest.config import ExtractionConfig
from pops_ingest.extractor import extract_workbook
from pops_ingest.ooxml import OOXMLIndex

from tests.fixture_workbook import create_pops_mini


def _json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def _jsonl(path: Path):
    with path.open(encoding="utf-8") as handle:
        return [json.loads(line) for line in handle if line.strip()]


class EndToEndExtractionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tempdir = tempfile.TemporaryDirectory()
        self.root = Path(self.tempdir.name)
        self.source = create_pops_mini(self.root / "pops_mini.xlsx")
        self.config = ExtractionConfig(
            table_score_threshold=0.45,
            uncertain_table_threshold=0.30,
            preview_rows=30,
            preview_columns=30,
        )

    def tearDown(self) -> None:
        self.tempdir.cleanup()

    def test_ooxml_inventory_preserves_sheet_order_and_formula_facts(self) -> None:
        index = OOXMLIndex.open(self.source, self.config)
        self.assertEqual(
            [sheet["name"] for sheet in index.sheet_catalog],
            ["SYNTHESIS -->", "OBS KPI", "MBS (OPEX)", "KPI", "Calc", "Internal"],
        )
        self.assertEqual(index.sheet_catalog[4]["state"], "hidden")
        self.assertEqual(index.sheet_catalog[5]["state"], "veryHidden")
        self.assertIn("E6", index.formula_records["OBS KPI"])
        raw = index.formula_records["OBS KPI"]["E6"]
        self.assertIn("formula_text", raw)
        self.assertIn("cached_value_present", raw)
        self.assertTrue(index.package_manifest["parts"])

    def test_selected_sheet_extraction_is_structured_and_auditable(self) -> None:
        index = OOXMLIndex.open(self.source, self.config)
        destination = extract_workbook(
            self.source,
            ["OBS KPI", "MBS (OPEX)", "KPI"],
            self.root / "run_one",
            config=self.config,
            ooxml_index=index,
            selection_method="test",
        )
        manifest = _json(destination / "manifest.json")
        self.assertEqual(
            manifest["selection"]["selected_sheets"],
            ["OBS KPI", "MBS (OPEX)", "KPI"],
        )
        self.assertIn("Calc", manifest["selection"]["unselected_sheets"])
        self.assertIn("Calc", manifest["selection"]["referenced_but_unselected_sheets"])
        self.assertEqual(len(manifest["sheets"]), 3)
        self.assertTrue((destination / "report.html").is_file())
        schema = _json(destination / "schema.json")
        self.assertEqual(schema["$id"], "urn:pops-ingest:schema:1.0.0")
        self.assertIn("cell", schema["$defs"])
        self.assertTrue((destination / "records_long.jsonl").is_file())
        self.assertTrue((destination / "tables_index.csv").is_file())
        self.assertGreater(sum(s["counts"]["detected_tables"] for s in manifest["sheets"]), 2)

        obs_manifest = next(sheet for sheet in manifest["sheets"] if sheet["name"] == "OBS KPI")
        obs_root = destination / obs_manifest["relative_path"]
        obs_sheet = _json(obs_root / "sheet.json")
        self.assertIn("B3:F3", obs_sheet["merged_ranges"])
        self.assertTrue(any(item["hidden"] for item in obs_sheet["row_dimensions"]))
        self.assertTrue(any(item["hidden"] for item in obs_sheet["column_dimensions"]))
        self.assertTrue(obs_sheet["data_validations"])
        self.assertTrue(obs_sheet["conditional_formats"])
        self.assertEqual(
            [(table["range"], table["title"]) for table in obs_sheet["detected_tables"]],
            [("B3:H11", "OBS KPIs EVOLUTION")],
        )
        self.assertEqual(obs_sheet["detected_tables"][0]["header_rows"], [3, 4])
        row_six = next(
            row for row in obs_sheet["detected_tables"][0]["rows"] if row["row"] == 6
        )
        self.assertEqual(row_six["section_path"], ["TRANSFORMATION KPI'S"])

        cells = _jsonl(obs_root / "cells.jsonl")
        by_coordinate = {cell["coordinate"]: cell for cell in cells}
        formula = by_coordinate["E6"]
        self.assertEqual(formula["formula"], "=D6-C6")
        self.assertEqual(formula["cached_value_status"], "missing")
        self.assertIsNone(formula["literal_value"])
        self.assertIsNone(formula["cached_value"])
        self.assertTrue(formula["formula_signature"])
        self.assertEqual(by_coordinate["C6"]["literal_value"], 0.628)
        self.assertIsNotNone(by_coordinate["B6"]["comment"])

        with (destination / "records_long.csv").open(
            encoding="utf-8", newline=""
        ) as handle:
            csv_rows = list(csv.DictReader(handle))
        formula_csv = next(row for row in csv_rows if row["source_cell"] == "E6")
        self.assertEqual(formula_csv["formula"], "'=D6-C6")

        kpi_manifest = next(sheet for sheet in manifest["sheets"] if sheet["name"] == "KPI")
        kpi_sheet = _json(destination / kpi_manifest["relative_path"] / "sheet.json")
        self.assertEqual(kpi_sheet["explicit_excel_tables"][0]["name"], "KPIReference")
        self.assertEqual(kpi_sheet["explicit_excel_tables"][0]["range"], "A1:F5")
        native_matches = [
            table for table in kpi_sheet["detected_tables"] if table["range"] == "A1:F5"
        ]
        self.assertEqual(len(native_matches), 1)
        self.assertEqual(native_matches[0]["title"], "KPIReference")

        mbs_manifest = next(
            sheet for sheet in manifest["sheets"] if sheet["name"] == "MBS (OPEX)"
        )
        mbs_sheet = _json(destination / mbs_manifest["relative_path"] / "sheet.json")
        self.assertEqual(
            [table["range"] for table in mbs_sheet["detected_tables"]],
            ["B3:O9", "B13:O18"],
        )
        self.assertTrue(
            any(
                annotation["range"] == "Q4"
                and annotation["kind"] in {"instruction", "annotation"}
                for annotation in mbs_sheet["annotations"]
            )
        )
        instruction = next(
            annotation
            for annotation in mbs_sheet["annotations"]
            if annotation["range"] == "Q4"
        )
        self.assertIn(
            instruction["nearest_table_id"],
            {table["table_id"] for table in mbs_sheet["detected_tables"]},
        )

        styles = _json(destination / "styles.json")
        self.assertGreater(len(styles["styles"]), 1)
        report = (destination / "report.html").read_text(encoding="utf-8")
        self.assertIn("POPS workbook extraction", report)
        self.assertIn("OBS KPI", report)
        self.assertIn('href="schema.json"', report)

    def test_structure_hashes_and_table_ids_are_repeatable(self) -> None:
        first = extract_workbook(
            self.source,
            ["OBS KPI", "KPI"],
            self.root / "deterministic_one",
            config=self.config,
            ooxml_index=OOXMLIndex.open(self.source, self.config),
            selection_method="test",
        )
        second = extract_workbook(
            self.source,
            ["OBS KPI", "KPI"],
            self.root / "deterministic_two",
            config=self.config,
            ooxml_index=OOXMLIndex.open(self.source, self.config),
            selection_method="test",
        )
        first_manifest = _json(first / "manifest.json")
        second_manifest = _json(second / "manifest.json")
        self.assertEqual(first_manifest["fingerprints"], second_manifest["fingerprints"])
        first_tables = [
            table["table_id"]
            for sheet in first_manifest["sheets"]
            for table in sheet["tables"]
        ]
        second_tables = [
            table["table_id"]
            for sheet in second_manifest["sheets"]
            for table in sheet["tables"]
        ]
        self.assertEqual(first_tables, second_tables)

    def test_broken_defined_names_warn_and_do_not_abort_extraction(self) -> None:
        workbook = load_workbook(self.source, data_only=False)
        try:
            workbook.defined_names.add(
                DefinedName("BrokenGlobal", attr_text="=#REF!:#REF!")
            )
            workbook["OBS KPI"].defined_names.add(
                DefinedName("BrokenLocal", attr_text="=#REF!:#REF!")
            )
            workbook.save(self.source)
        finally:
            workbook.close()

        destination = extract_workbook(
            self.source,
            ["OBS KPI"],
            self.root / "broken_names",
            config=self.config,
            ooxml_index=OOXMLIndex.open(self.source, self.config),
            selection_method="test",
        )
        manifest = _json(destination / "manifest.json")
        names = {
            (item["name"], item["scope"]): item
            for item in manifest["workbook"]["defined_names"]
        }
        for key in (
            ("BrokenGlobal", "workbook"),
            ("BrokenLocal", "sheet:OBS KPI"),
        ):
            self.assertEqual(names[key]["definition"], "=#REF!:#REF!")
            self.assertEqual(names[key]["status"], "broken")
            self.assertEqual(names[key]["destinations"], [])
            self.assertEqual(names[key]["destination_parse_status"], "failed")
            self.assertEqual(
                names[key]["destination_parse_error"]["type"], "TokenizerError"
            )

        warnings = [
            item
            for item in manifest["warnings"]
            if item["code"] == "DEFINED_NAME_BROKEN_REFERENCE"
        ]
        self.assertCountEqual(
            [item["defined_name"] for item in warnings],
            ["BrokenGlobal", "BrokenLocal"],
        )
        local_warning = next(
            item for item in warnings if item["defined_name"] == "BrokenLocal"
        )
        self.assertEqual(local_warning["sheet"], "OBS KPI")


if __name__ == "__main__":
    unittest.main()
