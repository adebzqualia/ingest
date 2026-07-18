from __future__ import annotations

import unittest

from pops_ingest.selection import SheetSelectionError, resolve_sheet_selection


CATALOG = [
    {"name": "SYNTHESIS -->", "state": "visible"},
    {"name": "OBS KPI", "state": "visible"},
    {"name": "Résumé, FR", "state": "visible"},
    {"name": "Calc", "state": "hidden"},
    {"name": "Internal", "state": "veryHidden"},
]


class SheetSelectionTests(unittest.TestCase):
    def test_exact_names_preserve_workbook_order(self) -> None:
        selected = resolve_sheet_selection(
            CATALOG, sheet_names=["Résumé, FR", "SYNTHESIS -->"]
        )
        self.assertEqual(selected, ["SYNTHESIS -->", "Résumé, FR"])

    def test_hidden_sheet_can_be_explicitly_selected(self) -> None:
        self.assertEqual(
            resolve_sheet_selection(CATALOG, sheet_names=["Calc"]), ["Calc"]
        )

    def test_visible_filter_excludes_hidden_states(self) -> None:
        self.assertEqual(
            resolve_sheet_selection(CATALOG, visible_sheets=True),
            ["SYNTHESIS -->", "OBS KPI", "Résumé, FR"],
        )

    def test_index_ranges(self) -> None:
        self.assertEqual(
            resolve_sheet_selection(CATALOG, index_selectors=["2,4-5"]),
            ["OBS KPI", "Calc", "Internal"],
        )

    def test_glob(self) -> None:
        self.assertEqual(
            resolve_sheet_selection(CATALOG, sheet_globs=["*KPI"]), ["OBS KPI"]
        )

    def test_unknown_and_duplicate_fail(self) -> None:
        with self.assertRaises(SheetSelectionError):
            resolve_sheet_selection(CATALOG, sheet_names=["Missing"])
        with self.assertRaises(SheetSelectionError):
            resolve_sheet_selection(
                CATALOG, sheet_names=["OBS KPI"], sheet_globs=["OBS*"]
            )

    def test_descending_or_out_of_range_index_fails(self) -> None:
        for selector in ("4-2", "0", "6", "abc"):
            with self.subTest(selector=selector), self.assertRaises(SheetSelectionError):
                resolve_sheet_selection(CATALOG, index_selectors=[selector])


if __name__ == "__main__":
    unittest.main()

