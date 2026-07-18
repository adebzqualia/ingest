from __future__ import annotations

from contextlib import redirect_stderr, redirect_stdout
from io import StringIO
from pathlib import Path
import tempfile
import unittest

from pops_ingest.cli import main
from tests.fixture_workbook import create_pops_mini


class CliTests(unittest.TestCase):
    def test_list_and_noninteractive_extract_commands(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = create_pops_mini(root / "mini.xlsx")
            stdout = StringIO()
            with redirect_stdout(stdout):
                result = main([str(source), "--list-sheets"])
            self.assertEqual(result, 0)
            self.assertIn("OBS KPI", stdout.getvalue())
            self.assertIn("veryHidden", stdout.getvalue())

            output = root / "selected"
            stdout = StringIO()
            with redirect_stdout(stdout):
                result = main(
                    [
                        str(source),
                        "--sheet",
                        "OBS KPI",
                        "--output",
                        str(output),
                        "--table-threshold",
                        "0.45",
                        "--uncertain-threshold",
                        "0.30",
                    ]
                )
            self.assertEqual(result, 0)
            self.assertTrue((output / "report.html").is_file())
            self.assertIn("Extraction complete", stdout.getvalue())

    def test_missing_exact_sheet_fails_without_output(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = create_pops_mini(root / "mini.xlsx")
            output = root / "must_not_exist"
            stderr = StringIO()
            with redirect_stderr(stderr):
                result = main(
                    [
                        str(source),
                        "--sheet",
                        "Missing sheet",
                        "--output",
                        str(output),
                    ]
                )
            self.assertEqual(result, 2)
            self.assertFalse(output.exists())
            self.assertIn("no sheet named", stderr.getvalue())


if __name__ == "__main__":
    unittest.main()

