from __future__ import annotations

from pathlib import Path
import tempfile
import unittest
import zipfile

from pops_ingest.config import ExtractionConfig
from pops_ingest.ooxml import OOXMLFormatError, OOXMLIndex, OOXMLSecurityError


class OoxmlSecurityTests(unittest.TestCase):
    def test_legacy_and_non_zip_inputs_fail_actionably(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            legacy = root / "legacy.xls"
            legacy.write_bytes(b"not an OOXML workbook")
            with self.assertRaisesRegex(OOXMLFormatError, "Legacy .xls"):
                OOXMLIndex.open(legacy, ExtractionConfig())

            fake = root / "fake.xlsx"
            fake.write_bytes(b"not a zip")
            with self.assertRaisesRegex(OOXMLFormatError, "not a valid ZIP"):
                OOXMLIndex.open(fake, ExtractionConfig())

    def test_zip_path_traversal_is_rejected_before_xml_parsing(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            workbook = Path(directory) / "malicious.xlsx"
            with zipfile.ZipFile(workbook, "w") as archive:
                archive.writestr("../outside.xml", "<x/>")
            with self.assertRaises(OOXMLSecurityError):
                OOXMLIndex.open(workbook, ExtractionConfig())


if __name__ == "__main__":
    unittest.main()

