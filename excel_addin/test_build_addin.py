"""Exercise the Excel-free package builder with real VBA project files."""

from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
from xml.etree import ElementTree
from zipfile import ZipFile

from pyopenvba import ExcelFile

import build_addin


class BuildAddinTests(unittest.TestCase):
    def setUp(self):
        self.directory = TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.engine = self.root / "engine.bas"
        self.wrapper = self.root / "wrapper.bas"
        self.engine_source = (
            'Attribute VB_Name = "StaticRegexSingle"\r\n'
            'Option Private Module\r\n'
            'Public Function Example() As String\r\n'
            '    Example = "\u00e4\u00f6\u00fc"\r\n'
            'End Function\r\n'
        )
        self.wrapper_source = (
            'Attribute VB_Name = "VbaRegexWorksheetFunctions"\r\n'
            'Public Function RxExample() As String\r\n'
            '    RxExample = "\u00c4\u00d6\u00dc"\r\n'
            'End Function\r\n'
            'Public Sub RegisterRxFunctions()\r\n'
            'End Sub\r\n'
        )
        # Cover both Windows-exported modules and PowerShell 7/UTF-8 output.
        self.engine.write_bytes(self.engine_source.encode("utf-8-sig"))
        self.wrapper.write_bytes(self.wrapper_source.encode("cp1252"))

    def build(self, name):
        output = self.root / name
        build_addin.build_addin(self.engine, self.wrapper, output)
        return output

    def test_addin_and_workbook_preserve_source_and_document_bindings(self):
        for extension, expected_type in (
            ("xlam", "application/vnd.ms-excel.addin.macroEnabled.main+xml"),
            ("xlsm", "application/vnd.ms-excel.sheet.macroEnabled.main+xml"),
        ):
            with self.subTest(extension=extension):
                output = self.build(f"Example functions.{extension}")
                with ZipFile(output) as archive:
                    self.assertIsNone(archive.testzip())
                    types = ElementTree.fromstring(archive.read("[Content_Types].xml"))
                    workbook_type = next(
                        part.attrib["ContentType"] for part in types
                        if part.get("PartName") == "/xl/workbook.xml"
                    )
                    self.assertEqual(workbook_type, expected_type)
                    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
                    workbook_xml = ElementTree.fromstring(archive.read("xl/workbook.xml"))
                    self.assertEqual(workbook_xml.find("s:workbookPr", ns).get("codeName"), "ThisWorkbook")
                    relationships = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
                    self.assertTrue(any(
                        rel.get("Type").endswith("/vbaProject") and rel.get("Target") == "vbaProject.bin"
                        for rel in relationships
                    ))
                with ExcelFile(output) as workbook:
                    self.assertEqual(workbook.validate(), [])
                    self.assertEqual(set(workbook.module_names()), {
                        "ThisWorkbook", "Sheet1", "StaticRegexSingle", "VbaRegexWorksheetFunctions"
                    })
                    self.assertEqual(workbook.get_module("StaticRegexSingle"), self.engine_source)
                    self.assertEqual(workbook.get_module("VbaRegexWorksheetFunctions"), self.wrapper_source)
                    events = workbook.get_module("ThisWorkbook")
                    self.assertIn('Attribute VB_PredeclaredId = True', events)
                    self.assertIn('Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"', events)
                    for event in ("Workbook_Open", "Workbook_AddinInstall"):
                        self.assertIn(
                            f"Private Sub {event}()\r\n"
                            "    VbaRegexWorksheetFunctions.RegisterRxFunctions\r\nEnd Sub", events
                        )

    def test_rebuild_replaces_previous_source(self):
        output = self.build("updated.xlam")
        new_source = self.engine_source.replace("\u00e4\u00f6\u00fc", "fresh source")
        self.engine.write_bytes(new_source.encode("utf-8"))
        self.build(output.name)
        with ExcelFile(output) as workbook:
            self.assertEqual(workbook.get_module("StaticRegexSingle"), new_source)

    def test_unrepresentable_source_preserves_existing_output(self):
        output = self.build("existing.xlam")
        previous = output.read_bytes()
        self.engine.write_bytes(self.engine_source.replace("\u00e4", "\ufffd").encode("utf-8"))
        with self.assertRaisesRegex(ValueError, "outside Windows-1252"):
            self.build(output.name)
        self.assertEqual(output.read_bytes(), previous)

    def test_packaging_failure_preserves_existing_output(self):
        output = self.build("existing.xlam")
        previous = output.read_bytes()
        with patch.object(ExcelFile, "save", side_effect=OSError("simulated write failure")):
            with self.assertRaisesRegex(OSError, "simulated write failure"):
                self.build(output.name)
        self.assertEqual(output.read_bytes(), previous)
        self.assertEqual(list(self.root.glob(".addin-*")), [])

    def test_wrong_module_name_is_rejected(self):
        self.engine.write_bytes(self.engine_source.replace("StaticRegexSingle", "WrongName").encode("utf-8"))
        with self.assertRaisesRegex(ValueError, "expected one VB_Name"):
            self.build("wrong.xlam")
        self.assertFalse((self.root / "wrong.xlam").exists())

    def test_non_macro_output_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "extension"):
            self.build("wrong.xlsx")
        self.assertFalse((self.root / "wrong.xlsx").exists())


if __name__ == "__main__":
    unittest.main()
