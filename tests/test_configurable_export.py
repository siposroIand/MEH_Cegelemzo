from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import pandas as pd
from openpyxl import load_workbook

from core.constants import PRODUCT_GROUPS_JSON
from core.utils import split_address_parts
from services.aggregation_service import aggregate_data
from services.category_service import CategoryService
from services.configurable_export_service import (
    EXCEL_SECTIONS,
    UNKNOWN_COUNTY,
    ExcelExportOptions,
    available_counties,
    export_configured_excel,
    write_configured_excel,
)
from services.filter_service import Filters, filter_frame
from tests.fixtures import dataset


class ConfigurableExportTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source, self.raw, self.missing = dataset()
        self.categories = CategoryService(PRODUCT_GROUPS_JSON)

    def export(self, options, filters=Filters(), path=None):
        master = aggregate_data(self.raw, self.source, self.categories).master
        path = path or (self.root if options.separate_counties else self.root / "export.xlsx")
        return Path(export_configured_excel(
            path, self.source, self.raw, self.missing, master, options, filters,
            self.categories.all_groups(), "Teszt szűrés",
        ))

    def test_selected_sheets_filters_and_matrix_values(self):
        filters = Filters(years=(2025,), code="170405", types=("HULL_ATADAS",))
        book = load_workbook(self.export(ExcelExportOptions(sections=("matrix", "raw", "detail")), filters))
        self.assertEqual(set(book.sheetnames), {
            "Export beállításai", "Szűrés", "Cég–Év–HAK mátrix", "Nyers adatok (kg)", "Részletes",
        })
        expected = filter_frame(self.raw, filters, self.categories)
        raw_sheet = book["Nyers adatok (kg)"]
        self.assertEqual(raw_sheet.max_row, len(expected) + 1)
        self.assertEqual(raw_sheet["B2"].value, "001234567")
        self.assertEqual(raw_sheet["B2"].data_type, "s")
        matrix = book["Cég–Év–HAK mátrix"]
        self.assertEqual(matrix["B1"].value, 2025)
        self.assertEqual(matrix["B2"].value, "170405")
        self.assertAlmostEqual(matrix["B4"].value, expected.MENNYISEG.sum() / 1000)
        self.assertEqual(matrix["B4"].fill.fgColor.rgb, "00CCF0E7")
        self.assertIn("Megye", [cell.value for cell in book["Részletes"][1]])
        self.assertNotIn("Vármegye", [cell.value for cell in book["Részletes"][1]])
        book.close()

    def test_county_files_are_disjoint_and_keep_unknown_rows(self):
        path = self.export(ExcelExportOptions(sections=tuple(EXCEL_SECTIONS), separate_counties=True))
        files = sorted(path.glob("*.xlsx"))
        self.assertEqual(len(files), 3)
        raw_total = 0
        missing_total = 0
        for file in files:
            book = load_workbook(file)
            raw_total += book["Nyers adatok (kg)"].max_row - 1
            missing_total += book["Kimaradt lekérések"].max_row - 1
            if UNKNOWN_COUNTY in file.stem:
                self.assertEqual(book["Forráscégek"].max_row, 1)
                self.assertEqual(book["Kimaradt lekérések"]["A2"].value, "Kimaradt cég")
            else:
                names = {row[0] for row in list(book["Nyers adatok (kg)"].values)[1:]}
                self.assertEqual(len(names), 1)
                source_names = {row[0] for row in list(book["Forráscégek"].values)[1:]}
                self.assertEqual(names, source_names)
            book.close()
        self.assertEqual(raw_total, len(self.raw))
        self.assertEqual(missing_total, len(self.missing))
        self.assertFalse(list(self.root.glob(".meh-excel-*")))

    def test_all_data_scope_and_explicit_county_selection(self):
        options = ExcelExportOptions(
            sections=("raw", "source", "missing"), counties=("Pest",), use_filters=False,
        )
        book = load_workbook(self.export(options, Filters(companies=("nincs ilyen cég",), years=(1900,))))
        self.assertEqual(book["Nyers adatok (kg)"].max_row, 61)
        self.assertEqual(book["Forráscégek"].max_row, 2)
        self.assertEqual(book["Kimaradt lekérések"].max_row, 1)
        self.assertIn("Teljes adatállomány", book["Export beállításai"]["B2"].value)
        book.close()

    def test_source_and_missing_use_only_company_and_county_filters(self):
        options = ExcelExportOptions(sections=("source", "missing", "raw"))
        book = load_workbook(self.export(options, Filters(companies=("Minta Fém Kft.",), years=(1900,))))
        self.assertEqual(book["Forráscégek"].max_row, 2)
        self.assertEqual(book["Nyers adatok (kg)"].max_row, 1)
        self.assertEqual(book["Kimaradt lekérések"].max_row, 1)
        book.close()

    def test_categories_match_existing_filter_service(self):
        group = self.categories.resolve_groups("170405")[0]
        filters = Filters(categories=(group,), years=(2025,))
        expected = filter_frame(self.raw, filters, self.categories)
        book = load_workbook(self.export(ExcelExportOptions(sections=("raw",)), filters))
        self.assertEqual(book["Nyers adatok (kg)"].max_row, len(expected) + 1)
        book.close()

    def test_same_company_name_different_identifiers_not_mixed(self):
        self.source["CEGNEV"] = "Azonos név"
        self.raw["CEGNEV"] = "Azonos név"
        path = self.export(ExcelExportOptions(sections=("raw",), separate_counties=True))
        self.assertEqual(len(list(path.glob("*.xlsx"))), 2)
        for file in path.glob("*.xlsx"):
            book = load_workbook(file)
            self.assertEqual(len({row[1] for row in list(book["Nyers adatok (kg)"].values)[1:]}), 1)
            book.close()

    def test_missing_source_and_unknown_county(self):
        self.source = pd.DataFrame()
        self.assertEqual(available_counties(self.source, self.raw, self.missing), [UNKNOWN_COUNTY])
        path = self.export(ExcelExportOptions(sections=("raw", "missing"), separate_counties=True))
        self.assertEqual(len(list(path.glob("*.xlsx"))), 1)

    def test_conflicting_counties_rejected(self):
        duplicate = self.source.iloc[[0]].copy()
        duplicate["SZEKHELY"] = "Pest, 2000, Szentendre, Másik cím"
        self.source = pd.concat([self.source, duplicate], ignore_index=True)
        with self.assertRaisesRegex(ValueError, "Eltérő megyeadatok"):
            self.export(ExcelExportOptions())
        self.raw = pd.DataFrame()
        self.missing = pd.DataFrame()
        with self.assertRaisesRegex(ValueError, "Eltérő megyeadatok"):
            self.export(ExcelExportOptions(sections=("source",)))
        self.assertEqual(list(self.root.iterdir()), [])

    def test_unsafe_county_filenames_and_collisions_are_safe(self):
        self.source["SZEKHELY"] = ["../A:B, 1000, A, Cím", "..\\a?b, 1000, B, Cím"]
        path = self.export(ExcelExportOptions(sections=("raw",), separate_counties=True))
        files = list(path.glob("*.xlsx"))
        self.assertEqual(len(files), 2)
        for file in files:
            self.assertNotRegex(file.name, r'[<>:"/\\|?*]')
            self.assertEqual(file.parent, path)

    def test_batch_failure_publishes_nothing_and_does_not_replace_previous_export(self):
        options = ExcelExportOptions(sections=("raw",), separate_counties=True)
        existing = self.export(options)
        contents = {file: file.read_bytes() for file in existing.glob("*.xlsx")}
        calls = 0

        def fail_second(*args):
            nonlocal calls
            calls += 1
            if calls == 2:
                raise OSError("Nincs elég hely")
            return write_configured_excel(*args)

        with patch("services.configurable_export_service.write_configured_excel", side_effect=fail_second):
            with self.assertRaisesRegex(OSError, "Nincs elég hely"):
                self.export(options)
        self.assertEqual(list(self.root.iterdir()), [existing])
        for file, data in contents.items():
            self.assertEqual(file.read_bytes(), data)

    def test_validation_preserves_existing_workbook(self):
        path = self.root / "export.xlsx"
        path.write_bytes(b"existing")
        for options in (
            ExcelExportOptions(sections=()),
            ExcelExportOptions(sections=("invalid",)),
            ExcelExportOptions(sections=("raw", "raw")),
            ExcelExportOptions(counties=("Nincs ilyen megye",)),
        ):
            with self.assertRaises(ValueError):
                self.export(options)
            self.assertEqual(path.read_bytes(), b"existing")
        self.source.loc[0, "SRC_EMAIL"] = "x" * 32768
        with self.assertRaisesRegex(ValueError, "szövegkorlát"):
            self.export(ExcelExportOptions(sections=("source",)))
        self.assertEqual(path.read_bytes(), b"existing")
        self.assertEqual(list(self.root.iterdir()), [path])

    def test_formula_text_is_literal_and_county_label_is_megye(self):
        self.source.loc[0, "SRC_EMAIL"] = "=1+1"
        book = load_workbook(self.export(ExcelExportOptions(sections=("source",))))
        cells = [cell for row in book["Forráscégek"] for cell in row if cell.value == "=1+1"]
        self.assertEqual(len(cells), 1)
        self.assertEqual(cells[0].data_type, "s")
        self.assertEqual(split_address_parts("Pest, 2000, Szentendre, Cím")["Megye"], "Pest")
        self.assertEqual(split_address_parts(None)["Megye"], "")
        book.close()
