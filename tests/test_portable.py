import io
import json
from pathlib import Path
import pickle
import tempfile
import unittest
from unittest.mock import patch
import zipfile

import pandas as pd
from openpyxl import load_workbook

from services.export_service import write_dataset_excel
from services.ktj_service import KtjService
from services.package_service import MAX_PACKAGE_BYTES, read_package, read_packages, write_package
from services.persistence_service import PersistenceService
from services.report_service import csv_bytes, dataset_bytes, read_dataset
from tests.fixtures import dataset


def legacy_package(tables, state=None, metadata=None, protocol=4):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, frame in zip(("source", "raw", "missing"), tables):
            archive.writestr(name + ".pkl", pickle.dumps(frame, protocol=protocol))
        archive.writestr("state.json", json.dumps(state or {}))
        archive.writestr("ktj_meta.json", json.dumps(metadata or {}))
    return buffer.getvalue()


class ForbiddenObject:
    def __reduce__(self):
        return eval, ("1 + 1",)


class PortableTests(unittest.TestCase):
    def setUp(self):
        self.tables = dataset()
        self.state = {
            "selected_years": [2025], "selected_companies": ["Minta Fém Kft."],
            "selected_types": ["HULL_ATADAS"], "filter_code": "170405", "current_tab_index": 1,
        }
        self.metadata = {"1000000": {"MEGNEVEZES": "Árvíztűrő telephely", "IRSZAM": "0123"}}

    def assert_tables_equal(self, actual):
        for expected, restored in zip(self.tables, actual):
            pd.testing.assert_frame_equal(
                expected.fillna(""), restored.fillna(""), check_dtype=False,
            )

    def test_json_and_package_round_trip_preserves_data_state_and_metadata(self):
        self.tables[1].loc[0, "MENNYISEG"] = 123.12345678901234
        self.assert_tables_equal(read_dataset(dataset_bytes(*self.tables)))
        buffer = io.BytesIO()
        write_package(buffer, *self.tables, self.state, self.metadata)
        with zipfile.ZipFile(buffer) as archive:
            self.assertEqual(set(archive.namelist()), {"dataset.json", "state.json", "ktj_meta.json"})
        restored = read_package(buffer.getvalue())
        self.assert_tables_equal(restored[:3])
        self.assertEqual(restored[1].iloc[0].MENNYISEG, self.tables[1].iloc[0].MENNYISEG)
        self.assertEqual(restored[3:], (self.state, self.metadata))
        self.assertEqual(restored[0].iloc[0].KUJ, "001234567")

    def test_original_pickle_layout_is_readable_only_when_explicitly_enabled(self):
        for protocol in (4, 5):
            with self.subTest(protocol=protocol):
                data = legacy_package(self.tables, self.state, self.metadata, protocol)
                with self.assertRaisesRegex(ValueError, "asztali"):
                    read_package(data)
                result = read_package(data, allow_legacy=True)
                self.assert_tables_equal(result[:3])
                self.assertEqual(result[3:], (self.state, self.metadata))
                modern = io.BytesIO()
                write_package(modern, *result)
                self.assert_tables_equal(read_package(modern.getvalue())[:3])

    def test_partial_datasets_and_empty_package(self):
        source, _raw, missing = self.tables
        empty = pd.DataFrame()
        for tables in ((source, empty, empty), (empty, empty, missing)):
            buffer = io.BytesIO()
            write_package(buffer, *tables)
            result = read_package(buffer.getvalue())
            self.assertEqual([len(frame) for frame in result[:3]], [len(frame) for frame in tables])
            workbook = io.BytesIO()
            write_dataset_excel(workbook, *tables)
            self.assertEqual(len(load_workbook(workbook).worksheets), 3)
        with self.assertRaisesRegex(ValueError, "Nincs exportálható"):
            write_package(io.BytesIO(), empty, empty, empty)

    def test_legacy_reader_rejects_executable_globals_and_non_tables(self):
        for invalid in (ForbiddenObject(), {"not": "a table"}):
            data = legacy_package((invalid, self.tables[1], self.tables[2]))
            with self.assertRaises(ValueError):
                read_package(data, allow_legacy=True)

    def test_invalid_packages_and_expansion_limit(self):
        with self.assertRaisesRegex(ValueError, "Sérült"):
            read_package(b"not a zip")
        for entries in (
            {"raw.pkl": b"missing other tables"},
            {"../dataset.json": b"{}"},
            {"dataset.json": b"{}", "raw.pkl": b"mixed"},
            {"dataset.json": b'{"format":"unknown"}'},
            {"dataset.json": dataset_bytes(*self.tables), "state.json": b'{"selected_years":["bad"]}'},
            {"dataset.json": dataset_bytes(*self.tables), "ktj_meta.json": b'{"001":[]}'},
        ):
            buffer = io.BytesIO()
            with zipfile.ZipFile(buffer, "w") as archive:
                for name, content in entries.items():
                    archive.writestr(name, content)
            with self.subTest(entries=list(entries)), self.assertRaises(ValueError):
                read_package(buffer.getvalue(), allow_legacy=True)
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("dataset.json", " " * 10000)
        with patch("services.package_service.MAX_PACKAGE_BYTES", 1000):
            with self.assertRaisesRegex(ValueError, "kicsomagolt"):
                read_package(buffer.getvalue())
            with self.assertRaises(ValueError):
                write_package(io.BytesIO(), *self.tables)

    def test_json_rejects_non_finite_quantities_and_row_limit(self):
        payload = json.loads(dataset_bytes(*self.tables))
        payload["raw"][0]["MENNYISEG"] = "inf"
        with self.assertRaisesRegex(ValueError, "Nem véges"):
            read_dataset(json.dumps(payload))
        self.tables[1].loc[0, "MENNYISEG"] = float("inf")
        with self.assertRaises(ValueError):
            dataset_bytes(*self.tables)
        with patch("services.report_service.MAX_ROWS", 1), self.assertRaises(ValueError):
            read_dataset(dataset_bytes(*dataset()))

    def test_desktop_reads_packages_larger_than_100_mb(self):
        payload = dataset_bytes(*self.tables)
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "large.okirpkg"
            for compression in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED):
                with self.subTest(compression=compression):
                    with zipfile.ZipFile(path, "w", compression=compression) as archive:
                        with archive.open("dataset.json", "w") as entry:
                            for _ in range(100):
                                entry.write(b" " * (1024 * 1024))
                            entry.write(payload)
                    with zipfile.ZipFile(path) as archive:
                        self.assertGreater(archive.getinfo("dataset.json").file_size, MAX_PACKAGE_BYTES)
                    if compression == zipfile.ZIP_STORED:
                        self.assertGreater(path.stat().st_size, MAX_PACKAGE_BYTES)
                    with self.assertRaisesRegex(ValueError, "100 MB"):
                        read_package(path)
                    self.assert_tables_equal(read_package(path, enforce_limits=False)[:3])

    def test_desktop_package_round_trip_above_250000_rows(self):
        source = pd.DataFrame({"CEGNEV": ["Cég"] * 250_001, "KUJ": ["001"] * 250_001})
        buffer = io.BytesIO()
        with self.assertRaisesRegex(ValueError, "túl nagy"):
            write_package(buffer, source, pd.DataFrame(), pd.DataFrame())
        write_package(buffer, source, pd.DataFrame(), pd.DataFrame(), enforce_limits=False)
        with self.assertRaisesRegex(ValueError, "túl nagy"):
            read_package(buffer.getvalue())
        restored = read_package(buffer.getvalue(), enforce_limits=False)
        pd.testing.assert_frame_equal(source, restored[0])

    def test_unlimited_packages_still_validate_contents(self):
        with self.assertRaisesRegex(ValueError, "Nem támogatott típus"):
            read_package(legacy_package((ForbiddenObject(), *self.tables[1:])),
                         allow_legacy=True, enforce_limits=False)
        invalid = self.tables[1].copy()
        invalid.loc[0, "EV"] = 1800
        with self.assertRaisesRegex(ValueError, "Érvénytelen év"):
            write_package(io.BytesIO(), self.tables[0], invalid, self.tables[2], enforce_limits=False)
        with self.assertRaisesRegex(ValueError, "legalább egy"):
            read_packages([], enforce_limits=False)

    def test_csv_excel_compatibility_and_formula_protection(self):
        frame = pd.DataFrame({
            "Név": ["Árvíztűrő; cég", "=1+1", "  @SUM(1)", "\tcommand", "több\nsor"],
            "KÜJ": ["00123"] * 5, "Mennyiség (kg)": [1.25, -2.5, 0, 0.000001, 10],
        })
        data = csv_bytes(frame)
        self.assertTrue(data.startswith(b"\xef\xbb\xbf"))
        restored = pd.read_csv(io.BytesIO(data), sep=";", decimal=",", dtype={"KÜJ": str})
        self.assertEqual(restored.iloc[0]["Név"], "Árvíztűrő; cég")
        self.assertEqual(restored.iloc[0]["KÜJ"], "00123")
        self.assertEqual(restored.iloc[1]["Név"], "'=1+1")
        self.assertEqual(restored.iloc[2]["Név"], "'  @SUM(1)")
        self.assertEqual(restored.iloc[3]["Név"], "'\tcommand")
        self.assertEqual(restored.iloc[4]["Név"], "több\nsor")
        self.assertEqual(restored.iloc[1]["Mennyiség (kg)"], -2.5)
        headers = pd.read_csv(io.BytesIO(csv_bytes(pd.DataFrame({"=1+1": ["value"]}))), sep=";")
        self.assertEqual(list(headers.columns), ["'=1+1"])

    def test_full_excel_has_all_tables_and_literal_identifiers(self):
        source, raw, missing = self.tables
        source.loc[0, "SRC_EMAIL"] = "=1+1"
        buffer = io.BytesIO()
        write_dataset_excel(buffer, source, raw, missing)
        book = load_workbook(buffer)
        self.assertEqual(book.sheetnames, ["Forráscégek", "Nyers adatok (kg)", "Kimaradt lekérések"])
        self.assertEqual(book.worksheets[1].max_row, len(raw) + 1)
        sheet = book.worksheets[0]
        self.assertEqual(sheet["B2"].value, "001234567")
        self.assertEqual(sheet["B2"].data_type, "s")
        cell = sheet.cell(2, list(source.columns).index("SRC_EMAIL") + 1)
        self.assertEqual(cell.value, "=1+1")
        self.assertEqual(cell.data_type, "s")
        self.assertEqual(sheet.freeze_panes, "A2")
        self.assertEqual(sheet.auto_filter.ref, sheet.dimensions)
        with self.assertRaisesRegex(ValueError, "szövegkorlát"):
            write_dataset_excel(io.BytesIO(), pd.DataFrame({"name": ["x" * 32768]}), raw, missing)

    def test_package_metadata_includes_persisted_cache_without_network(self):
        with tempfile.TemporaryDirectory() as directory:
            persistence = PersistenceService(Path(directory) / "cache.db")
            try:
                service = KtjService(persistence)
                persistence.save_ktj_meta("1000000", self.metadata["1000000"])
                with patch("services.ktj_service.OkirClient") as client:
                    self.assertEqual(service.export_metadata(self.tables[1]), self.metadata)
                    client.assert_not_called()
            finally:
                persistence.close()
