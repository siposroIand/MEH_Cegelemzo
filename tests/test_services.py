import io
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import pandas as pd
from openpyxl import load_workbook

from services.aggregation_service import aggregate_data, normalize_records
from services.cache_db import CacheDB
from services.category_service import CategoryService
from services.export_service import build_detail_export, build_summary_export, write_detail_excel, write_hak_matrix_excel
from services.fetch_service import fetch_data
from services.filter_service import Filters, filter_frame
from services.import_service import build_source_from_excel_by_adoszam
from services.persistence_service import PersistenceService
from services.report_service import read_dataset, report_html, summary
from tests.fixtures import dataset, legacy_dataset_bytes


class ServiceTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        path = Path(self.temp.name) / "groups.json"
        path.write_text(json.dumps({"Vas": ["170405"], "Fém": ["170405"], "Papír": ["150101"], "Nem besorolható": []}), encoding="utf-8")
        self.categories = CategoryService(path)
        self.source, self.raw, self.missing = dataset()

    def test_overlap_filters_and_total(self):
        aggregates = aggregate_data(self.raw, self.source, self.categories)
        self.assertAlmostEqual(aggregates.base_unique.OSSZES_MENNYISEG.sum(), self.raw.MENNYISEG.sum())
        self.assertEqual(self.categories.resolve_group("170405"), "Vas; Fém")
        selected = Filters(categories=("Fém",), years=(2025,), code="1704;999")
        raw = filter_frame(self.raw, selected, self.categories)
        master = filter_frame(aggregates.master, selected)
        self.assertGreater(len(raw), 0)
        self.assertAlmostEqual(summary(master)["quantity_kg"], raw.MENNYISEG.sum())
        self.assertTrue(filter_frame(self.raw, Filters(code=".*"), self.categories).empty)

    def test_distinct_types_and_descriptions_are_not_dropped(self):
        rows = []
        for kind in ("HULL_ATADAS", "HULL_ATVETEL"):
            for name in ("Leírás 1", "Leírás 2"):
                rows.extend(normalize_records([{"EV": 2025, "HULLADEKKOD": "170405", "HULLADEK_MEGNEVEZES": name,
                                                "MENNYISEG": 40, "HULLADEK_VESZELYES_10": 0}],
                                              "Minta Fém Kft.", "001234567", kind))
        master = aggregate_data(pd.DataFrame(rows), self.source, self.categories).master
        out = build_detail_export(master, self.source)
        self.assertEqual(len(out), 4)
        self.assertEqual(set(out["Adattípus"]), {"Átadás", "Átvétel"})
        self.assertEqual(out["Mennyiség (kg)"].sum(), 160)
        self.assertEqual(summary(master)["quantity_kg"], 160)
        totals = build_summary_export(master, self.source, self.categories.all_groups())
        self.assertAlmostEqual(totals["Össz hulladék (tonna)"].sum(), .2)

    def test_excel_and_html_escape_untrusted_text(self):
        text = '=HYPERLINK("https://example.invalid")'
        raw = self.raw.copy()
        raw["CEGNEV"] = text
        master = aggregate_data(raw, pd.DataFrame(), self.categories).master
        buffer = io.BytesIO()
        write_detail_excel(buffer, master, pd.DataFrame())
        book = load_workbook(buffer)
        self.assertEqual(book["Részletes"]["A2"].data_type, "s")
        self.assertEqual(book["Részletes"]["A2"].value, text)
        self.assertEqual(book["Kimutatás"]["B2"].data_type, "s")
        report = report_html(master, self.missing, "<script>alert(1)</script>")
        self.assertNotIn("<script>", report)
        self.assertIn("&lt;script&gt;", report)
        self.assertIn("PDF", report)

    def test_legacy_json_import_and_validation(self):
        restored = read_dataset(legacy_dataset_bytes(self.source, self.raw, self.missing))
        self.assertEqual(restored[1].SOURCE_KUJ.iloc[0], "001234567")
        self.assertEqual(restored[1].MENNYISEG.sum(), self.raw.MENNYISEG.sum())
        for changes in ({"format": "pickle"}, {"quantity_unit": "t"}, {"raw": [{"foo": []}]}):
            payload = json.loads(legacy_dataset_bytes(self.source, self.raw, self.missing))
            payload.update(changes)
            with self.assertRaises(ValueError):
                read_dataset(json.dumps(payload))
        raw_with_nulls = self.raw.copy()
        raw_with_nulls.loc[0, ["EV", "MENNYISEG"]] = None
        restored = read_dataset(legacy_dataset_bytes(self.source, raw_with_nulls, self.missing))
        self.assertTrue(pd.isna(restored[1].loc[0, "EV"]))
        payload = json.loads(legacy_dataset_bytes(self.source, self.raw, self.missing))
        payload["raw"][0]["EV"] = 2025.5
        with self.assertRaises(ValueError):
            read_dataset(json.dumps(payload))

    def test_hak_matrix_shape_units_and_overlaps(self):
        rows = []
        for company, kuj, year, code, name, kind, quantity in (
            ("Alpha", "001", 2024, "001001", "Első", "HULL_ATADAS", 125),
            ("Alpha", "001", 2025, "001001", "Első", "HULL_ATADAS", 75),
            ("Alpha", "002", 2025, "001001", "Másik", "HULL_ATVETEL", 200),
            ("Alpha", "001", 2025, "170405", "Vas", "HULL_ATADAS", .125),
            ("Beta", "003", 2025, "170405", "Vas", "HULL_ATADAS", 1000),
            ("Beta", "003", 2025, "999999", None, "HULL_ATADAS", -25),
        ):
            rows.extend(normalize_records([{"EV": year, "HULLADEKKOD": code, "HULLADEK_MEGNEVEZES": name,
                                            "MENNYISEG": quantity}], company, kuj, kind))
        master = aggregate_data(pd.DataFrame(rows), pd.DataFrame(), self.categories).master
        buffer = io.BytesIO()
        write_hak_matrix_excel(buffer, master, "Minden év és adattípus")
        workbook = load_workbook(buffer)
        sheet = workbook["Cég–HAK mátrix"]
        self.assertEqual(list(sheet.values), [
            ("Cég", "001001", "170405", "999999"),
            ("Mennyiség (t)", "Első; Másik", "Vas", "Nincs megnevezés"),
            ("Alpha", .4, .000125, 0),
            ("Beta", 0, 1, -.025),
        ])
        self.assertEqual(sheet.freeze_panes, "B3")
        self.assertEqual(sheet["B1"].data_type, "s")
        self.assertEqual(sheet["B3"].data_type, "n")
        self.assertEqual(sheet["B3"].number_format, "#,##0.000")
        self.assertEqual(workbook["Szűrés"]["B2"].value, "Minden év és adattípus")
        self.assertAlmostEqual(sum(sum(row[1:]) for row in list(sheet.values)[2:]) * 1000,
                               summary(master)["quantity_kg"])

        buffer = io.BytesIO()
        write_hak_matrix_excel(buffer, master, "Minden év és adattípus", by_year=True)
        sheet = load_workbook(buffer)["Cég–Év–HAK mátrix"]
        self.assertEqual(list(sheet.values), [
            ("ÉV", 2024, 2025, 2025, 2025, "Összes év"),
            ("HAK", "001001", "001001", "170405", "999999", "Összes hulladék"),
            ("Cég / HAK MEGNEVEZÉS (t)", "Első", "Első; Másik", "Vas", "Nincs megnevezés", "Összes hulladék (t)"),
            ("Alpha", .125, .275, .000125, 0, .400125),
            ("Beta", 0, 0, 1, -.025, .975),
        ])
        self.assertEqual(sheet.freeze_panes, "B4")
        self.assertEqual(sheet.print_title_rows, "$1:$3")
        self.assertEqual(sheet["B2"].data_type, "s")
        self.assertEqual(sheet["E4"].value, 0)
        self.assertEqual(sheet["E4"].number_format.split(";")[2], '"-"')
        self.assertEqual(sheet["E4"].fill.patternType, None)
        self.assertEqual(sheet["B5"].fill.patternType, None)
        self.assertEqual(sheet["D4"].fill.fgColor.rgb, "00CCF0E7")
        self.assertEqual(sheet["E5"].fill.fgColor.rgb, "00FFE4CC")
        self.assertTrue(all(cell.data_type == "n" for row in sheet.iter_rows(min_row=4, min_col=2) for cell in row))

        selected = Filters(years=(2025,), categories=("Fém",), types=("HULL_ATADAS",), code="1704")
        buffer = io.BytesIO()
        write_hak_matrix_excel(buffer, filter_frame(master, selected))
        self.assertEqual(list(load_workbook(buffer).worksheets[0].values), [
            ("Cég", "170405"), ("Mennyiség (t)", "Vas"), ("Alpha", .000125), ("Beta", 1),
        ])

    def test_hak_matrix_rejects_empty_or_oversized_and_writes_literal_text(self):
        with self.assertRaisesRegex(ValueError, "Nincs"):
            write_hak_matrix_excel(io.BytesIO(), pd.DataFrame())
        too_wide = pd.DataFrame({"CEGNEV": ["A"] * 16384, "HULLADEKKOD": [str(i) for i in range(16384)]})
        with self.assertRaisesRegex(ValueError, "mérethatár"):
            write_hak_matrix_excel(io.BytesIO(), too_wide)
        raw = self.raw.copy()
        raw["CEGNEV"] = "=1+1"
        raw["HULLADEK_MEGNEVEZES"] = "=2+2"
        master = aggregate_data(raw, pd.DataFrame(), self.categories).master
        buffer = io.BytesIO()
        write_hak_matrix_excel(buffer, master, "=3+3")
        workbook = load_workbook(buffer)
        for sheet in workbook:
            for row in sheet:
                self.assertTrue(all(cell.data_type != "f" for cell in row))
        self.assertEqual(workbook.worksheets[0]["A3"].value, "=1+1")
        buffer = io.BytesIO()
        write_hak_matrix_excel(buffer, master, "=3+3", by_year=True)
        for sheet in load_workbook(buffer):
            self.assertTrue(all(cell.data_type != "f" for row in sheet for cell in row))

    def test_year_matrix_keeps_missing_year_and_rejects_excess_year_code_columns(self):
        master = aggregate_data(self.raw, self.source, self.categories).master
        master["EV"] = None
        buffer = io.BytesIO()
        write_hak_matrix_excel(buffer, master, by_year=True)
        sheet = load_workbook(buffer).worksheets[0]
        self.assertEqual(sheet["B1"].value, "Ismeretlen év")
        self.assertAlmostEqual(sum(row[-1] for row in list(sheet.values)[3:]),
                               summary(master)["quantity_kg"] / 1000)
        wide = pd.DataFrame({
            "CEGNEV": ["A"] * 16383, "EV": list(range(16383)), "HULLADEKKOD": ["001001"] * 16383,
        })
        with self.assertRaisesRegex(ValueError, "mérethatár"):
            write_hak_matrix_excel(io.BytesIO(), wide, by_year=True)

    def test_snapshot_roundtrip_and_index(self):
        persistence = PersistenceService(Path(self.temp.name) / "db.sqlite")
        self.addCleanup(persistence.close)
        sid = persistence.save_snapshot(self.source, self.raw, self.missing, "minta.xlsx", "teszt")
        loaded = persistence.load_snapshot(sid)
        self.assertEqual(len(loaded[1]), len(self.raw))
        self.assertEqual(loaded[1].SOURCE_KUJ.iloc[0], "001234567")
        indexes = persistence._conn().execute("PRAGMA index_list(snapshot_raw_rows)").fetchall()
        self.assertTrue(indexes)
        persistence.delete_snapshot(sid)
        self.assertEqual(persistence.list_snapshots(), [])
        with self.assertRaises(TypeError):
            persistence.save_snapshot(self.source, pd.DataFrame({"invalid": [object()]}), self.missing)
        self.assertEqual(persistence.list_snapshots(), [])

    @patch("services.fetch_service.OkirClient")
    def test_cached_fetch_no_network_and_force_refresh(self, client):
        path = Path(self.temp.name) / "cache.sqlite"
        cache = CacheDB(path)
        cached = [{"EV": 2025, "HULLADEKKOD": "170405", "MENNYISEG": 50}]
        cache.save("001234567", "HULL_ATADAS", "success", json.dumps(cached))
        cache.close()
        result = fetch_data(self.source.iloc[:1], True, False, ["HULL_ATADAS"], cache_path=path)
        self.assertEqual(result.raw.MENNYISEG.sum(), 50)
        client.return_value.fetch.assert_not_called()
        client.return_value.fetch.return_value = [{"EV": 2025, "HULLADEKKOD": "170405", "MENNYISEG": 75}]
        result = fetch_data(self.source.iloc[:1], True, True, ["HULL_ATADAS"], cache_path=path)
        self.assertEqual(result.raw.MENNYISEG.sum(), 75)
        client.return_value.fetch.assert_called_once()

    @patch("services.fetch_service.OkirClient")
    def test_parallel_fetch_initializes_new_cache_once(self, client):
        client.return_value.fetch.return_value = [{"EV": 2025, "HULLADEKKOD": "170405", "MENNYISEG": 75}]
        with patch("services.fetch_service.CacheDB", wraps=CacheDB) as cache_constructor:
            result = fetch_data(
                self.source, True, False, ["HULL_ATADAS", "HULL_ATVETEL"],
                max_workers=6, cache_path=Path(self.temp.name) / "new-cache.sqlite",
            )
        cache_constructor.assert_called_once()
        self.assertTrue(result.missing.empty)
        self.assertEqual(len(result.raw), 4)

    @patch("services.fetch_service.OkirClient")
    def test_fetch_failures_are_reported(self, client):
        client.return_value.fetch.side_effect = ValueError("OKIR nem elérhető")
        result = fetch_data(self.source, False, False, ["HULL_ATADAS"], cache_path=Path(self.temp.name) / "cache.sqlite")
        self.assertTrue(result.raw.empty)
        self.assertEqual(len(result.missing), 2)
        self.assertTrue(result.missing.Ok.str.contains("OKIR nem elérhető").all())

    def test_import_memoizes_tax_lookup_and_reports_invalid_rows(self):
        from unittest.mock import Mock
        client = Mock()
        client.fetch_companies_by_kshkod.return_value = [{"MEGNEVEZES": "Minta", "KUJ": "001", "KSHKOD": "12345678"}]
        frame = pd.DataFrame({"Teljes név": ["A", "A", "B"], "Adószám": ["12345678-2-08", "12345678-2-08", "x"],
                              "Megye": ["Pest"] * 3, "Irányítószám": ["2000"] * 3,
                              "Település": ["Minta"] * 3, "Cím": ["Utca"] * 3})
        buffer = io.BytesIO()
        frame.to_excel(buffer, index=False)
        buffer.seek(0)
        failures = []
        result = build_source_from_excel_by_adoszam(buffer, client, missing_fn=failures.append)
        self.assertEqual(len(result), 1)
        client.fetch_companies_by_kshkod.assert_called_once()
        self.assertEqual(len(failures), 1)

    def test_empty_exports(self):
        buffer = io.BytesIO()
        write_detail_excel(buffer, pd.DataFrame(), pd.DataFrame())
        self.assertEqual(load_workbook(buffer).sheetnames, ["Részletes", "Kimutatás"])
        self.assertEqual(summary(pd.DataFrame())["quantity_kg"], 0)


if __name__ == "__main__":
    unittest.main()
