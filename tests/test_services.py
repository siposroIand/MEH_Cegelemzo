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
from services.export_service import build_detail_export, build_summary_export, write_detail_excel
from services.fetch_service import fetch_data
from services.filter_service import Filters, filter_frame
from services.import_service import build_source_from_excel_by_adoszam
from services.persistence_service import PersistenceService
from services.report_service import csv_bytes, dataset_bytes, read_dataset, report_html, summary
from tests.fixtures import dataset


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

    def test_csv_excel_and_html_escape_untrusted_text(self):
        text = '=HYPERLINK("https://example.invalid")'
        frame = pd.DataFrame({"Szöveg": [text, "Árvíztűrő", "@SUM(1)", "\tformula"], "kg": [1, 2, -3, 0]})
        csv = csv_bytes(frame).decode("utf-8-sig")
        self.assertIn("'=HYPERLINK", csv)
        self.assertIn("'@SUM", csv)
        self.assertIn(";-3", csv)
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

    def test_json_roundtrip_and_validation(self):
        restored = read_dataset(dataset_bytes(self.source, self.raw, self.missing))
        self.assertEqual(restored[1].SOURCE_KUJ.iloc[0], "001234567")
        self.assertEqual(restored[1].MENNYISEG.sum(), self.raw.MENNYISEG.sum())
        for changes in ({"format": "pickle"}, {"quantity_unit": "t"}, {"raw": [{"foo": []}]}):
            payload = json.loads(dataset_bytes(self.source, self.raw, self.missing))
            payload.update(changes)
            with self.assertRaises(ValueError):
                read_dataset(json.dumps(payload))
            raw_with_nulls = self.raw.copy()
            raw_with_nulls.loc[0, ["EV", "MENNYISEG"]] = None
            restored = read_dataset(dataset_bytes(self.source, raw_with_nulls, self.missing))
            self.assertTrue(pd.isna(restored[1].loc[0, "EV"]))
        payload = json.loads(dataset_bytes(self.source, self.raw, self.missing))
        payload["raw"][0]["EV"] = 2025.5
        with self.assertRaises(ValueError):
            read_dataset(json.dumps(payload))

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
