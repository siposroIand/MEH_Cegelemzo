import io
from pathlib import Path
import tempfile
import threading
import time
import unittest
from unittest.mock import patch
from urllib.parse import urlencode

import pandas as pd
from openpyxl import load_workbook

from tests.fixtures import dataset, legacy_dataset_bytes
from web.app import create_app


class WebTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.config = {"TESTING": True, "DATABASE": Path(self.temp.name) / "test.sqlite",
                       "WEB_PASSWORD": "only-a-test-password", "SECRET_KEY": "test-key-not-for-production"}
        self.app = create_app(self.config)
        self.workspace = self.app.extensions["workspace"]
        self.client = self.app.test_client()
        self.addCleanup(self.temp.cleanup)
        self.addCleanup(self.workspace.close)
        self.source, self.raw, self.missing = dataset()

    def login(self):
        self.client.get("/login")
        with self.client.session_transaction() as session:
            csrf = session["csrf"]
        response = self.client.post("/login", data={"password": self.config["WEB_PASSWORD"], "csrf": csrf})
        self.assertEqual(response.status_code, 302)
        self.csrf = self.client.get("/api/state").json["csrf"]

    def wait(self):
        for _ in range(250):
            job = self.client.get("/api/state").json["job"]
            if job["status"] != "running":
                return job
            time.sleep(.02)
        self.fail("Background job did not finish")

    def upload(self, payload=None):
        response = self.client.post("/api/import", headers={"X-CSRF-Token": self.csrf},
                                    data={"file": (io.BytesIO(payload or legacy_dataset_bytes(self.source, self.raw, self.missing)), "minta.json")})
        self.assertEqual(response.status_code, 202)
        return self.wait()

    def test_login_csrf_logout_and_cache_headers(self):
        self.assertEqual(self.client.get("/api/view").status_code, 401)
        self.assertEqual(self.client.get("/").status_code, 302)
        self.login()
        self.assertEqual(self.client.post("/api/logout").status_code, 403)
        response = self.client.get("/api/state")
        self.assertEqual(response.headers["Cache-Control"], "no-store")
        self.assertIn("frame-ancestors 'none'", response.headers["Content-Security-Policy"])
        self.assertEqual(self.client.post("/api/logout", headers={"X-CSRF-Token": self.csrf}).status_code, 200)
        self.assertEqual(self.client.get("/api/state").status_code, 401)

    def test_import_pagination_filter_and_exports(self):
        self.login()
        self.assertEqual(self.upload()["status"], "done")
        result = self.client.get("/api/view?view=ktj&page_size=10000").json
        self.assertEqual(len(result["rows"]), 100)
        self.assertEqual(result["total"], 120)
        self.assertAlmostEqual(result["summary"]["quantity_kg"], self.raw.MENNYISEG.sum())
        next_page = self.client.get("/api/view?view=ktj&page=2&page_size=100").json
        self.assertEqual(len(next_page["rows"]), 20)
        query = urlencode({"years": 2025, "companies": "Minta Fém Kft."})
        view = self.client.get("/api/view?" + query).json
        expected = self.raw[(self.raw.EV == 2025) & (self.raw.CEGNEV == "Minta Fém Kft.")]
        self.assertEqual(view["summary"]["quantity_kg"], expected.MENNYISEG.sum())
        response = self.client.get("/api/export/matrix?" + query)
        self.assertEqual(response.status_code, 200, response.data[:100])
        self.assertIn("meh-ceg-hak-matrix.xlsx", response.headers["Content-Disposition"])
        book = load_workbook(io.BytesIO(response.data))
        matrix = list(book["Cég–HAK mátrix"].values)
        self.assertEqual(matrix[0], ("Cég", "170405"))
        self.assertEqual(matrix[1], ("Mennyiség (t)", "Vas és acél"))
        self.assertEqual(matrix[2][0], "Minta Fém Kft.")
        self.assertAlmostEqual(matrix[2][1], expected.MENNYISEG.sum() / 1000)
        self.assertIn("Év: 2025", book["Szűrés"]["B2"].value)
        for kind in ("summary", "detail", "ktj"):
            response = self.client.get(f"/api/export/{kind}?{query}")
            self.assertEqual(response.status_code, 200, response.data[:200])
            self.assertIn("attachment", response.headers["Content-Disposition"])
            self.assertTrue(load_workbook(io.BytesIO(response.data)).sheetnames)
        errors = self.client.get("/api/export/errors?" + query)
        self.assertEqual(load_workbook(io.BytesIO(errors.data)).worksheets[0]["A2"].value, "Kimaradt cég")
        report = self.client.get("/api/export/report?" + query)
        self.assertIn("Minta Fém Kft.", report.text)
        for removed in ("csv", "json", "backup", "missing"):
            self.assertEqual(self.client.get(f"/api/export/{removed}").status_code, 404)
        self.assertEqual(self.client.get("/api/export/matrix?years=1900").status_code, 400)
        html = self.client.get("/").text
        for removed in ("csv", "json", "backup", "missing"):
            self.assertNotIn(f'data-export="{removed}"', html)
        self.assertIn('data-export="matrix"', html)

    def test_invalid_upload_keeps_current_data(self):
        self.login()
        self.upload()
        self.assertEqual(self.upload(b'{"format":"wrong"}')["status"], "error")
        self.assertEqual(self.client.get("/api/state").json["rows"], 120)
        response = self.client.post("/api/import", headers={"X-CSRF-Token": self.csrf},
                                    data={"file": (io.BytesIO(b"pickle"), "unsafe.okirpkg")})
        self.assertEqual(response.status_code, 400)
        self.assertEqual(self.client.get("/api/view?years=abc").status_code, 400)

    def test_excel_import_and_ktj_metadata(self):
        self.login()
        workbook = io.BytesIO()
        pd.DataFrame([{"Teljes név": "Minta", "Adószám": "12345678-2-08", "Megye": "Pest",
                       "Irányítószám": "2000", "Település": "Szentendre", "Cím": "Minta utca"}]).to_excel(workbook, index=False)
        workbook.seek(0)
        with patch("web.app.OkirClient") as companies, patch("services.fetch_service.OkirClient") as waste:
            companies.return_value.fetch_companies_by_kshkod.return_value = [{"MEGNEVEZES": "Minta", "KUJ": "001234567"}]
            waste.return_value.fetch.return_value = [{"EV": 2025, "KTJ": "123456789",
                                                     "HULLADEKKOD": "170405", "MENNYISEG": 123.5}]
            response = self.client.post("/api/import", headers={"X-CSRF-Token": self.csrf},
                                        data={"file": (workbook, "cegek.xlsx"), "types": ["HULL_ATADAS"], "use_cache": "true"})
            self.assertEqual(response.status_code, 202)
            self.assertEqual(self.wait()["status"], "done")
        self.assertEqual(self.client.get("/api/view").json["summary"]["quantity_kg"], 123.5)
        with patch("services.ktj_service.OkirClient") as client:
            client.return_value.fetch_first_ktj_detail.return_value = {"MEGNEVEZES": "Minta telephely", "HELYSEG": "Szentendre"}
            self.client.post("/api/ktj", headers={"X-CSRF-Token": self.csrf})
            self.assertEqual(self.wait()["status"], "done")
            client.return_value.fetch_first_ktj_detail.assert_called_once()
            response = self.client.get("/api/view?view=ktj")
            self.assertEqual(response.json["rows"][0]["KTJ megnevezés"], "Minta telephely")
            self.client.get("/api/export/ktj")
            client.return_value.fetch_first_ktj_detail.assert_called_once()

    def test_upload_size_limit(self):
        self.login()
        self.app.config["MAX_CONTENT_LENGTH"] = 100
        response = self.client.post("/api/import", headers={"X-CSRF-Token": self.csrf},
                                    data={"file": (io.BytesIO(b"x" * 200), "large.json")})
        self.assertEqual(response.status_code, 413)

    def test_single_job_and_persistent_restart(self):
        self.login()
        self.upload()
        gate = threading.Event()
        self.workspace.submit(lambda: gate.wait(5))
        try:
            response = self.client.post("/api/snapshots", json={"note": "new"}, headers={"X-CSRF-Token": self.csrf})
            self.assertEqual(response.status_code, 409)
        finally:
            gate.set()
        self.wait()
        self.workspace.close()
        app = create_app(self.config)
        self.addCleanup(app.extensions["workspace"].close)
        self.assertEqual(len(app.extensions["workspace"].dataset.raw), 120)

    def test_snapshot_management(self):
        self.login()
        self.upload()
        headers = {"X-CSRF-Token": self.csrf}
        sid = self.client.get("/api/state").json["snapshots"][0][0]
        self.client.delete(f"/api/snapshots/{sid}", headers=headers)
        self.assertEqual(self.wait()["status"], "error")
        self.client.post("/api/snapshots", json={"note": "new"}, headers=headers)
        self.assertEqual(self.wait()["status"], "done")
        self.client.delete(f"/api/snapshots/{sid}", headers=headers)
        self.assertEqual(self.wait()["status"], "done")
        self.assertEqual(len(self.client.get("/api/state").json["snapshots"]), 1)

    def test_pwa_assets_and_offline_policy(self):
        for path in ("/sw.js", "/static/manifest.webmanifest", "/static/apple-touch-icon.png",
                     "/static/icon-192.png", "/static/icon-512.png", "/static/offline.html"):
            with self.client.get(path) as response:
                self.assertEqual(response.status_code, 200, path)
        with self.client.get("/static/manifest.webmanifest") as response:
            manifest = response.json
        self.assertEqual(manifest["display"], "standalone")
        with self.client.get("/sw.js") as response:
            script = response.text
        self.assertNotIn('"/api/', script)
        self.assertNotIn('"/login"', script)


if __name__ == "__main__":
    unittest.main()
