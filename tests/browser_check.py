"""Run explicitly: python -m tests.browser_check (Chromium + WebKit required)."""
import io
import tempfile
import threading
from pathlib import Path
from unittest.mock import patch

import pandas as pd
from openpyxl import load_workbook
from playwright.sync_api import expect, sync_playwright
from werkzeug.serving import make_server

from tests.fixtures import dataset
from web.app import create_app


def main():
    with tempfile.TemporaryDirectory() as directory, \
            patch("web.app.OkirClient") as company_client, patch("services.fetch_service.OkirClient") as waste_client:
        source, raw, _missing = dataset()
        def find_company(kshkod):
            if kshkod == "99999999":
                return []
            company = source.iloc[0 if kshkod == "12345678" else 1]
            return [{"MEGNEVEZES": company["CEGNEV"], "KUJ": company["KUJ"]}]
        company_client.return_value.fetch_companies_by_kshkod.side_effect = find_company
        waste_client.return_value.fetch.side_effect = lambda kuj, kind: raw[raw.SOURCE_KUJ == kuj].to_dict("records")
        workbook = io.BytesIO()
        pd.DataFrame([
            {"Teljes név": name, "Adószám": tax, "Megye": "Pest", "Irányítószám": "2000",
             "Település": "Szentendre", "Cím": "Minta utca"}
            for name, tax in (("Minta Fém Kft.", "12345678-2-08"), ("Minta Papír Kft.", "12345679-2-08"),
                              ("Kimaradt cég", "99999999-2-08"))
        ]).to_excel(workbook, index=False)
        app = create_app({
            "DATABASE": Path(directory) / "browser.sqlite",
            "WEB_PASSWORD": "browser-test-password", "SECRET_KEY": "browser-test-only",
        })
        server = make_server("127.0.0.1", 0, app, threaded=True)
        thread = threading.Thread(target=server.serve_forever)
        thread.start()
        try:
            with sync_playwright() as playwright:
                cases = [
                    ("chromium-desktop", playwright.chromium, {"viewport": {"width": 1440, "height": 1000}}),
                    ("webkit-iphone", playwright.webkit, playwright.devices["iPhone 13"]),
                    ("webkit-ipad", playwright.webkit, playwright.devices["iPad (gen 7)"]),
                ]
                for name, engine, device in cases:
                    browser = engine.launch()
                    context = browser.new_context(**device, accept_downloads=True)
                    page = context.new_page()
                    errors = []
                    page.on("pageerror", lambda error: errors.append(str(error)))
                    try:
                        page.goto(f"http://127.0.0.1:{server.server_port}")
                        page.get_by_label("Munkaterület jelszava").fill("browser-test-password")
                        page.get_by_role("button", name="Belépés").click()
                        expect(page.locator("#status")).not_to_have_text("Betöltés...")
                        page.get_by_role("link", name="Adatok importálása").click()
                        page.locator("#file").set_input_files({
                            "name": "minta.xlsx",
                            "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            "buffer": workbook.getvalue(),
                        })
                        page.get_by_role("button", name="Adatok betöltése", exact=True).click()
                        expect(page.locator("#status")).to_contain_text("Import kész", timeout=20000)
                        expect(page.locator("#kpi-companies")).to_have_text("2")
                        if not page.locator("#reset").is_visible():
                            page.locator("#filter-panel > summary").click()
                        page.locator("#reset").click()
                        expect(page.locator("#years option:checked")).to_have_count(0)
                        page.get_by_role("button", name="Telephelyek / KTJ", exact=True).click()
                        expect(page.locator("#page")).to_contain_text("120 sor")
                        expect(page.locator("#table tbody tr")).to_have_count(50)
                        page.locator("#next").click()
                        expect(page.locator("#page")).to_contain_text("2 / 3")
                        page.locator("#code").fill("170405")
                        expect(page.locator("#page")).to_contain_text("60 sor")
                        with page.expect_download() as downloaded:
                            page.get_by_role("link", name="Cég–HAK mátrix letöltése (.xlsx)").click()
                        download = downloaded.value
                        assert download.failure() is None
                        assert download.suggested_filename == "meh-ceg-hak-matrix.xlsx"
                        matrix = load_workbook(io.BytesIO(Path(download.path()).read_bytes()))["Cég–HAK mátrix"]
                        assert matrix["B1"].value == "170405"
                        assert matrix["B2"].value == "Vas és acél"
                        assert matrix["A3"].value == "Minta Fém Kft."
                        assert abs(matrix["B3"].value - raw.loc[raw.HULLADEKKOD == "170405", "MENNYISEG"].sum() / 1000) < .000001
                        assert not page.locator('[data-export="summary"]').is_visible()
                        expect(page.locator('[data-export="csv"],[data-export="json"],[data-export="backup"]')).to_have_count(0)
                        with page.expect_popup() as popup_info:
                            page.get_by_role("link", name="Nyomtatható riport / PDF").click()
                        popup = popup_info.value
                        expect(popup.get_by_role("heading", name="MÉH • Vezetői riport")).to_be_visible()
                        popup.close()
                        page.get_by_role("button", name="Kimaradt / hibák", exact=True).click()
                        expect(page.locator("#page")).to_contain_text("1 sor")
                        expect(page.locator("#table")).to_contain_text("Nincs OKIR találat")
                        for width in ([320, 390, 844] if "iphone" in name else [device["viewport"]["width"]]):
                            page.set_viewport_size({"width": width, "height": 844})
                            assert page.evaluate("document.documentElement.scrollWidth <= window.innerWidth"), f"{name}: horizontal overflow at {width}"
                        page.evaluate("navigator.serviceWorker.ready")
                        page.evaluate("new Promise(resolve => { if (navigator.serviceWorker.controller) resolve(); else navigator.serviceWorker.addEventListener('controllerchange', resolve, {once:true}); })")
                        context.set_offline(True)
                        page.reload()
                        expect(page.get_by_role("heading", name="Nincs kapcsolat")).to_be_visible()
                        context.set_offline(False)
                        page.get_by_role("link", name="Újrapróbálkozás").click()
                        expect(page.locator("#filename")).to_contain_text("minta.xlsx")
                        page.locator("#logout").click()
                        expect(page.get_by_label("Munkaterület jelszava")).to_be_visible()
                        assert not errors, errors
                        print(f"PASS {name}: Excel import, filters, pagination, matrix values/download, report, responsive layout, offline, logout")
                    finally:
                        context.close()
                        browser.close()
        finally:
            server.shutdown()
            thread.join()
            app.extensions["workspace"].close()


if __name__ == "__main__":
    main()
