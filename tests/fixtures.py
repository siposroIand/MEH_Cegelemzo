import json

import pandas as pd

from services.aggregation_service import normalize_records


def dataset(count=120):
    source = pd.DataFrame([
        {"CEGNEV": "Minta Fém Kft.", "KUJ": "001234567", "ADOSZAM": "12345678-2-08",
         "SZEKHELY": "Győr-Moson-Sopron, 9028, Győr, Minta utca 1.", "SRC_EMAIL": "minta@example.invalid",
         "SRC_NETTO_ARBEV": "1 500 000", "SRC_SAJAT_TOKE": "500 000"},
        {"CEGNEV": "Minta Papír Kft.", "KUJ": "001234568", "SZEKHELY": "Pest, 2000, Szentendre, Minta utca 2."},
    ])
    rows = []
    for index in range(count):
        company = source.iloc[index % 2]
        rows.extend(normalize_records([{
            "KUJ": company["KUJ"], "KTJ": f"00{index + 1000000}",
            "EV": 2024 + index % 3, "HULLADEKKOD": "170405" if index % 2 == 0 else "150101",
            "HULLADEK_MEGNEVEZES": "Vas és acél" if index % 2 == 0 else "Papír csomagolás",
            "HULLADEK_VESZELYES_10": int(index % 5 == 0),
            "MENNYISEG": 25.25 + index,
        }], company["CEGNEV"], company["KUJ"], "HULL_ATADAS"))
    missing = pd.DataFrame([{"Cég": "Kimaradt cég", "KÜJ": "009999999", "Adattípus": "Átadás", "Ok": "Teszt: nincs adat"}])
    return source, pd.DataFrame(rows), missing


def legacy_dataset_bytes(source, raw, missing):
    return json.dumps({
        "format": "meh-cegelemzo", "version": 1, "quantity_unit": "kg",
        "source": json.loads(source.to_json(orient="records")),
        "raw": json.loads(raw.to_json(orient="records")),
        "missing": json.loads(missing.to_json(orient="records")),
    }, ensure_ascii=False).encode("utf-8")
