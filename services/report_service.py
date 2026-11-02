import html
import json
import math
from datetime import datetime

import pandas as pd

from core.constants import ADATTIPUS_LABELS
from services.filter_service import unique_quantities

FORMAT = "meh-cegelemzo"
MAX_ROWS = 250_000
RAW_COLUMNS = {
    "CEGNEV", "SOURCE_KUJ", "ADATTIPUS", "ADATTIPUS_LABEL", "EV", "HULLADEKKOD",
    "HULLADEK_MEGNEVEZES", "HULLADEK_VESZELYES_10", "MENNYISEG",
}


def records(frame):
    return json.loads(frame.to_json(orient="records", force_ascii=False))


def read_dataset(data):
    """Read legacy backups only; JSON is no longer an export format."""
    payload = json.loads(data)
    if not isinstance(payload, dict) or payload.get("format") != FORMAT or payload.get("version") != 1:
        raise ValueError("Nem támogatott MÉH JSON formátum vagy verzió.")
    if payload.get("quantity_unit") != "kg":
        raise ValueError("A JSON mennyiségi egysége kg kell legyen.")
    frames = []
    for key in ("source", "raw", "missing"):
        rows = payload.get(key)
        if not isinstance(rows, list) or len(rows) > MAX_ROWS:
            raise ValueError(f"Érvénytelen vagy túl nagy JSON tábla: {key}")
        for row in rows:
            if not isinstance(row, dict) or len(row) > 100:
                raise ValueError(f"Érvénytelen JSON sor: {key}")
            for name, value in row.items():
                if not isinstance(name, str) or not isinstance(value, (str, int, float, bool, type(None))):
                    raise ValueError(f"Nem támogatott JSON mező: {key}")
                if isinstance(value, float) and not math.isfinite(value):
                    raise ValueError("A JSON nem tartalmazhat végtelen/NaN értékeket.")
        frames.append(pd.DataFrame(rows))
    source, raw, missing = frames
    for frame, required in ((source, {"CEGNEV", "KUJ"}), (raw, RAW_COLUMNS),
                            (missing, {"Cég", "KÜJ", "Adattípus", "Ok"})):
        if not frame.empty and not required.issubset(frame.columns):
            raise ValueError("Hiányzó JSON oszlopok: " + ", ".join(sorted(required - set(frame.columns))))
    if not raw.empty:
        for column in ("EV", "MENNYISEG", "HULLADEK_VESZELYES_10"):
            raw[column] = pd.to_numeric(raw[column], errors="raise")
        years = raw["EV"].dropna()
        if not years.between(1900, 2200).all() or not (years % 1 == 0).all():
            raise ValueError("Érvénytelen év.")
        if not raw["ADATTIPUS"].isin(ADATTIPUS_LABELS).all():
            raise ValueError("Érvénytelen adattípus.")
        if not raw["HULLADEK_VESZELYES_10"].dropna().isin([0, 1]).all():
            raise ValueError("Érvénytelen veszélyesség.")
        raw["ADATTIPUS_LABEL"] = raw["ADATTIPUS"].map(ADATTIPUS_LABELS)
        for column in ("CEGNEV", "SOURCE_KUJ", "HULLADEKKOD"):
            if not raw[column].map(lambda value: isinstance(value, str) and bool(value.strip())).all():
                raise ValueError(f"Érvénytelen szövegazonosító: {column}")
    if not source.empty:
        for column in ("CEGNEV", "KUJ"):
            if not source[column].map(lambda value: isinstance(value, str) and bool(value.strip())).all():
                raise ValueError(f"Érvénytelen forrásazonosító: {column}")
    return source, raw, missing


def summary(master):
    base = unique_quantities(master)
    if base.empty:
        return {"quantity_kg": 0, "companies": 0, "codes": 0, "dangerous_percent": 0,
                "years": [], "yearly": [], "categories": [], "company_totals": [], "waste_codes": []}
    total = float(base["OSSZES_MENNYISEG"].sum())
    dangerous = float(base.loc[base["HULLADEK_VESZELYES_10"] == 1, "OSSZES_MENNYISEG"].sum())

    def grouped(frame, column, limit=None):
        out = frame.groupby(column, as_index=False, dropna=False)["OSSZES_MENNYISEG"].sum()
        out = out.sort_values(column if column == "EV" else "OSSZES_MENNYISEG", ascending=column == "EV")
        return records(out if limit is None else out.head(limit))

    return {
        "quantity_kg": total, "companies": int(base["CEGNEV"].nunique()),
        "codes": int(base["HULLADEKKOD"].nunique()),
        "dangerous_percent": dangerous / total * 100 if total else 0,
        "years": sorted(int(year) for year in base["EV"].dropna().unique()),
        "yearly": grouped(base, "EV"), "categories": grouped(master, "PRODUCT_GROUP"),
        "company_totals": grouped(base, "CEGNEV", 15),
        "waste_codes": grouped(base, "HULLADEKKOD", 15),
    }


def report_html(master, missing, filter_description="Összes adat"):
    data = summary(master)
    total = f'{data["quantity_kg"] / 1000:,.2f}'.replace(",", " ")
    sections = []
    for title, key, label in (("Éves mennyiségek", "yearly", "EV"),
                              ("Termékcsoportok", "categories", "PRODUCT_GROUP"),
                              ("Legnagyobb cégek (legfeljebb 15)", "company_totals", "CEGNEV")):
        frame = pd.DataFrame(data[key])
        if not frame.empty:
            frame = frame.rename(columns={label: title, "OSSZES_MENNYISEG": "Mennyiség (kg)"})
            sections.append(f"<h2>{title}</h2>" + frame.to_html(index=False, escape=True))
    return f"""<!doctype html><html lang="hu"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>MÉH – Vezetői riport</title><style>
body{{font:16px system-ui,sans-serif;color:#142d37;max-width:1000px;margin:32px auto;padding:20px}}
h1{{color:#087f74}}table{{border-collapse:collapse;width:100%;margin-bottom:24px}}
th,td{{border-bottom:1px solid #ccd8dc;padding:9px;text-align:left;overflow-wrap:anywhere}}
small{{color:#48616b}}@media print{{body{{margin:0}}.print-help{{display:none}}
thead{{display:table-header-group}}tr{{break-inside:avoid}}}}
</style></head><body><h1>MÉH • Vezetői riport</h1>
<p class="print-help">PDF mentés: böngésző → Nyomtatás → Mentés PDF-ként.
iOS: Megosztás → Nyomtatás → előnézet megnyitása → Megosztás → Mentés a Fájlokba.</p>
<small>{html.escape(datetime.now().isoformat(timespec="minutes"))} |
{html.escape(filter_description)}</small>
<p><b>{total} t</b> összes mennyiség · {data["companies"]} cég · {data["codes"]} HAK kód</p>
<p>Veszélyes hulladék aránya: {data["dangerous_percent"]:.2f}%.
Kimaradt/hibás lekérések a teljes munkamenetben: {len(missing)}.</p>
<p>A termékcsoportok átfedhetnek, ezért összegeik nem adhatók össze.
Az összmennyiség a kiválasztott adattípusok összege, nem nettó anyagmérleg.</p>
{''.join(sections)}</body></html>"""
