import re
import tempfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from uuid import uuid4

import pandas as pd

from core.utils import safe_str, split_address_parts
from services.desktop_export_service import atomic_export
from services.export_service import (
    build_detail_export,
    build_summary_export,
    style_excel_worksheet,
    validate_excel_tables,
    write_excel_table,
    write_hak_matrix_excel,
)
from services.filter_service import Filters, filter_frame


EXCEL_SECTIONS = {
    "matrix": "Cég × Év / HAK mátrix (t)",
    "summary": "Termékcsoportos összesítő (t)",
    "detail": "Részletes HAK-lista (kg és t)",
    "source": "Forráscégek",
    "raw": "Nyers adatok (kg)",
    "missing": "Kimaradt / hibás lekérések",
}
UNKNOWN_COUNTY = "Ismeretlen megye"


@dataclass(frozen=True)
class ExcelExportOptions:
    sections: tuple[str, ...] = ("matrix",)
    use_filters: bool = True
    counties: tuple[str, ...] = ()
    separate_counties: bool = False


def source_counties(source):
    return source.get("SZEKHELY", pd.Series("", index=source.index)).map(
        lambda value: split_address_parts(value)["Megye"] or UNKNOWN_COUNTY
    )


def county_labels(source, frame, name_column, kuj_column):
    lookup = {}
    for name, kuj, county in zip(
        source.get("CEGNEV", pd.Series("", index=source.index)),
        source.get("KUJ", pd.Series("", index=source.index)),
        source_counties(source),
    ):
        key = (safe_str(name), safe_str(kuj))
        if key in lookup and lookup[key] != county:
            raise ValueError(f"Eltérő megyeadatok ugyanahhoz a céghez: {key[0]} ({key[1]}).")
        lookup[key] = county
    if frame.empty:
        return pd.Series(index=frame.index, dtype=str)
    return pd.Series([
        lookup.get((safe_str(name), safe_str(kuj)), UNKNOWN_COUNTY)
        for name, kuj in zip(
            frame.get(name_column, pd.Series("", index=frame.index)),
            frame.get(kuj_column, pd.Series("", index=frame.index)),
        )
    ], index=frame.index, dtype=str)


def available_counties(source, raw, missing):
    return sorted(
        set(source_counties(source))
        | set(county_labels(source, raw, "CEGNEV", "SOURCE_KUJ"))
        | set(county_labels(source, missing, "Cég", "KÜJ"))
    )


def write_configured_excel(path, master, source, raw, missing, sections, groups, description):
    tables = []
    for section in sections:
        if section == "summary":
            tables.append(("Összesítő", build_summary_export(master, source, groups)))
        elif section == "detail":
            tables.append(("Részletes", build_detail_export(master, source)))
        elif section == "source":
            tables.append(("Forráscégek", source))
        elif section == "raw":
            tables.append(("Nyers adatok (kg)", raw))
        elif section == "missing":
            tables.append(("Kimaradt lekérések", missing))
    validate_excel_tables(tables)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        # A first sheet also keeps writer cleanup valid if a later sheet fails.
        write_excel_table(writer, "Export beállításai", pd.DataFrame([
            ("Adatkör", description),
            ("Munkalapok", "; ".join(EXCEL_SECTIONS[key] for key in sections)),
            ("Megyebontás alapja", "A forráscég székhelye, nem a KTJ telephelye."),
            ("Forrás és hibák szűrése", "Cég és megye szerint; az év, HAK, adattípus és termékcsoport csak a hulladékadatokra hat."),
        ], columns=["Beállítás", "Érték"]))
        if "matrix" in sections:
            if master.empty:
                write_excel_table(writer, "Cég–Év–HAK mátrix", pd.DataFrame(
                    {"Tájékoztatás": ["Nincs a szűrésnek megfelelő hulladékadat."]}
                ))
            else:
                write_hak_matrix_excel(path, master, description, by_year=True, workbook=writer.book)
        for name, frame in tables:
            write_excel_table(writer, name, frame)
            if name in {"Összesítő", "Részletes"}:
                style_excel_worksheet(writer.sheets[name])


def export_configured_excel(destination, source, raw, missing, master, options, filters, groups, description):
    if not options.sections or any(key not in EXCEL_SECTIONS for key in options.sections):
        raise ValueError("Válassz legalább egy érvényes munkalapot.")
    if len(set(options.sections)) != len(options.sections):
        raise ValueError("Egy munkalap csak egyszer választható ki.")
    active = filters if options.use_filters else Filters()
    master = filter_frame(master, active)
    # The aggregated master already carries category membership.
    raw_filters = Filters(companies=active.companies, years=active.years, code=active.code, types=active.types)
    raw = filter_frame(raw, raw_filters)
    if active.categories and not raw.empty:
        raw = raw[raw["HULLADEKKOD"].isin(master.get("HULLADEKKOD", []))]
    selected_source = source
    if active.companies:
        selected_source = source[source["CEGNEV"].isin(active.companies)] if not source.empty else source
        missing = missing[missing["Cég"].isin(active.companies)] if not missing.empty else missing
    frames = (master, selected_source, raw, missing)
    labels = (
        county_labels(source, master, "CEGNEV", "SOURCE_KUJ"),
        source_counties(selected_source),
        county_labels(source, raw, "CEGNEV", "SOURCE_KUJ"),
        county_labels(source, missing, "Cég", "KÜJ"),
    )
    if options.counties:
        frames = tuple(frame.loc[label.isin(options.counties)] for frame, label in zip(frames, labels))
        labels = tuple(label.loc[label.isin(options.counties)] for label in labels)
    section_frame = {"matrix": 0, "summary": 0, "detail": 0, "source": 1, "raw": 2, "missing": 3}
    used = {section_frame[key] for key in options.sections}
    if all(frames[index].empty for index in used):
        raise ValueError("Nincs a kiválasztott munkalapoknak és szűrésnek megfelelő adat.")
    description = description if options.use_filters else "Teljes adatállomány (a főablak szűrőitől függetlenül)"
    description += "; Megyék: " + (", ".join(options.counties) or "Összes")
    destination = Path(destination)
    if not options.separate_counties:
        return atomic_export(destination, write_configured_excel, *frames, options.sections, groups, description)

    counties = sorted(set().union(*(set(labels[index]) for index in used)))
    output = destination / f"meh-excel-{datetime.now():%Y%m%d-%H%M%S}-{uuid4().hex[:8]}"
    # Publish the entire batch only after every workbook has been written.
    with tempfile.TemporaryDirectory(prefix=".meh-excel-", dir=destination) as temporary:
        staged = Path(temporary) / "export"
        staged.mkdir()
        for number, county in enumerate(counties, 1):
            filename = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", county).strip(" .")[:80] or "megye"
            subset = tuple(frame.loc[label.eq(county)] for frame, label in zip(frames, labels))
            write_configured_excel(
                staged / f"{number:02d}-meh-{filename}.xlsx", *subset, options.sections,
                groups, description + f"; Fájl megyéje: {county}",
            )
        staged.rename(output)
    return str(output)
