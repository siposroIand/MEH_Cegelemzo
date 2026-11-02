from typing import List

import pandas as pd
from services.filter_service import unique_quantities

from core.utils import (
    clean_int_like,
    format_phone,
    format_quantity_auto_unit,
    parse_hu_numeric,
    split_address_parts,
)


def build_detail_export(master: pd.DataFrame, source_df: pd.DataFrame) -> pd.DataFrame:
    if master.empty:
        return pd.DataFrame()
    src_cols = [
        "CEGNEV", "KUJ", "SZEKHELY", "ADOSZAM",
        "SRC_EMAIL", "SRC_TELEFON", "SRC_WEB",
    ]
    src = source_df.copy()
    for c in src_cols:
        if c not in src.columns:
            src[c] = ""
    src = src[src_cols].drop_duplicates(subset=["CEGNEV", "KUJ"])

    out = unique_quantities(master).merge(
        src, how="left",
        left_on=["CEGNEV", "SOURCE_KUJ"],
        right_on=["CEGNEV", "KUJ"],
    )

    parts = out["SZEKHELY"].apply(split_address_parts).apply(pd.Series)

    out["SRC_TELEFON"] = out["SRC_TELEFON"].apply(format_phone)
    out["Mennyiség (kg)"] = pd.to_numeric(out["OSSZES_MENNYISEG"], errors="coerce").fillna(0).round(2)
    out["Mennyiség (t)"] = (pd.to_numeric(out["OSSZES_MENNYISEG"], errors="coerce").fillna(0) / 1000).round(1)
    out["Veszélyes"] = out["HULLADEK_VESZELYES_10"].apply(
        lambda x: "Igen" if pd.to_numeric(x, errors="coerce") == 1 else "Nem"
    )
    out = pd.concat([out, parts], axis=1)
    out.rename(columns={"Vármegye": "Megye"}, inplace=True)

    out = out[[
        "CEGNEV", "Megye", "IRSZ", "Település", "Cím",
        "SRC_EMAIL", "SRC_TELEFON", "ADOSZAM", "SOURCE_KUJ",
        "EV", "ADATTIPUS_LABEL", "HULLADEKKOD", "HULLADEK_MEGNEVEZES",
        "Veszélyes", "Mennyiség (kg)", "Mennyiség (t)"
    ]].copy()

    out.rename(columns={
        "CEGNEV": "Cégnév", "SRC_EMAIL": "Email", "SRC_TELEFON": "Telefonszám",
        "ADOSZAM": "Adószám", "SOURCE_KUJ": "KÜJ",
        "EV": "Év",
        "ADATTIPUS_LABEL": "Adattípus",
        "HULLADEKKOD": "Hulladékkód", "HULLADEK_MEGNEVEZES": "Hulladék megnevezés",
    }, inplace=True)

    return out.sort_values(
        ["Cégnév", "Év", "Hulladékkód"],
        ascending=[True, False, True]
    ).reset_index(drop=True)


def build_detail_pivot_export(master: pd.DataFrame) -> pd.DataFrame:
    if master is None or master.empty:
        return pd.DataFrame(columns=["Cég", "Év", "Termékcsoport", "HAK", "HAK megnevezés", "Mennyiség (t)"])

    df = master.copy()
    df["Mennyiség (t)"] = pd.to_numeric(df["OSSZES_MENNYISEG"], errors="coerce").fillna(0) / 1000

    group_cols = ["CEGNEV", "EV", "PRODUCT_GROUP", "HULLADEKKOD"]
    if "HULLADEK_MEGNEVEZES" in df.columns:
        group_cols.append("HULLADEK_MEGNEVEZES")

    grouped = (
        df.groupby(group_cols, as_index=False)["Mennyiség (t)"]
        .sum()
        .sort_values(group_cols)
        .reset_index(drop=True)
    )

    rename_map = {
        "CEGNEV": "Cég",
        "EV": "Év",
        "PRODUCT_GROUP": "Termékcsoport",
        "HULLADEKKOD": "HAK",
        "HULLADEK_MEGNEVEZES": "HAK megnevezés",
    }
    grouped.rename(columns=rename_map, inplace=True)

    final_cols = ["Cég", "Év", "Termékcsoport", "HAK"]
    if "HAK megnevezés" in grouped.columns:
        final_cols.append("HAK megnevezés")
    final_cols.append("Mennyiség (t)")

    return grouped[final_cols]


def build_outline_report_sheet(ws, grouped_df: pd.DataFrame):
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

    thin = Side(style="thin", color="BFBFBF")
    bold_font = Font(bold=True)
    fill_company = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
    fill_year = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")
    fill_group = PatternFill(start_color="EDEDED", end_color="EDEDED", fill_type="solid")

    has_megnevezes = grouped_df is not None and "HAK megnevezés" in grouped_df.columns

    if has_megnevezes:
        headers = ["Szint", "Megnevezés", "Év", "HAK", "HAK megnevezés", "Mennyiség (t)"]
        COL_HAK_MEGNEV = 5
        COL_QTY = 6
        filter_ref = "A1:F1"
    else:
        headers = ["Szint", "Megnevezés", "Év", "HAK", "Mennyiség (t)"]
        COL_HAK_MEGNEV = None
        COL_QTY = 5
        filter_ref = "A1:E1"

    ws.sheet_properties.outlinePr.summaryBelow = True

    for col_idx, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=col_idx, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color="404040", end_color="404040", fill_type="solid")
        c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        c.alignment = Alignment(horizontal="center", vertical="center")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = filter_ref

    row_idx = 2

    if grouped_df is None or grouped_df.empty:
        return

    def _border_row(r: int):
        for c in range(1, COL_QTY + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            if cell.data_type == "f":
                cell.data_type = "s"

    for company, df_company in grouped_df.groupby("Cég", sort=True):
        company_total = round(float(df_company["Mennyiség (t)"].sum()), 1)

        ws.cell(row=row_idx, column=1, value="Cég")
        ws.cell(row=row_idx, column=2, value=company)
        ws.cell(row=row_idx, column=COL_QTY, value=company_total)
        for c in range(1, COL_QTY + 1):
            ws.cell(row=row_idx, column=c).font = bold_font
            ws.cell(row=row_idx, column=c).fill = fill_company
        _border_row(row_idx)

        company_row = row_idx
        row_idx += 1
        company_start = row_idx

        for year, df_year in df_company.groupby("Év", sort=True):
            year_total = round(float(df_year["Mennyiség (t)"].sum()), 1)

            ws.cell(row=row_idx, column=1, value="Év")
            ws.cell(row=row_idx, column=2, value=company)
            ws.cell(row=row_idx, column=3, value=int(year) if pd.notna(year) else "")
            ws.cell(row=row_idx, column=COL_QTY, value=year_total)
            for c in range(1, COL_QTY + 1):
                ws.cell(row=row_idx, column=c).fill = fill_year
            _border_row(row_idx)

            year_row = row_idx
            row_idx += 1
            year_start = row_idx

            for group, df_group in df_year.groupby("Termékcsoport", sort=True):
                group_total = round(float(df_group["Mennyiség (t)"].sum()), 1)

                ws.cell(row=row_idx, column=1, value="Termékcsoport")
                ws.cell(row=row_idx, column=2, value=group)
                ws.cell(row=row_idx, column=3, value=int(year) if pd.notna(year) else "")
                ws.cell(row=row_idx, column=COL_QTY, value=group_total)
                for c in range(1, COL_QTY + 1):
                    ws.cell(row=row_idx, column=c).fill = fill_group
                _border_row(row_idx)

                group_row = row_idx
                row_idx += 1
                group_start = row_idx

                for _, r in df_group.iterrows():
                    ws.cell(row=row_idx, column=1, value="HAK")
                    ws.cell(row=row_idx, column=2, value=r["Termékcsoport"])
                    ws.cell(row=row_idx, column=3, value=int(r["Év"]) if pd.notna(r["Év"]) else "")
                    ws.cell(row=row_idx, column=4, value=r["HAK"])
                    if COL_HAK_MEGNEV is not None:
                        ws.cell(row=row_idx, column=COL_HAK_MEGNEV, value=r.get("HAK megnevezés", ""))
                    ws.cell(row=row_idx, column=COL_QTY, value=round(float(r["Mennyiség (t)"]), 1))
                    _border_row(row_idx)
                    ws.row_dimensions[row_idx].outlineLevel = 3
                    row_idx += 1

                if group_start <= row_idx - 1:
                    for rr in range(group_start, row_idx):
                        ws.row_dimensions[rr].hidden = True
                    ws.row_dimensions[group_row].outlineLevel = 2

            if year_start <= row_idx - 1:
                for rr in range(year_start, row_idx):
                    ws.row_dimensions[rr].hidden = True
                ws.row_dimensions[year_row].outlineLevel = 1

        if company_start <= row_idx - 1:
            for rr in range(company_start, row_idx):
                ws.row_dimensions[rr].hidden = True
            ws.row_dimensions[company_row].outlineLevel = 0

    tonna_format = '#,##0.0\ "t"'
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=COL_QTY, max_col=COL_QTY):
        for cell in row:
            cell.number_format = tonna_format

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 16
    if has_megnevezes:
        ws.column_dimensions["E"].width = 45
        ws.column_dimensions["F"].width = 18
    else:
        ws.column_dimensions["E"].width = 18


def build_summary_export(master: pd.DataFrame, source_df: pd.DataFrame, groups: List[str]) -> pd.DataFrame:
    if master is None or master.empty:
        return pd.DataFrame()

    src_cols = [
        "CEGNEV", "KUJ", "SZEKHELY", "ADOSZAM",
        "SRC_EMAIL", "SRC_TELEFON", "SRC_SAJAT_TOKE", "SRC_NETTO_ARBEV"
    ]
    src = source_df.copy()
    for c in src_cols:
        if c not in src.columns:
            src[c] = ""
    src = src[src_cols].drop_duplicates(subset=["CEGNEV", "KUJ"])

    df = master.copy()
    df["Mennyiség (t)"] = pd.to_numeric(df["OSSZES_MENNYISEG"], errors="coerce").fillna(0) / 1000

    pivot = (
        df.groupby(["CEGNEV", "SOURCE_KUJ", "EV", "PRODUCT_GROUP"], as_index=False)["Mennyiség (t)"]
        .sum()
        .pivot_table(
            index=["CEGNEV", "SOURCE_KUJ", "EV"],
            columns="PRODUCT_GROUP",
            values="Mennyiség (t)",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
    )

    total_base = unique_quantities(master).copy()
    total_base["Mennyiség (t)"] = pd.to_numeric(total_base["OSSZES_MENNYISEG"], errors="coerce").fillna(0) / 1000
    total_by_year = (
        total_base.groupby(["CEGNEV", "SOURCE_KUJ", "EV"], as_index=False)["Mennyiség (t)"]
        .sum()
        .rename(columns={"Mennyiség (t)": "Össz hulladék (tonna)"})
    )

    out = total_by_year.merge(pivot, on=["CEGNEV", "SOURCE_KUJ", "EV"], how="left")
    out = out.merge(src, how="left", left_on=["CEGNEV", "SOURCE_KUJ"], right_on=["CEGNEV", "KUJ"])

    parts = out["SZEKHELY"].apply(split_address_parts).apply(pd.Series)
    out = pd.concat([out, parts], axis=1)
    out.rename(columns={"Vármegye": "Megye"}, inplace=True)

    out["Telefonszám"] = out["SRC_TELEFON"].apply(format_phone)
    out["SRC_SAJAT_TOKE_NUM"] = parse_hu_numeric(out["SRC_SAJAT_TOKE"]).fillna(0).round(0).astype("Int64")
    out["SRC_NETTO_ARBEV_NUM"] = parse_hu_numeric(out["SRC_NETTO_ARBEV"]).fillna(0).round(0).astype("Int64")

    for g in groups:
        if g not in out.columns:
            out[g] = 0.0

    final_cols = [
        "CEGNEV", "Megye", "IRSZ", "Település", "Cím",
        "SRC_EMAIL", "Telefonszám", "ADOSZAM", "SOURCE_KUJ",
        "SRC_SAJAT_TOKE_NUM", "SRC_NETTO_ARBEV_NUM",
        "EV", "Össz hulladék (tonna)",
    ] + groups

    out = out[final_cols].copy()
    out.rename(columns={
        "CEGNEV": "Cégnév", "SRC_EMAIL": "Email", "ADOSZAM": "Adószám",
        "SOURCE_KUJ": "KÜJ", "SRC_SAJAT_TOKE_NUM": "Saját tőke",
        "SRC_NETTO_ARBEV_NUM": "Nettó árbevétel", "EV": "Év",
    }, inplace=True)

    numeric_cols = ["Össz hulladék (tonna)"] + groups
    for c in numeric_cols:
        if c in out.columns:
            out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0.0).round(1)

    return out.sort_values(["Cégnév", "Év"], ascending=[True, False]).reset_index(drop=True)


def build_ktj_export(df: pd.DataFrame, ktj_service, category_service, load_metadata=True) -> pd.DataFrame:
    """KTJ szerinti export DataFrame előkészítése – UI-független."""
    out = df.copy()

    if "KUJ" not in out.columns:
        out["KUJ"] = out["SOURCE_KUJ"] if "SOURCE_KUJ" in out.columns else ""
    if "KTJ" not in out.columns:
        out["KTJ"] = out["SOURCE_KTJ"] if "SOURCE_KTJ" in out.columns else ""

    out["KUJ"] = out["KUJ"].apply(clean_int_like)
    out["KTJ"] = out["KTJ"].apply(clean_int_like)

    if "HULLADEK_VESZELYES_10" in out.columns:
        out["Veszélyes"] = out["HULLADEK_VESZELYES_10"].apply(
            lambda x: "Igen" if pd.to_numeric(x, errors="coerce") == 1 else "Nem"
        )
    else:
        out["Veszélyes"] = ""

    if "PRODUCT_GROUP" not in out.columns and "HULLADEKKOD" in out.columns:
        out["PRODUCT_GROUP"] = out["HULLADEKKOD"].astype(str).apply(category_service.resolve_group)

    if "OSSZES_MENNYISEG" in out.columns:
        out["Összes mennyiség"] = out["OSSZES_MENNYISEG"].apply(lambda x: format_quantity_auto_unit(x, 2))
        qty_col = "OSSZES_MENNYISEG"
    elif "MENNYISEG" in out.columns:
        out["Összes mennyiség"] = out["MENNYISEG"].apply(lambda x: format_quantity_auto_unit(x, 2))
        qty_col = "MENNYISEG"
    else:
        out["Összes mennyiség"] = ""
        qty_col = None

    if load_metadata:
        ktj_service.ensure_metadata_loaded(out)
    out = ktj_service.enrich_with_metadata(out)

    export_cols = [
        "CEGNEV", "KUJ", "KTJ", "KTJ megnevezés", "Besorolás",
        "Irányítószám", "Helység", "Cím", "Megye", "Régió",
        "EV", "ADATTIPUS_LABEL", "PRODUCT_GROUP", "HULLADEKKOD",
        "HULLADEK_MEGNEVEZES", "Veszélyes",
    ]
    if qty_col:
        export_cols.append(qty_col)

    existing = [c for c in export_cols if c in out.columns]
    out = out[existing].copy()

    rename_map = {
        "CEGNEV": "Cég", "KUJ": "KÜJ", "EV": "Év",
        "ADATTIPUS_LABEL": "Adattípus", "PRODUCT_GROUP": "Termékcsoport",
        "HULLADEKKOD": "Hulladékkód", "HULLADEK_MEGNEVEZES": "Megnevezés",
    }
    if qty_col:
        rename_map[qty_col] = "Mennyiség (kg)"
    out.rename(columns=rename_map, inplace=True)

    sort_cols = [c for c in ["KÜJ", "KTJ", "Cég", "Év", "Termékcsoport", "Hulladékkód"] if c in out.columns]
    if sort_cols:
        out.sort_values(sort_cols, inplace=True, na_position="last")

    return out.reset_index(drop=True)


# ---------------------------------------------------------------------------
# Excel írók
# ---------------------------------------------------------------------------

def write_summary_excel(path: str, master: pd.DataFrame, source_df: pd.DataFrame, groups: List[str]) -> None:
    out = build_summary_export(master, source_df, groups)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="Összesítő")
        style_excel_worksheet(writer.book["Összesítő"])


def write_detail_excel(path: str, master: pd.DataFrame, source_df: pd.DataFrame) -> None:
    out = build_detail_export(master, source_df)
    pivot = build_detail_pivot_export(master)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="Részletes")
        style_excel_worksheet(writer.book["Részletes"])
        ws = writer.book.create_sheet("Kimutatás")
        build_outline_report_sheet(ws, pivot)


def write_ktj_excel(path: str, df: pd.DataFrame, ktj_service, category_service, load_metadata=True) -> None:
    out = build_ktj_export(df, ktj_service, category_service, load_metadata=load_metadata)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="KTJ szerint")
        style_excel_worksheet(writer.book["KTJ szerint"])


# ---------------------------------------------------------------------------
# Stílusok
# ---------------------------------------------------------------------------

def style_excel_worksheet(ws):
    from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

    header_fill = PatternFill(start_color="404040", end_color="404040", fill_type="solid")
    sub_fill_odd = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    sub_fill_even = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    thin = Side(style="thin", color="BFBFBF")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        fill = sub_fill_odd if row_idx % 2 == 0 else sub_fill_even
        for cell in row:
            if cell.data_type == "f":
                cell.data_type = "s"
            cell.fill = fill
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    for col in ws.columns:
        letter = col[0].column_letter
        max_len = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 38)

    tonna_format = '#,##0.0\ "t"'
    for cell in ws[1]:
        if cell.value and ("(t)" in str(cell.value) or "tonna" in str(cell.value).lower()):
            col_letter = cell.column_letter
            for data_cell in ws[col_letter][1:]:
                data_cell.number_format = tonna_format