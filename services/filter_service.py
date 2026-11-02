import re
from dataclasses import dataclass

import pandas as pd


@dataclass(frozen=True)
class Filters:
    companies: tuple = ()
    years: tuple = ()
    categories: tuple = ()
    code: str = ""
    types: tuple = ()


def filter_frame(df: pd.DataFrame, filters: Filters, category_service=None) -> pd.DataFrame:
    if df.empty:
        return df
    mask = pd.Series(True, index=df.index)
    for column, values in (
        ("CEGNEV", filters.companies),
        ("EV", filters.years),
        ("ADATTIPUS", filters.types),
    ):
        if values:
            mask &= df[column].isin(values)
    if filters.categories:
        if "PRODUCT_GROUP" in df:
            mask &= df["PRODUCT_GROUP"].isin(filters.categories)
        else:
            if category_service is None:
                raise ValueError("A nyers adatok szűréséhez kategóriaszolgáltatás szükséges.")
            selected = set(filters.categories)
            matches = {
                code: bool(selected.intersection(category_service.resolve_groups(code)))
                for code in df["HULLADEKKOD"].unique()
            }
            mask &= df["HULLADEKKOD"].map(matches)
    codes = [re.escape(code.strip()) for code in filters.code.split(";") if code.strip()]
    if codes:
        mask &= df["HULLADEKKOD"].astype(str).str.contains("|".join(codes), case=False, na=False)
    return df.loc[mask]


def unique_quantities(master: pd.DataFrame) -> pd.DataFrame:
    """Only remove duplicates introduced by overlapping product groups."""
    return master.drop(columns=["PRODUCT_GROUP"], errors="ignore").drop_duplicates()
