from typing import Dict, Optional
import pandas as pd

from core.utils import clean_int_like
from services.okir_client import OkirClient
from services.persistence_service import PersistenceService


class KtjService:
    def __init__(self, persistence: PersistenceService):
        self.persistence = persistence
        self.cache: Dict[str, dict] = {}

    def ensure_metadata_loaded(self, df: pd.DataFrame) -> None:
        if df is None or df.empty or "KTJ" not in df.columns:
            return

        unique_ktjs = []
        for val in df["KTJ"].dropna().tolist():
            ktj = clean_int_like(val)
            if ktj and ktj not in unique_ktjs:
                unique_ktjs.append(ktj)

        if not unique_ktjs:
            return

        client = OkirClient()
        for ktj in unique_ktjs:
            if ktj in self.cache:
                continue
            cached = None
            try:
                cached = self.persistence.load_ktj_meta(ktj)
            except Exception:
                cached = None
            if cached is not None:
                self.cache[ktj] = cached
                continue
            try:
                payload = client.fetch_first_ktj_detail(ktj) or {}
                self.cache[ktj] = payload
                try:
                    self.persistence.save_ktj_meta(ktj, payload)
                except Exception:
                    pass
            except Exception:
                self.cache[ktj] = {}

    def enrich_with_metadata(self, df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty:
            return df

        out = df.copy()
        if "KTJ" not in out.columns:
            out["KTJ"] = out["SOURCE_KTJ"] if "SOURCE_KTJ" in out.columns else ""
        out["KTJ"] = out["KTJ"].apply(clean_int_like)

        def meta_value(ktj_value, key):
            item = self.cache.get(clean_int_like(ktj_value), {})
            return item.get(key, "") if isinstance(item, dict) else ""

        out["KTJ megnevezés"] = out["KTJ"].apply(lambda x: meta_value(x, "MEGNEVEZES"))
        out["Besorolás"]      = out["KTJ"].apply(lambda x: meta_value(x, "BESOROLAS"))
        out["Irányítószám"]   = out["KTJ"].apply(lambda x: clean_int_like(meta_value(x, "IRSZAM")))
        out["Helység"]        = out["KTJ"].apply(lambda x: meta_value(x, "HELYSEG"))
        out["Cím"]            = out["KTJ"].apply(lambda x: meta_value(x, "CIM"))
        out["Megye"]          = out["KTJ"].apply(lambda x: meta_value(x, "MEGYE_NEV"))
        out["Régió"]          = out["KTJ"].apply(lambda x: meta_value(x, "REGIO_NEV"))
        return out

    def load_external_cache(self, data: dict) -> None:
        """Megosztott csomag betöltésekor az elmentett cache visszatöltéséhez."""
        if isinstance(data, dict):
            self.cache.update(data)