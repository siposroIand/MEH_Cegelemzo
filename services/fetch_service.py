import json
import logging
import threading
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass

import pandas as pd

from core.constants import ADATTIPUS_LABELS, CACHE_DB, DEFAULT_MAX_WORKERS
from services.aggregation_service import extract_records, normalize_records
from services.cache_db import CacheDB
from services.okir_client import OkirClient


@dataclass
class FetchResult:
    source: pd.DataFrame
    raw: pd.DataFrame
    missing: pd.DataFrame
    max_year: int | None


def fetch_data(source_df, use_cache, force_refresh, selected_types,
               max_workers=DEFAULT_MAX_WORKERS, progress=None, cache_path=CACHE_DB):
    if not selected_types or any(kind not in ADATTIPUS_LABELS for kind in selected_types):
        raise ValueError("Válassz érvényes adattípust.")
    tasks = [
        (str(row.CEGNEV).strip(), str(row.KUJ).strip(), kind)
        for row in source_df[["CEGNEV", "KUJ"]].drop_duplicates().itertuples(index=False)
        for kind in dict.fromkeys(selected_types)
    ]
    workers = min(6, max(1, int(max_workers)))
    lock = threading.Lock()
    completed = 0

    def partition_jobs(partition):
        nonlocal completed
        rows, missing = [], []
        cache = CacheDB(cache_path)
        client = OkirClient()
        try:
            for company, kuj, kind in partition:
                label = ADATTIPUS_LABELS[kind]
                try:
                    data = None
                    if use_cache and not force_refresh:
                        cached = cache.get(kuj, kind)
                        if cached and cached["status"] == "success" and cached["raw_json"]:
                            try:
                                data = json.loads(cached["raw_json"])
                            except json.JSONDecodeError:
                                logging.getLogger(__name__).warning("Sérült OKIR cache: %s / %s", kuj, kind)
                    if data is None:
                        data = client.fetch(kuj, kind)
                        if use_cache:
                            cache.save(kuj, kind, "success", json.dumps(data, ensure_ascii=False))
                    records = extract_records(data, kind)
                    if not records:
                        raise ValueError("Nincs feldolgozható rekord")
                    rows.extend(normalize_records(records, company, kuj, kind))
                except Exception as exc:
                    logging.getLogger(__name__).warning("OKIR lekérési hiba: %s / %s: %s", kuj, kind, exc)
                    missing.append({"Cég": company, "KÜJ": kuj, "Adattípus": label, "Ok": str(exc)})
                with lock:
                    completed += 1
                    if progress:
                        progress(completed, len(tasks), f"Lekérés: {company} / {label}")
        finally:
            cache.close()
            client.session.close()
        return rows, missing

    all_rows, all_missing = [], []
    with ThreadPoolExecutor(max_workers=workers) as executor:
        for rows, missing in executor.map(partition_jobs, [tasks[i::workers] for i in range(workers)]):
            all_rows.extend(rows)
            all_missing.extend(missing)
    raw = pd.DataFrame(all_rows)
    years = raw["EV"].dropna() if not raw.empty else pd.Series(dtype=float)
    return FetchResult(source_df, raw, pd.DataFrame(all_missing), int(years.max()) if not years.empty else None)
