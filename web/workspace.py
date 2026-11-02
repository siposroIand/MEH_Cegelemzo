import logging
import threading
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from functools import lru_cache

import pandas as pd

from models.records import Aggregates
from services.aggregation_service import aggregate_data
from services.filter_service import filter_frame
from services.ktj_service import KtjService
from services.report_service import summary


@dataclass(frozen=True)
class Dataset:
    source: pd.DataFrame
    raw: pd.DataFrame
    missing: pd.DataFrame
    aggregates: Aggregates
    filename: str


class Workspace:
    """One password-protected shared workspace, served by one process."""

    def __init__(self, persistence, categories):
        self.persistence = persistence
        self.categories = categories
        self.ktj = KtjService(persistence)
        self.lock = threading.RLock()
        self.executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="meh-job")
        empty = pd.DataFrame()
        self.dataset = Dataset(empty, empty, empty, aggregate_data(empty, empty, categories), "")
        self.filtered = lru_cache(maxsize=8)(self._filtered)
        self.job = {"status": "idle", "step": 0, "total": 1, "message": "Készen áll."}
        snapshots = persistence.list_snapshots()
        if snapshots:
            active = persistence.get_setting("web_active_snapshot")
            selected = next((row for row in snapshots if str(row[0]) == active), snapshots[0])
            self.replace(*persistence.load_snapshot(selected[0]), selected[2], persist=False)

    def replace(self, source, raw, missing, filename, persist=True, metadata=None):
        aggregates = aggregate_data(raw, source, self.categories)
        if persist:
            for ktj, payload in (metadata or {}).items():
                self.persistence.save_ktj_meta(ktj, payload)
            sid = self.persistence.save_snapshot(source, raw, missing, filename, "Automatikus webes mentés")
            self.persistence.set_setting("web_active_snapshot", str(sid))
        with self.lock:
            self.dataset = Dataset(source, raw, missing, aggregates, filename)
            self.ktj.cache.update(metadata or {})
            self.filtered.cache_clear()

    def progress(self, step, total, message):
        with self.lock:
            self.job.update(step=step, total=max(total, 1), message=message)

    def submit(self, operation):
        with self.lock:
            if self.job["status"] == "running":
                raise RuntimeError("Már fut egy művelet. Várd meg a befejezését.")
            self.job = {"status": "running", "step": 0, "total": 1, "message": "Feldolgozás..."}

            def run():
                try:
                    message = operation()
                    with self.lock:
                        self.job.update(status="done", step=1, total=1, message=message or "Kész.")
                except Exception as exc:
                    logging.getLogger(__name__).exception("Webes háttérfeladat sikertelen")
                    with self.lock:
                        self.job.update(status="error", message=str(exc))
                finally:
                    self.persistence.close()

            self.executor.submit(run)

    def _filtered(self, filters):
        master = filter_frame(self.dataset.aggregates.master, filters)
        raw = filter_frame(self.dataset.raw, filters, self.categories)
        return master, raw, summary(master)

    def close(self):
        self.executor.shutdown(wait=True)
        self.filtered.cache_clear()
        self.persistence.close()
