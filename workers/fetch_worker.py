import pandas as pd
from PySide6.QtCore import QThread, Signal

from core.constants import CACHE_DB, DEFAULT_MAX_WORKERS, PRODUCT_GROUPS_JSON
from core.logger import logger
from services.aggregation_service import aggregate_data
from services.category_service import CategoryService
from services.fetch_service import fetch_data
from services.import_service import build_source_from_excel_by_adoszam
from services.okir_client import OkirClient
from services.persistence_service import PersistenceService


class FetchWorker(QThread):
    progress_signal = Signal(int, int, str)
    finished_signal = Signal(object, object, object, object, object)
    error_signal = Signal(str)

    def __init__(self, source_df, use_cache, force_refresh, selected_types,
                 max_workers=DEFAULT_MAX_WORKERS, file_path=None):
        super().__init__()
        self.source_df = source_df
        self.use_cache = use_cache
        self.force_refresh = force_refresh
        self.selected_types = selected_types
        self.max_workers = max_workers
        self.file_path = file_path

    def run(self):
        try:
            missing = []
            if self.file_path:
                client = OkirClient()
                try:
                    self.source_df = build_source_from_excel_by_adoszam(
                        self.file_path, client, log_fn=logger.write,
                        progress_fn=self.progress_signal.emit, missing_fn=missing.append,
                    )
                finally:
                    client.session.close()
            if self.source_df.empty:
                raise ValueError("Nem sikerült forrásadatot előállítani. " +
                                 "; ".join(row["Ok"] for row in missing[:5]))
            result = fetch_data(
                self.source_df, self.use_cache, self.force_refresh,
                self.selected_types, self.max_workers, self.progress_signal.emit,
            )
            failures = pd.concat([pd.DataFrame(missing), result.missing], ignore_index=True)
            self.progress_signal.emit(0, 1, "Aggregálás és automatikus helyi mentés...")
            aggregates = aggregate_data(result.raw, result.source, CategoryService(PRODUCT_GROUPS_JSON))
            persistence = PersistenceService(CACHE_DB)
            try:
                sid = persistence.save_snapshot(
                    result.source, result.raw, failures, self.file_path or "", "Automatikus asztali mentés",
                )
                persistence.set_setting("desktop_active_snapshot", str(sid))
            finally:
                persistence.close()
            self.finished_signal.emit(result.source, result.raw, failures, result.max_year, aggregates)
        except Exception as exc:
            logger.write(f"Lekérési hiba: {exc}")
            self.error_signal.emit(str(exc))
