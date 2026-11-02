import os
from pathlib import Path
import tempfile
import threading
import time
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QCloseEvent
from PySide6.QtWidgets import QApplication


class DesktopTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.db_path = Path(self.temp.name) / "cache.sqlite"
        for target in ("ui.main_window.CACHE_DB", "workers.fetch_worker.CACHE_DB"):
            patcher = patch(target, self.db_path)
            patcher.start()
            self.addCleanup(patcher.stop)

    def wait_for(self, condition, timeout=10):
        deadline = time.monotonic() + timeout
        while not condition() and time.monotonic() < deadline:
            self.app.processEvents()
            time.sleep(.005)
        self.assertTrue(condition(), "Az asztali háttérművelet nem fejeződött be.")

    def test_table_model_and_window(self):
        import pandas as pd
        from ui.table_model import DataFrameModel
        from ui.main_window import MainWindow
        from tests.fixtures import dataset
        from services.aggregation_service import aggregate_data

        model = DataFrameModel(pd.DataFrame({"kg": [10, 2, 30]}))
        model.sort(0, Qt.AscendingOrder)
        self.assertEqual(model.data(model.index(0, 0)), "2")
        mixed = DataFrameModel(pd.DataFrame({"KTJ": [None, "003", 2]}))
        mixed.sort(0)
        self.assertEqual(mixed.data(mixed.index(0, 0)), "003")
        self.assertEqual(mixed.data(mixed.index(2, 0)), "")
        window = MainWindow()
        try:
            window.source_df, window.raw_df, window.missing_df = dataset()
            window.agg = aggregate_data(window.raw_df, window.source_df, window.category_service)
            window.populate_filters()
            window.refresh_all_views()
            window.tabs.setCurrentIndex(1)
            self.assertGreater(window.data_year_tabs.count(), 1)
            window.tabs.setCurrentIndex(3)
            self.assertEqual(window.missing_table.model().rowCount(), 1)
            actions = [action.text() for action in window.menuBar().actions()[1].menu().actions()]
            self.assertTrue(any("mátrix" in action for action in actions))
            self.assertFalse(any("CSV" in action or "JSON" in action for action in actions))
            window.apply_default_year_selection()
            self.assertTrue(window.current_filters().years)
            window.reset_filters()
            self.assertFalse(window.current_filters().years)
            window.filter_type.set_checked_by_user_data({"HULL_ATADAS"})
            self.assertEqual(window.current_filters().types, ("HULL_ATADAS",))
        finally:
            window.close()

    def test_worker_delegates_to_shared_service(self):
        from workers.fetch_worker import FetchWorker
        from services.fetch_service import FetchResult
        from tests.fixtures import dataset
        source, raw, missing = dataset()
        received = []
        worker = FetchWorker(source, True, False, ["HULL_ATADAS"])
        worker.finished_signal.connect(lambda *args: received.append(args))
        with patch("workers.fetch_worker.fetch_data", return_value=FetchResult(source, raw, missing, 2026)):
            worker.run()
        self.assertEqual(len(received), 1)
        self.assertEqual(len(received[0][1]), 120)
        from services.persistence_service import PersistenceService
        persistence = PersistenceService(self.db_path)
        try:
            self.assertEqual(len(persistence.list_snapshots()), 1)
            self.assertEqual(persistence.get_setting("desktop_active_snapshot"), "1")
        finally:
            persistence.close()

    def test_matrix_button_exports_current_filters_in_background(self):
        from openpyxl import load_workbook
        from ui.main_window import MainWindow
        from tests.fixtures import dataset
        from services.aggregation_service import aggregate_data

        window = MainWindow()
        path = str(Path(self.temp.name) / "matrix.xlsx")
        try:
            window.source_df, window.raw_df, window.missing_df = dataset()
            window.agg = aggregate_data(window.raw_df, window.source_df, window.category_service)
            window.populate_filters()
            window.filter_year.set_checked_by_user_data({2025})
            window.filter_code.setText("170405")
            window.filter_type.set_checked_by_user_data({"HULL_ATADAS"})
            with patch("ui.main_window.QFileDialog.getSaveFileName", return_value=(path, "")), \
                    patch("ui.main_window.QMessageBox.information") as success:
                window.matrix_button.click()
                self.assertIsNotNone(window._task_worker)
                self.assertFalse(window.matrix_button.isEnabled())
                self.wait_for(lambda: window._task_worker is None)
                success.assert_called_once()
            book = load_workbook(path)
            sheet = book["Cég–Év–HAK mátrix"]
            self.assertEqual(sheet["B1"].value, 2025)
            self.assertEqual(sheet["B2"].value, "170405")
            total = sum(row[-1] for row in list(sheet.values)[3:])
            expected = window.get_filtered_base_unique().OSSZES_MENNYISEG.sum() / 1000
            self.assertAlmostEqual(total, expected)
            self.assertIn("2025", book["Szűrés"]["B2"].value)
        finally:
            window.close()

    def test_background_task_keeps_event_loop_alive_and_blocks_close(self):
        from ui.main_window import MainWindow

        window = MainWindow()
        release = threading.Event()
        ticks = []
        results = []
        timer = QTimer()
        timer.timeout.connect(lambda: ticks.append(1))
        timer.start(5)
        try:
            window._run_task("Teszt", lambda: release.wait(5), results.append)
            with patch("ui.main_window.QMessageBox.warning") as warning:
                close_event = QCloseEvent()
                window.closeEvent(close_event)
                self.assertFalse(close_event.isAccepted())
                warning.assert_called_once()
            self.wait_for(lambda: len(ticks) >= 3)
            self.assertTrue(window._is_busy())
            release.set()
            self.wait_for(lambda: window._task_worker is None)
            self.assertEqual(results, [True])
            self.assertTrue(window.matrix_button.isEnabled())
        finally:
            release.set()
            self.wait_for(lambda: window._task_worker is None)
            timer.stop()
            window.close()

    def test_failed_export_preserves_existing_file_and_shows_error(self):
        from ui.main_window import MainWindow

        window = MainWindow()
        path = Path(self.temp.name) / "matrix.xlsx"
        path.write_bytes(b"existing workbook")

        def broken_writer(temporary):
            Path(temporary).write_bytes(b"partial")
            raise OSError("Nincs elegendő hely")

        try:
            with patch("ui.main_window.QFileDialog.getSaveFileName", return_value=(str(path), "")), \
                    patch("ui.main_window.QMessageBox.critical") as error, \
                    patch("ui.main_window.QMessageBox.information") as success:
                window._export_excel("matrix.xlsx", broken_writer)
                self.wait_for(lambda: window._task_worker is None)
                error.assert_called_once()
                self.assertIn("Nincs elegendő hely", error.call_args.args[-1])
                success.assert_not_called()
            self.assertEqual(path.read_bytes(), b"existing workbook")
            self.assertEqual(list(path.parent.glob(".matrix-*")), [])
        finally:
            window.close()

    def test_last_saved_session_restores_on_startup(self):
        from services.persistence_service import PersistenceService
        from tests.fixtures import dataset
        from ui.main_window import MainWindow

        persistence = PersistenceService(self.db_path)
        source, raw, missing = dataset()
        sid = persistence.save_snapshot(source, raw, missing, "minta.xlsx")
        persistence.set_setting("web_active_snapshot", str(sid))
        persistence.close()
        window = MainWindow()
        try:
            self.wait_for(lambda: not window.raw_df.empty and window._task_worker is None)
            self.assertEqual(len(window.raw_df), len(raw))
            self.assertEqual(window.current_file_label.text(), "minta.xlsx")
        finally:
            window.close()


if __name__ == "__main__":
    unittest.main()
