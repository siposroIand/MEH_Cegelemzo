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
            export_menu = window.menuBar().actions()[1].menu()
            actions = [action.text() for action in export_menu.actions()]
            self.assertEqual(actions[0], "Excel készítése…")
            for action in export_menu.actions():
                if action.menu():
                    actions.extend(child.text() for child in action.menu().actions())
            self.assertTrue(any("mátrix" in action for action in actions))
            self.assertFalse(any("CSV" in action for action in actions))
            for label in ("PKG", "JSON", "Teljes adatállomány"):
                self.assertTrue(any(label in action for action in actions))
            for name in ("import_button", "load_button", "matrix_button"):
                self.assertFalse(hasattr(window, name))
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

    def test_configured_export_captures_filters_and_creates_county_files(self):
        from openpyxl import load_workbook
        from PySide6.QtWidgets import QDialog
        from services.aggregation_service import aggregate_data
        from services.configurable_export_service import ExcelExportOptions
        from tests.fixtures import dataset
        from ui.main_window import MainWindow

        window = MainWindow()
        try:
            window.source_df, window.raw_df, window.missing_df = dataset()
            window.agg = aggregate_data(window.raw_df, window.source_df, window.category_service)
            window.populate_filters()
            window.filter_year.set_checked_by_user_data({2025})
            options = ExcelExportOptions(sections=("raw",), separate_counties=True)
            with patch("ui.main_window.ExcelExportDialog") as dialog_class, \
                    patch("ui.main_window.QFileDialog.getExistingDirectory", return_value=self.temp.name), \
                    patch("ui.main_window.QMessageBox.information") as success:
                dialog_class.return_value.exec.return_value = QDialog.Accepted
                dialog_class.return_value.options.return_value = options
                window.export_configured()
                self.assertIsNotNone(window._task_worker)
                self.assertFalse(window.menuBar().isEnabled())
                window.reset_filters()
                self.wait_for(lambda: window._task_worker is None)
                success.assert_called_once()
            files = list(Path(self.temp.name).glob("meh-excel-*/*.xlsx"))
            self.assertEqual(len(files), 2)
            for file in files:
                book = load_workbook(file)
                rows = list(book["Nyers adatok (kg)"].values)
                self.assertEqual(len(rows), 21)
                year_column = rows[0].index("EV")
                self.assertEqual({row[year_column] for row in rows[1:]}, {2025})
                book.close()
        finally:
            window.close()

    def test_configured_export_cancellation_and_empty_selection_error(self):
        from PySide6.QtWidgets import QDialog
        from services.aggregation_service import aggregate_data
        from services.configurable_export_service import ExcelExportOptions
        from tests.fixtures import dataset
        from ui.main_window import MainWindow

        window = MainWindow()
        try:
            window.source_df, window.raw_df, window.missing_df = dataset()
            window.agg = aggregate_data(window.raw_df, window.source_df, window.category_service)
            window.populate_filters()
            path = str(Path(self.temp.name) / "empty")
            with patch("ui.main_window.ExcelExportDialog") as dialog_class, \
                    patch("ui.main_window.QFileDialog.getSaveFileName", return_value=(path, "")) as chooser, \
                    patch("ui.main_window.QMessageBox.critical") as error:
                dialog = dialog_class.return_value
                dialog.exec.return_value = QDialog.Rejected
                window.export_configured()
                chooser.assert_not_called()
                self.assertIsNone(window._task_worker)
                dialog.exec.return_value = QDialog.Accepted
                dialog.options.return_value = ExcelExportOptions(counties=("Nincs ilyen megye",))
                window.export_configured()
                self.wait_for(lambda: window._task_worker is None)
                error.assert_called_once()
                self.assertIn("Nincs", error.call_args.args[-1])
                self.assertFalse(Path(path + ".xlsx").exists())
        finally:
            window.close()

    def test_export_dialog_requires_a_sheet_and_preserves_options(self):
        from PySide6.QtWidgets import QDialogButtonBox
        from ui.widgets import ExcelExportDialog

        dialog = ExcelExportDialog(["Pest", "Ismeretlen megye"], "Év: 2025")
        try:
            save = dialog.buttons.button(QDialogButtonBox.Save)
            self.assertTrue(save.isEnabled())
            dialog.section_checks["matrix"].setChecked(False)
            self.assertFalse(save.isEnabled())
            dialog.section_checks["raw"].setChecked(True)
            dialog.counties.set_checked_by_user_data({"Pest"})
            dialog.use_filters.setChecked(False)
            dialog.separate_counties.setChecked(True)
            options = dialog.options()
            self.assertTrue(save.isEnabled())
            self.assertEqual(options.sections, ("raw",))
            self.assertEqual(options.counties, ("Pest",))
            self.assertFalse(options.use_filters)
            self.assertTrue(options.separate_counties)
        finally:
            dialog.close()

    def test_responsive_dashboard_and_linear_filter_population(self):
        from ui.main_window import MainWindow
        from ui.widgets import MultiSelectComboBox

        combo = MultiSelectComboBox()
        with patch.object(combo, "_update_summary", wraps=combo._update_summary) as summary:
            combo.add_check_items((f"Cég {index}", index) for index in range(2000))
            summary.assert_called_once()
        self.assertEqual(combo.count(), 2000)
        combo.set_checked_by_user_data({0, 1999})
        self.assertEqual(combo.checked_values(), [0, 1999])
        window = MainWindow()
        try:
            window.show()
            window.resize(1600, 980)
            self.app.processEvents()
            self.assertEqual(window._dashboard_columns, (5, 2))
            window.resize(860, 600)
            self.app.processEvents()
            self.assertEqual(window.width(), 860)
            self.assertEqual(window._dashboard_columns, (2, 1))
            self.assertEqual(window.chart_grid.count(), 6)
            self.assertEqual(window.kpi_grid.count(), 5)
            self.assertEqual(window.dashboard_scroll.horizontalScrollBar().maximum(), 0)
            window.resize(1600, 980)
            self.app.processEvents()
            self.assertEqual(window._dashboard_columns, (5, 2))
            self.assertEqual(window.chart_grid.count(), 6)
        finally:
            window.close()

    def test_matrix_menu_exports_current_filters_in_background(self):
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
                window.matrix_action.trigger()
                self.assertIsNotNone(window._task_worker)
                self.assertFalse(window.menuBar().isEnabled())
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
            self.assertTrue(window.menuBar().isEnabled())
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

    def test_portable_exports_use_full_dataset_or_captured_filters(self):
        from openpyxl import load_workbook
        from services.aggregation_service import aggregate_data
        from services.package_service import read_package
        from services.report_service import read_dataset
        from tests.fixtures import dataset
        from ui.main_window import MainWindow

        window = MainWindow()
        try:
            window.source_df, window.raw_df, window.missing_df = dataset()
            window.agg = aggregate_data(window.raw_df, window.source_df, window.category_service)
            window.populate_filters()
            for kind, suffix in (("pkg", ".okirpkg"), ("json", ".json"), ("dataset", ".xlsx")):
                window.filter_year.set_checked_by_user_data({2025})
                window.filter_company.set_checked_by_user_data({"Minta Fém Kft."})
                window.filter_code.setText("170405")
                path = Path(self.temp.name) / kind
                with patch("ui.main_window.QFileDialog.getSaveFileName", return_value=(str(path), "")), \
                        patch("ui.main_window.QMessageBox.information") as success:
                    window.export_portable(kind)
                    self.assertIsNotNone(window._task_worker)
                    window.reset_filters()
                    self.wait_for(lambda: window._task_worker is None)
                    success.assert_called_once()
                path = path.with_suffix(suffix)
                self.assertTrue(path.exists())
                if kind == "pkg":
                    result = read_package(path.read_bytes())
                    self.assertEqual(len(result[1]), 120)
                    self.assertEqual(result[3]["selected_years"], [2025])
                    self.assertEqual(result[3]["filter_code"], "170405")
                elif kind == "json":
                    self.assertEqual(len(read_dataset(path.read_bytes())[1]), 120)
                elif kind == "dataset":
                    self.assertEqual(load_workbook(path)["Nyers adatok (kg)"].max_row, 121)
        finally:
            window.close()

    def test_package_import_restores_state_metadata_and_persistent_snapshot(self):
        from services.package_service import write_package
        from tests.fixtures import dataset
        from tests.test_portable import legacy_package
        from ui.main_window import MainWindow
        from PySide6.QtWidgets import QMessageBox

        window = MainWindow()
        self.app.processEvents()
        tables = dataset()
        state = {"selected_years": [2025], "filter_code": "170405"}
        metadata = {"1000000": {"MEGNEVEZES": "Mentett telephely"}}
        try:
            for legacy in (False, True):
                path = Path(self.temp.name) / ("legacy.okirpkg" if legacy else "modern.okirpkg")
                if legacy:
                    path.write_bytes(legacy_package(tables, state, metadata))
                    with patch("ui.main_window.QFileDialog.getOpenFileName", return_value=(str(path), "")), \
                            patch("ui.main_window.QMessageBox.question", return_value=QMessageBox.No):
                        window.import_shared_package()
                        self.assertIsNone(window._task_worker)
                else:
                    write_package(path, *tables, state, metadata)
                with patch("ui.main_window.QFileDialog.getOpenFileName", return_value=(str(path), "")), \
                        patch("ui.main_window.QMessageBox.question", return_value=QMessageBox.Yes), \
                        patch("ui.main_window.QMessageBox.information"), \
                        patch("ui.main_window.QMessageBox.critical") as failure:
                    window.import_shared_package()
                    self.wait_for(lambda: window._task_worker is None)
                    failure.assert_not_called()
                self.assertEqual(len(window.raw_df), 120)
                self.assertEqual(window.current_filters().years, (2025,))
                self.assertEqual(window.filter_code.text(), "170405")
                self.assertEqual(window.ktj_service.cache, metadata)
                sid = int(window.persistence.get_setting("desktop_active_snapshot"))
                self.assertEqual(len(window.persistence.load_snapshot(sid)[1]), 120)
                self.assertEqual(window.persistence.load_ktj_meta("1000000"), metadata["1000000"])
        finally:
            window.close()

    def test_invalid_package_preserves_active_data(self):
        import zipfile
        from tests.fixtures import dataset
        from ui.main_window import MainWindow

        window = MainWindow()
        self.app.processEvents()
        window.source_df, window.raw_df, window.missing_df = dataset()
        original = window.raw_df
        path = Path(self.temp.name) / "bad.okirpkg"
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("dataset.json", '{"format":"invalid"}')
        try:
            with patch("ui.main_window.QFileDialog.getOpenFileName", return_value=(str(path), "")), \
                    patch("ui.main_window.QMessageBox.critical") as failure:
                window.import_shared_package()
                self.wait_for(lambda: window._task_worker is None)
                failure.assert_called_once()
            self.assertIs(window.raw_df, original)
            self.assertEqual(window.persistence.list_snapshots(), [])
        finally:
            window.close()


if __name__ == "__main__":
    unittest.main()
