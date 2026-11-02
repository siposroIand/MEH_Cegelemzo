import os
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

try:
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QApplication
except ImportError:
    QApplication = None


@unittest.skipIf(QApplication is None, "Desktop dependencies are optional for the web app")
class DesktopTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

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
            self.assertTrue(any("CSV" in action for action in actions))
            self.assertTrue(any("JSON" in action for action in actions))
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


if __name__ == "__main__":
    unittest.main()
