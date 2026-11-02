from pathlib import Path
import zipfile

import pandas as pd
from PySide6.QtCore import QEvent, Qt, QTimer
from PySide6.QtGui import QAction
from PySide6.QtWidgets import (
    QFileDialog,
    QDialog,
    QFormLayout,
    QGroupBox,
    QGridLayout,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSplitter,
    QTabWidget,
    QTableView,
    QTextEdit,
    QVBoxLayout,
    QWidget,
    QHeaderView,
)

from core.constants import (
    CACHE_DB,
    DEFAULT_QUERY_TYPE_LABEL,
    ENDPOINTS,
    FILTER_DEBOUNCE_MS,
    LABEL_TO_ADATTIPUS,
    OUTPUT_DIR,
    PRODUCT_GROUPS_JSON,
)
from core.logger import logger
from core.utils import (
    add_quantity_columns,
    axis_int_formatter,
    axis_number_formatter,
    clean_int_like,
    convert_for_axis,
    format_hu_number,
    format_phone,
    format_quantity_auto_unit,
    parse_hu_numeric,
    safe_str,
    shorten_label,
    split_address_parts,
)
from services.aggregation_service import aggregate_data
from services.cache_db import CacheDB
from services.category_service import CategoryService
from services.export_service import (
    write_dataset_excel,
    write_detail_excel,
    write_hak_matrix_excel,
    write_ktj_excel,
    write_missing_excel,
    write_summary_excel,
)
from services.desktop_export_service import atomic_export
from services.configurable_export_service import available_counties, export_configured_excel
from services.filter_service import Filters, filter_frame, unique_quantities
from services.report_service import dataset_bytes, read_dataset, report_html
from services.package_service import read_packages, write_package
from services.ktj_service import KtjService
from services.persistence_service import PersistenceService
from ui.widgets import (
    ChartPanel,
    DashboardCard,
    ExcelExportDialog,
    FullscreenChartDialog,
    ImportDatabaseDialog,
    LogDialog,
    MultiSelectComboBox,
)
from workers.fetch_worker import FetchWorker
from workers.task_worker import TaskWorker
from ui.table_model import DataFrameModel


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MÉH Zrt. - Cégelemző")
        self.resize(1600, 980)
        self.setMinimumSize(860, 600)
        self._task_worker = None
        self._task_success = None

        self.source_df = pd.DataFrame()
        self.raw_df = pd.DataFrame()
        self.missing_df = pd.DataFrame()

        self.category_service = CategoryService(PRODUCT_GROUPS_JSON)
        self.agg = aggregate_data(pd.DataFrame(), pd.DataFrame(), self.category_service)
        self.max_data_year = None

        self.cache = CacheDB(CACHE_DB)
        self.persistence = PersistenceService(CACHE_DB)
        self.ktj_service = KtjService(self.persistence)

        self.current_file_path = ""
        self.current_query_label = DEFAULT_QUERY_TYPE_LABEL
        self.current_use_cache = True

        self._filtered_cache_key = None
        self._filtered_cache = {}
        self._drill_category = None
        self._ktj_confirmed = False
        self._dirty_views = set()

        self.category_colors = {
            "Vas": "#64748b",
            "Alu": "#3b82f6",
            "Réz": "#14b8a6",
            "Egyéb fémek": "#f59e0b",
            "Papír": "#f97316",
            "Műanyag": "#06b6d4",
            "Elektronika": "#8b5cf6",
            "Fa": "#ef4444",
            "Bontási építési": "#10b981",
            "Nem besorolható": "#84cc16",
        }

        self._build_ui()
        self._build_menu()

        self._filter_timer = QTimer(self)
        self._filter_timer.setSingleShot(True)
        self._filter_timer.timeout.connect(self.refresh_all_views)

        self._apply_styles()
        self.set_status("Készen áll.")
        QTimer.singleShot(0, self._restore_last_snapshot)

    # -------------------------------------------------------------------------
    # Helpers
    # -------------------------------------------------------------------------
    def set_status(self, text: str, progress_value=None, progress_max=None, log=False):
        self.status_label.setText(text)
        if progress_max is not None:
            self.progress.setMaximum(max(progress_max, 1))
        if progress_value is not None:
            self.progress.setValue(progress_value)
        if log:
            try:
                logger.write(text)
            except Exception:
                pass

    def selected_types_from_label(self, label: str):
        return list(ENDPOINTS.keys()) if label == "Mindhárom" else [LABEL_TO_ADATTIPUS[label]]

    def _filters_key(self):
        return (
            tuple(sorted(map(str, self.filter_company.checked_values()))),
            tuple(sorted(map(int, self.filter_year.checked_values()))),
            tuple(sorted(map(str, self.filter_category.checked_values()))),
            tuple(sorted(map(str, self.filter_type.checked_values()))),
            self.filter_code.text().strip(),
        )

    def _invalidate_filtered_cache(self):
        self._filtered_cache_key = None
        self._filtered_cache = {}

    def get_filtered_snapshot(self):
        key = self._filters_key()
        if key == self._filtered_cache_key and self._filtered_cache:
            return self._filtered_cache

        master = self.get_filtered_master()
        base = unique_quantities(master)

        yearly = pd.DataFrame()
        if base is not None and not base.empty and "EV" in base.columns and "OSSZES_MENNYISEG" in base.columns:
            tmp = base.copy()
            tmp["OSSZES_MENNYISEG"] = pd.to_numeric(tmp["OSSZES_MENNYISEG"], errors="coerce").fillna(0)
            yearly = tmp.groupby("EV", as_index=False)["OSSZES_MENNYISEG"].sum().sort_values("EV")

        category_totals = pd.DataFrame()
        if master is not None and not master.empty and "PRODUCT_GROUP" in master.columns and "OSSZES_MENNYISEG" in master.columns:
            tmp = master.copy()
            tmp["OSSZES_MENNYISEG"] = pd.to_numeric(tmp["OSSZES_MENNYISEG"], errors="coerce").fillna(0)
            category_totals = (
                tmp.groupby("PRODUCT_GROUP", as_index=False)["OSSZES_MENNYISEG"]
                .sum()
                .sort_values("OSSZES_MENNYISEG", ascending=False)
            )

        self._filtered_cache_key = key
        self._filtered_cache = {
            "master": master,
            "base": base,
            "yearly": yearly,
            "category_totals": category_totals,
        }
        return self._filtered_cache

    def capture_state(self):
        return {
            "version": 1,
            "current_file_path": self.current_file_path,
            "current_query_label": self.current_query_label,
            "current_use_cache": self.current_use_cache,
            "selected_companies": list(self.filter_company.checked_values()),
            "selected_years": list(self.filter_year.checked_values()),
            "selected_categories": list(self.filter_category.checked_values()),
            "selected_types": list(self.filter_type.checked_values()),
            "filter_code": self.filter_code.text().strip(),
            "current_tab_index": self.tabs.currentIndex(),
            "drill_category": self._drill_category,
        }

    def restore_state(self, state: dict):
        if not state:
            return

        self.current_file_path = state.get("current_file_path", "")
        self.current_query_label = state.get("current_query_label", DEFAULT_QUERY_TYPE_LABEL)
        self.current_use_cache = bool(state.get("current_use_cache", True))

        self.current_file_label.setText(self.current_file_path or "-")
        self.current_type_label.setText(self.current_query_label or "-")
        self.current_cache_label.setText("Igen" if self.current_use_cache else "Nem")

        self.populate_filters()

        companies = set(str(x) for x in state.get("selected_companies", []))
        years = set(int(x) for x in state.get("selected_years", []))
        categories = set(str(x) for x in state.get("selected_categories", []))

        try:
            self.filter_company.blockSignals(True)
            self.filter_year.blockSignals(True)
            self.filter_category.blockSignals(True)
            self.filter_type.blockSignals(True)
            self.filter_company.set_checked_by_user_data(companies)
            self.filter_year.set_checked_by_user_data(years)
            self.filter_category.set_checked_by_user_data(categories)
            self.filter_type.set_checked_by_user_data(set(state.get("selected_types", [])))
        finally:
            self.filter_company.blockSignals(False)
            self.filter_year.blockSignals(False)
            self.filter_category.blockSignals(False)
            self.filter_type.blockSignals(False)

        self.filter_code.setText(state.get("filter_code", ""))
        self._drill_category = state.get("drill_category")

        self._invalidate_filtered_cache()
        self.refresh_all_views()

        idx = int(state.get("current_tab_index", 0))
        idx = max(0, min(idx, self.tabs.count() - 1))
        self.tabs.setCurrentIndex(idx)

    # -------------------------------------------------------------------------
    # Menu
    # -------------------------------------------------------------------------
    def _build_menu(self):
        menubar = self.menuBar()

        import_menu = menubar.addMenu("Importálás")
        browse_new_action = QAction("Excel betöltése…", self)
        browse_new_action.setShortcut("Ctrl+O")
        browse_new_action.triggered.connect(self.open_import_dialog)
        import_menu.addAction(browse_new_action)

        load_snapshot_action = QAction("Korábbi helyi mentés betöltése…", self)
        load_snapshot_action.triggered.connect(self.load_snapshot)
        import_menu.addAction(load_snapshot_action)

        import_json_action = QAction("MÉH JSON mentés betöltése", self)
        import_json_action.triggered.connect(self.import_json)
        import_menu.addAction(import_json_action)
        package_action = QAction("Megosztható PKG-k betöltése (régi csomag is)", self)
        package_action.triggered.connect(self.import_shared_package)
        import_menu.addAction(package_action)

        export_menu = menubar.addMenu("Exportálás")
        configure_action = QAction("Excel készítése…", self)
        configure_action.setShortcut("Ctrl+Shift+E")
        configure_action.triggered.connect(self.export_configured)
        export_menu.addAction(configure_action)
        export_menu.addSeparator()

        excel_menu = export_menu.addMenu("Gyors Excel-kimutatások")
        self.matrix_action = QAction("Cég × Év / HAK mátrix", self)
        self.matrix_action.triggered.connect(self.export_hak_matrix)
        excel_menu.addAction(self.matrix_action)
        export_summary_action = QAction("Excel összesítő (termékcsoportonként)", self)
        export_summary_action.triggered.connect(self.export_summary_excel)
        excel_menu.addAction(export_summary_action)

        export_detail_action = QAction("Excel részletes (HAK szerint)", self)
        export_detail_action.triggered.connect(self.export_detail_excel)
        excel_menu.addAction(export_detail_action)

        export_ktj_action = QAction("Excel export (KTJ szerint)", self)
        export_ktj_action.triggered.connect(self.export_ktj_excel)
        excel_menu.addAction(export_ktj_action)
        missing_action = QAction("Kimaradt / hibás lekérések (Excel)", self)
        missing_action.triggered.connect(self.export_missing_excel)
        excel_menu.addAction(missing_action)
        dataset_action = QAction("Teljes adatállomány (szűréstől függetlenül)", self)
        dataset_action.triggered.connect(lambda: self.export_portable("dataset"))
        excel_menu.addSeparator()
        excel_menu.addAction(dataset_action)
        report_menu = export_menu.addMenu("Riportok")
        report_action = QAction("Vezetői riport (HTML / nyomtatás PDF-be)", self)
        report_action.triggered.connect(self.export_report)
        report_menu.addAction(report_action)
        export_menu.addSeparator()
        save_menu = export_menu.addMenu("Mentés és megosztás")
        save_snapshot_action = QAction("Helyi mentés készítése", self)
        save_snapshot_action.triggered.connect(self.save_snapshot)
        save_menu.addAction(save_snapshot_action)
        save_menu.addSeparator()
        for label, kind in (
            ("Megosztható PKG (teljes munkamenet)", "pkg"),
            ("MÉH JSON (teljes adatállomány, visszaimportálható)", "json"),
        ):
            action = QAction(label, self)
            action.triggered.connect(lambda checked=False, selected=kind: self.export_portable(selected))
            save_menu.addAction(action)

        program_menu = menubar.addMenu("Program")
        clear_cache_action = QAction("Gyorsítótár törlése", self)
        clear_cache_action.triggered.connect(self.clear_cache)
        program_menu.addAction(clear_cache_action)

        delete_snapshot_action = QAction("Mentések kezelése (törlés)", self)
        delete_snapshot_action.triggered.connect(self.delete_snapshot)
        program_menu.addAction(delete_snapshot_action)

        log_action = QAction("Műveleti napló", self)
        log_action.triggered.connect(lambda: LogDialog(self).exec())
        program_menu.addAction(log_action)

    # -------------------------------------------------------------------------
    # UI
    # -------------------------------------------------------------------------
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)

        root = QVBoxLayout(central)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(12)

        header = QHBoxLayout()
        title = QLabel("MÉH Zrt. - Cégelemző szoftver")
        title.setObjectName("appTitle")
        header.addWidget(title)
        header.addStretch()
        root.addLayout(header)

        self.status_label = QLabel("Készen áll.")
        self.status_label.setWordWrap(True)
        self.progress = QProgressBar()
        self.progress.setFixedSize(160, 6)
        self.progress.setTextVisible(False)
        self.progress.hide()
        self.statusBar().setSizeGripEnabled(True)
        self.statusBar().addWidget(self.status_label, 1)
        self.statusBar().addPermanentWidget(self.progress)

        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(8)
        splitter.setOpaqueResize(False)
        root.addWidget(splitter, 1)

        # bal oldal
        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(8)

        source_group = QGroupBox("Aktuális lekérdezés")
        source_form = QFormLayout(source_group)
        source_form.setRowWrapPolicy(QFormLayout.WrapLongRows)
        self.current_file_label = QLabel("-")
        self.current_file_label.setWordWrap(True)
        self.current_type_label = QLabel("-")
        self.current_cache_label = QLabel("-")
        source_form.addRow("Fájl", self.current_file_label)
        source_form.addRow("Lekérdezés típusa", self.current_type_label)
        source_form.addRow("Gyorsítótár", self.current_cache_label)

        filter_group = QGroupBox("Szűrők")
        form = QFormLayout(filter_group)
        form.setRowWrapPolicy(QFormLayout.WrapLongRows)
        form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        self.filter_company = MultiSelectComboBox("Összes")
        self.filter_year = MultiSelectComboBox("Összes")
        self.filter_category = MultiSelectComboBox("Összes")
        self.filter_type = MultiSelectComboBox("Összes")
        for label, value in LABEL_TO_ADATTIPUS.items():
            self.filter_type.add_check_item(label, False, value)
        self.filter_code = QLineEdit()
        self.filter_code.setPlaceholderText("pl. 170401;170402")
        self.filter_company.selectionChanged.connect(self.apply_filters)
        self.filter_year.selectionChanged.connect(self.apply_filters)
        self.filter_category.selectionChanged.connect(self.apply_filters)
        self.filter_type.selectionChanged.connect(self.apply_filters)
        self.filter_code.textChanged.connect(self.apply_filters)
        form.addRow("Cég", self.filter_company)
        form.addRow("Év", self.filter_year)
        form.addRow("Termékcsoport", self.filter_category)
        form.addRow("Adattípus", self.filter_type)
        form.addRow("HAK", self.filter_code)
        reset_btn = QPushButton("Szűrők törlése")
        reset_btn.clicked.connect(self.reset_filters)
        form.addRow("", reset_btn)

        self.exec_text = QTextEdit()
        self.exec_text.setReadOnly(True)
        self.company_info = QTextEdit()
        self.company_info.setReadOnly(True)

        left_layout.addWidget(source_group)
        left_layout.addWidget(filter_group)
        left_layout.addWidget(QLabel("Összefoglaló"))
        left_layout.addWidget(self.exec_text, 1)
        left_layout.addWidget(QLabel("Cégadatok"))
        left_layout.addWidget(self.company_info, 1)
        sidebar = QScrollArea()
        sidebar.setWidgetResizable(True)
        sidebar.setFrameShape(QScrollArea.NoFrame)
        sidebar.setMinimumWidth(250)
        sidebar.setWidget(left)
        splitter.addWidget(sidebar)

        # jobb oldal
        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)

        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        self.tab_dashboard = QWidget()
        self.tab_data = QWidget()
        self.tab_ktj = QWidget()
        self.tab_missing = QWidget()

        self.tabs.addTab(self.tab_dashboard, "Áttekintés")
        self.tabs.addTab(self.tab_data, "Adatok")
        self.tabs.addTab(self.tab_ktj, "KTJ szerint")
        self.tabs.addTab(self.tab_missing, "Kimaradt")
        self.tabs.currentChanged.connect(self.on_tab_changed)

        right_layout.addWidget(self.tabs)
        splitter.addWidget(right)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([330, 1270])

        self._build_dashboard_tab()
        self._build_data_tab()
        self._build_ktj_tab()
        self._build_missing_tab()

    def _apply_styles(self):
        self.setStyleSheet("""
            QWidget {
                background: #f6f8fb;
                color: #0f172a;
                font-size: 13px;
            }
            QGroupBox, #dashboardCard, #chartPanel, #fileDropWidget {
                background: white;
                border: 1px solid #dde5ef;
                border-radius: 12px;
            }
            QGroupBox {
                margin-top: 10px;
                padding-top: 12px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 6px;
                color: #334155;
                font-weight: 700;
            }
            QTabWidget::pane {
                border: 1px solid #dde5ef;
                border-radius: 12px;
                background: white;
                top: -1px;
            }
            QTabBar::tab {
                background: #eef2f7;
                border: 1px solid #dde5ef;
                padding: 7px 14px;
                margin-right: 4px;
                border-top-left-radius: 10px;
                border-top-right-radius: 10px;
                color: #334155;
            }
            QTabBar::tab:selected {
                background: white;
                color: #0f172a;
                font-weight: 700;
            }
            QPushButton {
                background: white;
                border: 1px solid #cbd5e1;
                border-radius: 8px;
                padding: 6px 10px;
            }
            QPushButton:hover { background: #f1f5f9; }
            QPushButton:focus, QLineEdit:focus, QComboBox:focus {
                border: 1px solid #087f74;
            }
            QPushButton:disabled { color: #94a3b8; background: #eef2f7; }
            #primaryButton { background: #087f74; color: white; border: none; padding: 10px 18px; font-weight: 700; }
            #primaryButton:hover { background: #06665d; }
            #primaryButton:disabled { background: #b6d5d0; }
            #appTitle { font-size: 22px; font-weight: 700; }
            #mutedLabel { color: #64748b; background: transparent; }
            QMenuBar { background: white; padding: 4px; border-bottom: 1px solid #e2e8f0; }
            QMenuBar::item { padding: 6px 12px; border-radius: 6px; }
            QMenuBar::item:selected, QMenu::item:selected { background: #e4f2ee; color: #065f56; }
            QMenu { background: white; border: 1px solid #dde5ef; padding: 6px; }
            QMenu::item { padding: 8px 26px; border-radius: 4px; }
            QMenu::separator { height: 1px; background: #e2e8f0; margin: 5px 8px; }
            QStatusBar { background: white; border-top: 1px solid #e2e8f0; padding: 4px; }
            QStatusBar::item { border: none; }
            QProgressBar { border: none; background: #e4f2ee; border-radius: 3px; }
            QProgressBar::chunk { background: #087f74; border-radius: 3px; }
            QSplitter::handle { background: transparent; }
            QSplitter::handle:hover { background: #d7e9e5; }
            QScrollBar:vertical { width: 10px; background: transparent; margin: 0; }
            QScrollBar::handle:vertical { background: #cbd5e1; border-radius: 5px; min-height: 28px; }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }
            QTableView { gridline-color: #edf2f7; selection-background-color: #e4f2ee; selection-color: #0f172a; }
            QHeaderView::section { background: #f1f5f9; border: none; border-bottom: 1px solid #dde5ef; padding: 8px; }
            QLineEdit, QComboBox, QTextEdit, QTableView {
                background: white;
                border: 1px solid #cbd5e1;
                border-radius: 8px;
                padding: 4px 6px;
            }
            #dashboardCardTitle {
                color: #64748b;
                font-size: 12px;
                font-weight: 600;
            }
            #dashboardCardValue {
                color: #0f172a;
                font-size: 20px;
                font-weight: 800;
            }
            #dashboardCardSub {
                color: #94a3b8;
                font-size: 11px;
            }
            #chartPanelTitle {
                font-size: 13px;
                font-weight: 700;
                color: #0f172a;
            }
        """)

    def _build_dashboard_tab(self):
        root = QVBoxLayout(self.tab_dashboard)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(6)

        self.dashboard_scroll = QScrollArea()
        self.dashboard_scroll.setWidgetResizable(True)
        self.dashboard_scroll.setFrameShape(QScrollArea.NoFrame)

        self.dashboard_content = QWidget()
        self.dashboard_layout = QVBoxLayout(self.dashboard_content)
        self.dashboard_layout.setContentsMargins(4, 4, 4, 4)
        self.dashboard_layout.setSpacing(6)

        # KPI sor
        self.kpi_grid = QGridLayout()
        self.kpi_grid.setSpacing(10)
        self.kpi_total = DashboardCard("Összes mennyiség", "Valós összesített mennyiség")
        self.kpi_companies = DashboardCard("Cégek száma", "Szűrt eredmény")
        self.kpi_codes = DashboardCard("Hulladékkódok", "Egyedi HAK")
        self.kpi_danger = DashboardCard("Veszélyességi arány", "Összmennyiség alapján")
        self.kpi_years = DashboardCard("Vizsgált évek", "Aktuális szűrés")
        self.kpi_cards = [self.kpi_total, self.kpi_companies, self.kpi_codes, self.kpi_danger, self.kpi_years]
        self.dashboard_layout.addLayout(self.kpi_grid)

        self.chart_grid = QGridLayout()
        self.chart_grid.setSpacing(10)
        self.panel_yearly = ChartPanel("Éves trend", lambda: self.open_chart("Éves trend", self.render_yearly_chart))
        self.panel_finance = ChartPanel("Árbevétel / saját tőke", lambda: self.open_chart("Árbevétel / saját tőke", self.render_finance_chart))

        self.panel_company = ChartPanel("Top cégek összetétele", lambda: self.open_chart("Top cégek összetétele", self.render_company_chart))
        self.panel_category = ChartPanel("Termékcsoportok megoszlása", lambda: self.open_chart("Termékcsoportok megoszlása", self.render_category_chart))

        self.panel_danger = ChartPanel("Veszélyességi megoszlás", lambda: self.open_chart("Veszélyességi megoszlás", self.render_danger_chart))
        self.panel_hak_breakdown = ChartPanel("Kategória bontás (HAK)", lambda: self.open_chart("Kategória bontás (HAK)", self.render_hak_breakdown_chart))
        self.chart_panels = [
            self.panel_yearly, self.panel_finance, self.panel_company,
            self.panel_category, self.panel_danger, self.panel_hak_breakdown,
        ]
        for panel in self.chart_panels:
            panel.setMinimumHeight(260)
        self.dashboard_layout.addLayout(self.chart_grid)
        self._dashboard_columns = None
        self._reflow_dashboard(1200)

        self.dashboard_scroll.setWidget(self.dashboard_content)
        self.dashboard_scroll.viewport().installEventFilter(self)
        root.addWidget(self.dashboard_scroll, 1)

        # drill-down összekötések
        self.panel_yearly.canvas.clickedTarget.connect(self.on_dashboard_drilldown)
        self.panel_category.canvas.clickedTarget.connect(self.on_dashboard_drilldown)
        self.panel_company.canvas.clickedTarget.connect(self.on_dashboard_drilldown)
        self.panel_finance.canvas.clickedTarget.connect(self.on_dashboard_drilldown)
        self.panel_hak_breakdown.canvas.clickedTarget.connect(self.on_dashboard_drilldown)

    def eventFilter(self, watched, event):
        if watched is self.dashboard_scroll.viewport() and event.type() == QEvent.Resize:
            self._reflow_dashboard(event.size().width())
        return super().eventFilter(watched, event)

    def _reflow_dashboard(self, width):
        columns = (5 if width >= 1000 else 3 if width >= 650 else 2, 2 if width >= 900 else 1)
        if columns == self._dashboard_columns:
            return
        self._dashboard_columns = columns
        for grid, widgets, count in (
            (self.kpi_grid, self.kpi_cards, columns[0]),
            (self.chart_grid, self.chart_panels, columns[1]),
        ):
            while grid.count():
                grid.takeAt(0)
            for index in range(grid.columnCount()):
                grid.setColumnStretch(index, 0)
            for index, widget in enumerate(widgets):
                grid.addWidget(widget, index // count, index % count)
            for index in range(count):
                grid.setColumnStretch(index, 1)

    def _build_data_tab(self):
        layout = QVBoxLayout(self.tab_data)
        self.data_year_tabs = QTabWidget()
        layout.addWidget(self.data_year_tabs)

    def _build_ktj_tab(self):
        layout = QVBoxLayout(self.tab_ktj)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        self.ktj_table = QTableView()
        self.configure_table(self.ktj_table)
        layout.addWidget(self.ktj_table)

    def _build_missing_tab(self):
        layout = QVBoxLayout(self.tab_missing)
        self.missing_table = QTableView()
        self.configure_table(self.missing_table)
        layout.addWidget(self.missing_table)

    def configure_table(self, table: QTableView):
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QTableView.SelectRows)
        table.setEditTriggers(QTableView.NoEditTriggers)
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        table.horizontalHeader().setResizeContentsPrecision(100)
        table.horizontalHeader().setStretchLastSection(True)

    # -------------------------------------------------------------------------
    # Drill-down
    # -------------------------------------------------------------------------
    def on_dashboard_drilldown(self, payload):
        if not isinstance(payload, dict):
            return
        kind = payload.get("kind")
        value = payload.get("value")

        if kind == "year" and value is not None:
            try:
                self.filter_year.set_checked_by_user_data({int(value)})
                self.tabs.setCurrentIndex(1)
                self._invalidate_filtered_cache()
                self.refresh_all_views()
            except Exception:
                pass
            return

        if kind == "category" and value:
            self._drill_category = str(value)
            self.render_hak_breakdown_chart(self.panel_hak_breakdown.canvas)
            self.set_status(f"Drill-down kategória: {self._drill_category}")
            return

        if kind == "company" and value:
            try:
                self.filter_company.set_checked_by_user_data({str(value)})
                self._invalidate_filtered_cache()
                self.refresh_all_views()
            except Exception:
                pass
            return

        if kind == "hak" and value:
            code = str(value).strip()
            current = self.filter_code.text().strip()
            parts = [p.strip() for p in current.split(";") if p.strip()]
            if code and code not in parts:
                parts.append(code)
            self.filter_code.setText(";".join(parts))
            self._invalidate_filtered_cache()
            self.refresh_all_views()
            return

    # -------------------------------------------------------------------------
    # Import / Fetch
    # -------------------------------------------------------------------------
    def open_import_dialog(self):
        if self._is_busy():
            QMessageBox.warning(self, "Folyamatban", "Várd meg a folyamatban lévő lekérést.")
            return
        dlg = ImportDatabaseDialog(self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        payload = dlg.payload()
        self.current_file_path = payload["file_path"]
        self.current_query_label = payload["query_type"]
        self.current_use_cache = payload["use_cache"]

        self.current_file_label.setText(self.current_file_path)
        self.current_type_label.setText(self.current_query_label)
        self.current_cache_label.setText("Igen" if self.current_use_cache else "Nem")

        self.start_fetch()

    def start_fetch(self):
        if self._is_busy():
            QMessageBox.warning(self, "Folyamatban", "Várd meg a folyamatban lévő műveletet.")
            return
        file_path = self.current_file_path.strip()
        if not file_path:
            QMessageBox.warning(self, "Hiányzó fájl", "Előbb válassz ki egy Excel fájlt.")
            return

        self._set_busy(True)
        self.set_status("Forrásimport és hulladékadatok lekérése a háttérben...")
        self.worker = FetchWorker(
            self.source_df.copy(),
            self.current_use_cache,
            False,
            self.selected_types_from_label(self.current_query_label),
            max_workers=4,
            file_path=file_path,
        )
        self.worker.progress_signal.connect(self.on_progress)
        self.worker.finished_signal.connect(self.on_fetch_finished)
        self.worker.error_signal.connect(self.on_fetch_error)
        self.worker.finished.connect(self._fetch_thread_finished)
        self.worker.start()

    def _fetch_thread_finished(self):
        self._set_busy(False)
        if self.tabs.currentIndex() == 2:
            self.on_tab_changed(2)

    def on_source_progress(self, step, total, text):
        self.progress.setMaximum(max(total, 1))
        self.progress.setValue(step)
        self.status_label.setText(text)

    def on_progress(self, step, total, text):
        self.progress.setMaximum(max(total, 1))
        self.progress.setValue(step)
        self.status_label.setText(text)

    def on_fetch_finished(self, source_df, raw_df, missing_df, max_year, aggregates=None):
        self._ktj_confirmed = False
        self._drill_category = None
        self.source_df = source_df
        self.raw_df = raw_df
        self.missing_df = missing_df
        self.max_data_year = max_year

        self.set_status("Adatok aggregálása...", 90, 100)
        self.agg = aggregates if aggregates is not None else aggregate_data(
            self.raw_df, self.source_df, self.category_service,
        )

        self.set_status("Szűrők frissítése...", 95, 100)
        self.populate_filters()
        self.filter_type.set_all_checked(False)
        self.apply_default_year_selection()

        self.set_status("Nézetek frissítése...", 98, 100)
        self._invalidate_filtered_cache()
        self.refresh_all_views()

        self.set_status("Lekérdezés kész.", 100, 100)
        QMessageBox.information(self, "Kész", "Az adatok lekérése befejeződött.")

    def on_fetch_error(self, err):
        self.set_status(f"Hiba: {err}", 0, 100)
        QMessageBox.critical(self, "Hiba", err)

    # -------------------------------------------------------------------------
    # Szűrők
    # -------------------------------------------------------------------------
    def populate_filters(self):
        master = self.agg.master

        for combo in [self.filter_company, self.filter_year, self.filter_category]:
            combo.blockSignals(True)
            combo.clear_items()
            combo.blockSignals(False)

        if master is None or master.empty:
            return

        for combo in [self.filter_company, self.filter_year, self.filter_category]:
            combo.blockSignals(True)

        try:
            if "CEGNEV" in master.columns:
                self.filter_company.add_check_items(
                    (c, c) for c in sorted(master["CEGNEV"].dropna().astype(str).unique())
                )

            if "EV" in master.columns:
                self.filter_year.add_check_items(
                    (str(y), int(y)) for y in sorted(master["EV"].dropna().astype(int).unique(), reverse=True)
                )

            if "PRODUCT_GROUP" in master.columns:
                self.filter_category.add_check_items(
                    (g, g) for g in sorted(master["PRODUCT_GROUP"].dropna().astype(str).unique())
                )
        finally:
            for combo in [self.filter_company, self.filter_year, self.filter_category]:
                combo.blockSignals(False)

    def apply_default_year_selection(self):
        master = self.agg.master
        if master is None or master.empty or "EV" not in master.columns:
            return
        years = sorted(master["EV"].dropna().astype(int).unique(), reverse=True)
        if years:
            self.filter_year.set_checked_by_user_data(set(years[:2]))

    def apply_filters(self):
        self._filter_timer.start(FILTER_DEBOUNCE_MS)

    def reset_filters(self):
        self._ktj_confirmed = False
        for combo in [self.filter_company, self.filter_year, self.filter_category, self.filter_type]:
            combo.blockSignals(True)
            combo.set_all_checked(False)
            combo.blockSignals(False)

        self.filter_code.blockSignals(True)
        self.filter_code.clear()
        self.filter_code.blockSignals(False)

        self._drill_category = None
        self.set_status("Szűrők alaphelyzetbe állítva.")
        self._invalidate_filtered_cache()
        self.refresh_all_views()

    def current_filters(self):
        return Filters(
            companies=tuple(self.filter_company.checked_values()),
            years=tuple(self.filter_year.checked_values()),
            categories=tuple(self.filter_category.checked_values()),
            code=self.filter_code.text().strip(),
            types=tuple(self.filter_type.checked_values()),
        )

    def get_filtered_master(self):
        return filter_frame(self.agg.master, self.current_filters())

    def get_filtered_raw(self):
        return filter_frame(self.raw_df, self.current_filters(), self.category_service)

    def get_filtered_base_unique(self):
        return unique_quantities(self.get_filtered_master())

    # -------------------------------------------------------------------------
    # Refresh
    # -------------------------------------------------------------------------
    def refresh_all_views(self):
        self._invalidate_filtered_cache()
        self._dirty_views = {0, 1, 2, 3}
        self.refresh_exec_summary()
        self.refresh_company_info()
        self.on_tab_changed(self.tabs.currentIndex())

    def on_tab_changed(self, idx):
        if idx not in self._dirty_views:
            return
        if idx == 2:
            if self._is_busy():
                return
            if not self.raw_df.empty:
                self._maybe_load_ktj()
                if not self._ktj_confirmed:
                    return
            else:
                self._set_table_df(self.ktj_table, pd.DataFrame())
        else:
            {0: self.refresh_dashboard, 1: self.refresh_data_year_tabs, 3: self.refresh_missing_table}[idx]()
        self._dirty_views.discard(idx)

    def _maybe_load_ktj(self):
        if self._is_busy():
            return
        if self._ktj_confirmed:
            self.refresh_ktj_tab()
            return
        reply = QMessageBox.question(
            self,
            "KTJ adatok betöltése",
            "A KTJ nézet megjelenítéséhez az OKIR rendszerből kell metaadatokat lekérdezni.\n\n"
            "Ez az adatok mennyiségétől függően több percig is tarthat.\n\n"
            "Szeretnéd most betölteni az adatokat?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            self._ktj_confirmed = True
            self.refresh_ktj_tab()

    def refresh_dashboard(self):
        self.refresh_kpis()
        self.render_yearly_chart(self.panel_yearly.canvas)
        self.render_finance_chart(self.panel_finance.canvas)
        self.render_company_chart(self.panel_company.canvas)
        self.render_category_chart(self.panel_category.canvas)
        self.render_danger_chart(self.panel_danger.canvas)
        self.render_hak_breakdown_chart(self.panel_hak_breakdown.canvas)

    def refresh_kpis(self):
        snap = self.get_filtered_snapshot()
        master = snap.get("master")
        base = snap.get("base")

        if master is None or master.empty or base is None or base.empty or "OSSZES_MENNYISEG" not in base.columns:
            for kpi in [self.kpi_total, self.kpi_companies, self.kpi_codes, self.kpi_danger, self.kpi_years]:
                kpi.set_value("-")
            return

        total_qty = pd.to_numeric(base["OSSZES_MENNYISEG"], errors="coerce").fillna(0).sum()
        company_count = base["CEGNEV"].nunique() if "CEGNEV" in base.columns else 0
        waste_count = base["HULLADEKKOD"].nunique() if "HULLADEKKOD" in base.columns else 0

        if "HULLADEK_VESZELYES_10" in base.columns:
            dangerous_qty = base.loc[
                pd.to_numeric(base["HULLADEK_VESZELYES_10"], errors="coerce").fillna(0) == 1,
                "OSSZES_MENNYISEG",
            ]
            dangerous_qty = pd.to_numeric(dangerous_qty, errors="coerce").fillna(0).sum()
        else:
            dangerous_qty = 0

        dangerous_ratio = (dangerous_qty / total_qty * 100) if total_qty else 0
        years = sorted(base["EV"].dropna().astype(int).unique()) if "EV" in base.columns else []
        year_text = f"{years[0]} – {years[-1]}" if len(years) > 1 else (str(years[0]) if years else "-")

        self.kpi_total.set_value(format_quantity_auto_unit(total_qty, 2))
        self.kpi_companies.set_value(str(company_count))
        self.kpi_codes.set_value(str(waste_count))
        self.kpi_danger.set_value(f"{format_hu_number(dangerous_ratio, 2)} %")
        self.kpi_years.set_value(year_text)

    # -------------------------------------------------------------------------
    # Chart helpers
    # -------------------------------------------------------------------------
    def style_axes(self, ax):
        ax.clear()
        ax.grid(axis="y", linestyle="--", alpha=0.15, color="#cbd5e1")
        ax.tick_params(axis="both", labelsize=9, color="#94a3b8")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color("#e2e8f0")
        ax.spines["bottom"].set_color("#e2e8f0")

    def draw_empty(self, canvas, title, text="Nincs megjeleníthető adat"):
        ax = canvas.ax
        ax.clear()
        canvas.reset_hover()
        ax.set_facecolor("#fafafa")
        ax.text(
            0.5, 0.5, text,
            ha="center", va="center",
            transform=ax.transAxes,
            fontsize=10, color="#94a3b8",
            style="italic",
        )
        ax.set_xticks([])
        ax.set_yticks([])
        for spine in ax.spines.values():
            spine.set_visible(False)
        canvas.figure.subplots_adjust(left=0.08, right=0.98, top=0.94, bottom=0.08)
        canvas.draw_idle()

    def open_chart(self, title, renderer):
        FullscreenChartDialog(title, renderer, self).exec()

    # -------------------------------------------------------------------------
    # Grafikonok
    # -------------------------------------------------------------------------
    def render_yearly_chart(self, canvas):
        snap = self.get_filtered_snapshot()
        yearly = snap.get("yearly")

        if yearly is None or yearly.empty:
            self.draw_empty(canvas, "Éves trend")
            return

        ax = canvas.ax
        self.style_axes(ax)
        canvas.reset_hover()

        xs = yearly["EV"].tolist()
        ys_raw = yearly["OSSZES_MENNYISEG"].tolist()
        ys, unit = convert_for_axis(ys_raw)

        ax.plot(xs, ys, marker="o", linewidth=2.4, color="#2563eb", markersize=7, markerfacecolor="white", markeredgewidth=2)
        ax.fill_between(xs, ys, alpha=0.08, color="#2563eb")
        ax.set_xlabel("Év", fontsize=9, color="#64748b")
        ax.set_ylabel(f"Mennyiség ({unit})", fontsize=9, color="#64748b")
        ax.xaxis.set_major_formatter(axis_int_formatter())
        ax.yaxis.set_major_formatter(axis_number_formatter(0))
        ax.grid(axis="x", linestyle="--", alpha=0.10, color="#cbd5e1")

        canvas.add_point_targets(
            xs, ys,
            [f"Év: {int(x)}\nMennyiség: {format_quantity_auto_unit(y, 2)}" for x, y in zip(xs, ys_raw)],
            payloads=[{"kind": "year", "value": int(x)} for x in xs],
        )
        canvas.figure.subplots_adjust(left=0.12, right=0.98, top=0.94, bottom=0.18)
        canvas.draw_idle()

    def render_finance_chart(self, canvas):
        if self.source_df.empty or "CEGNEV" not in self.source_df.columns:
            self.draw_empty(canvas, "Árbevétel / saját tőke")
            return

        src = self.source_df.copy()
        if "SRC_NETTO_ARBEV" not in src.columns:
            src["SRC_NETTO_ARBEV"] = ""
        if "SRC_SAJAT_TOKE" not in src.columns:
            src["SRC_SAJAT_TOKE"] = ""

        src["ARBEV_NUM"] = pd.to_numeric(parse_hu_numeric(src["SRC_NETTO_ARBEV"]), errors="coerce")
        src["SAJAT_TOKE_NUM"] = pd.to_numeric(parse_hu_numeric(src["SRC_SAJAT_TOKE"]), errors="coerce")

        plot_df = (
            src[["CEGNEV", "ARBEV_NUM", "SAJAT_TOKE_NUM"]]
            .drop_duplicates("CEGNEV")
            .dropna(subset=["ARBEV_NUM", "SAJAT_TOKE_NUM"], how="all")
            .copy()
        )

        if plot_df.empty:
            self.draw_empty(canvas, "Árbevétel / saját tőke")
            return

        plot_df = plot_df.fillna({"ARBEV_NUM": 0, "SAJAT_TOKE_NUM": 0})
        plot_df = plot_df.sort_values("ARBEV_NUM", ascending=False).head(8).reset_index(drop=True)

        ax = canvas.ax
        self.style_axes(ax)
        canvas.reset_hover()

        x = range(len(plot_df))
        width = 0.38
        arbev_m = (plot_df["ARBEV_NUM"] / 1_000_000).tolist()
        toke_m = (plot_df["SAJAT_TOKE_NUM"] / 1_000_000).tolist()
        labels = [shorten_label(v, 11) for v in plot_df["CEGNEV"].astype(str).tolist()]
        full_labels = plot_df["CEGNEV"].astype(str).tolist()

        bars1 = ax.bar([i - width / 2 for i in x], arbev_m, width=width, label="Árbevétel", color="#2563eb", alpha=0.90)
        bars2 = ax.bar([i + width / 2 for i in x], toke_m, width=width, label="Saját tőke", color="#8b5cf6", alpha=0.90)

        ax.set_xticks(list(x))
        ax.set_xticklabels(labels, rotation=22, ha="right", fontsize=8)
        ax.set_ylabel("Millió Ft", fontsize=9, color="#64748b")
        ax.yaxis.set_major_formatter(axis_number_formatter(0))
        ax.legend(frameon=False, fontsize=8, loc="upper right")

        canvas.add_bar_targets(
            bars1,
            [f"{c}\nNettó árbevétel: {format_hu_number(a, 0)} Ft" for c, a in zip(full_labels, plot_df["ARBEV_NUM"].tolist())],
            plot_df["ARBEV_NUM"].tolist(),
            horizontal=False,
            payloads=[{"kind": "company", "value": c} for c in full_labels],
            show_quantity_line=False,
        )
        canvas.add_bar_targets(
            bars2,
            [f"{c}\nSaját tőke: {format_hu_number(t, 0)} Ft" for c, t in zip(full_labels, plot_df["SAJAT_TOKE_NUM"].tolist())],
            plot_df["SAJAT_TOKE_NUM"].tolist(),
            horizontal=False,
            payloads=[{"kind": "company", "value": c} for c in full_labels],
            show_quantity_line=False,
        )
        canvas.figure.subplots_adjust(left=0.12, right=0.98, top=0.94, bottom=0.30)
        canvas.draw_idle()

    def render_company_chart(self, canvas):
        master = self.get_filtered_master()
        base = self.get_filtered_base_unique()

        if (
            master is None or master.empty or base.empty
            or "CEGNEV" not in base.columns or "OSSZES_MENNYISEG" not in base.columns
            or "PRODUCT_GROUP" not in master.columns
        ):
            self.draw_empty(canvas, "Top cégek összetétele")
            return

        base_tmp = base.copy()
        base_tmp["OSSZES_MENNYISEG"] = pd.to_numeric(base_tmp["OSSZES_MENNYISEG"], errors="coerce").fillna(0)
        top_companies = (
            base_tmp.groupby("CEGNEV", as_index=False)["OSSZES_MENNYISEG"]
            .sum()
            .sort_values("OSSZES_MENNYISEG", ascending=False)
            .head(6)["CEGNEV"].tolist()
        )

        tmp = master.copy()
        tmp["OSSZES_MENNYISEG"] = pd.to_numeric(tmp["OSSZES_MENNYISEG"], errors="coerce").fillna(0)
        sub = (
            tmp[tmp["CEGNEV"].isin(top_companies)]
            .groupby(["CEGNEV", "PRODUCT_GROUP"], as_index=False)["OSSZES_MENNYISEG"]
            .sum()
        )
        pivot = sub.pivot(index="CEGNEV", columns="PRODUCT_GROUP", values="OSSZES_MENNYISEG").fillna(0)

        if pivot.empty:
            self.draw_empty(canvas, "Top cégek összetétele")
            return

        pivot = pivot.loc[pivot.sum(axis=1).sort_values(ascending=True).index]

        ax = canvas.ax
        ax.clear()
        canvas.reset_hover()

        full_labels = pivot.index.tolist()
        short_labels = [shorten_label(v, 20) for v in full_labels]
        lefts = [0.0] * len(pivot.index)

        for col in pivot.columns:
            vals_raw = pivot[col].tolist()
            vals_t = [v / 1000 for v in vals_raw]
            color = self.category_colors.get(str(col), "#94a3b8")
            bars = ax.barh(short_labels, vals_t, left=lefts, label=str(col), color=color, alpha=0.90)
            for i, bar in enumerate(bars):
                bar._x0 = lefts[i]
            canvas.add_bar_targets(
                bars,
                [f"Cég: {c}\nKategória: {col}" for c in full_labels],
                vals_raw,
                horizontal=True,
                payloads=[{"kind": "company", "value": c} for c in full_labels],
            )
            lefts = [l + v for l, v in zip(lefts, vals_t)]

        ax.set_xlabel("Mennyiség (t)", fontsize=9, color="#64748b")
        ax.tick_params(axis="y", labelsize=8)
        ax.xaxis.set_major_formatter(axis_number_formatter(0))
        ax.legend(frameon=False, fontsize=7, loc="lower right", ncols=2)
        ax.grid(axis="x", linestyle="--", alpha=0.15, color="#cbd5e1")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        canvas.figure.subplots_adjust(left=0.28, right=0.97, top=0.94, bottom=0.14)
        canvas.draw_idle()

    def render_category_chart(self, canvas):
        snap = self.get_filtered_snapshot()
        cat_totals = snap.get("category_totals")

        if cat_totals is None or cat_totals.empty:
            self.draw_empty(canvas, "Termékcsoportok megoszlása")
            return

        cat = cat_totals.head(8).sort_values("OSSZES_MENNYISEG", ascending=True).copy()
        if cat.empty:
            self.draw_empty(canvas, "Termékcsoportok megoszlása")
            return

        total = float(cat["OSSZES_MENNYISEG"].sum()) if "OSSZES_MENNYISEG" in cat.columns else 0.0
        cat["PCT"] = (cat["OSSZES_MENNYISEG"] / total * 100) if total else 0.0

        ax = canvas.ax
        self.style_axes(ax)
        canvas.reset_hover()

        full_labels = cat["PRODUCT_GROUP"].astype(str).tolist()
        short_labels = [shorten_label(v, 20) for v in full_labels]
        vals_raw = cat["OSSZES_MENNYISEG"].tolist()
        vals, unit = convert_for_axis(vals_raw)
        percents = cat["PCT"].tolist()

        colors = [self.category_colors.get(l, "#94a3b8") for l in full_labels]
        bars = ax.barh(short_labels, vals, color=colors, alpha=0.90)
        ax.set_xlabel(f"Mennyiség ({unit})", fontsize=9, color="#64748b")
        ax.tick_params(axis="y", labelsize=8)
        ax.xaxis.set_major_formatter(axis_number_formatter(0))
        ax.grid(axis="x", linestyle="--", alpha=0.15, color="#cbd5e1")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)

        for bar, pct in zip(bars, percents):
            ax.text(
                bar.get_width(), bar.get_y() + bar.get_height() / 2,
                f"  {format_hu_number(pct, 1)}%",
                va="center", ha="left", fontsize=8, color="#475569",
            )

        canvas.add_bar_targets(
            bars,
            [f"Kategória: {l}\nKattints a HAK bontáshoz!" for l in full_labels],
            vals_raw,
            horizontal=True,
            percents=percents,
            payloads=[{"kind": "category", "value": l} for l in full_labels],
        )
        canvas.figure.subplots_adjust(left=0.30, right=0.97, top=0.94, bottom=0.12)
        canvas.draw_idle()

    def render_danger_chart(self, canvas):
        base = self.get_filtered_base_unique()
        if base.empty or "OSSZES_MENNYISEG" not in base.columns or "HULLADEK_VESZELYES_10" not in base.columns:
            self.draw_empty(canvas, "Veszélyességi megoszlás")
            return

        ax = canvas.ax
        ax.clear()
        canvas.reset_hover()

        qty = pd.to_numeric(base["OSSZES_MENNYISEG"], errors="coerce").fillna(0)
        danger = pd.to_numeric(base["HULLADEK_VESZELYES_10"], errors="coerce").fillna(0)
        dangerous_qty = qty[danger == 1].sum()
        non_dangerous_qty = qty[danger != 1].sum()

        vals = [dangerous_qty, non_dangerous_qty]
        labels = ["Veszélyes", "Nem veszélyes"]
        total = sum(vals)
        percents = [(v / total * 100) if total else 0 for v in vals]

        wedges, _ = ax.pie(
            vals,
            labels=labels,
            startangle=90,
            colors=["#ef4444", "#94a3b8"],
            wedgeprops=dict(width=0.50, edgecolor="white", linewidth=2),
        )
        canvas.add_pie_targets(
            wedges, labels, vals, percents=percents,
            payloads=[{"kind": "danger", "value": 1}, {"kind": "danger", "value": 0}],
        )
        canvas.figure.subplots_adjust(left=0.06, right=0.94, top=0.94, bottom=0.06)
        canvas.draw_idle()

    def render_hak_breakdown_chart(self, canvas):
        cat = str(self._drill_category or "").strip()
        if not cat:
            self.panel_hak_breakdown.set_title("Kategória bontás (HAK)")
            self.draw_empty(
                canvas, "Kategória bontás (HAK)",
                text="Kattints egy kategóriára\na 'Termékcsoportok megoszlása' grafikonon.",
            )
            return

        self.panel_hak_breakdown.set_title(f"HAK bontás — {cat}")

        master = self.get_filtered_master()
        if master is None or master.empty:
            self.draw_empty(canvas, "HAK bontás")
            return

        needed = {"PRODUCT_GROUP", "HULLADEKKOD", "OSSZES_MENNYISEG"}
        if not needed.issubset(set(master.columns)):
            self.draw_empty(canvas, "HAK bontás")
            return

        df = master[master["PRODUCT_GROUP"].astype(str) == cat].copy()
        if df.empty:
            self.draw_empty(canvas, f"{cat} bontás (HAK)", text="Nincs adat ehhez a kategóriához.")
            return

        if "HULLADEK_MEGNEVEZES" not in df.columns:
            df["HULLADEK_MEGNEVEZES"] = ""

        df["OSSZES_MENNYISEG"] = pd.to_numeric(df["OSSZES_MENNYISEG"], errors="coerce").fillna(0)
        df["HAK_LABEL"] = (
            df["HULLADEKKOD"].astype(str).str.strip()
            + "  "
            + df["HULLADEK_MEGNEVEZES"].fillna("").astype(str).str.strip()
        )

        grouped = (
            df.groupby(["HULLADEKKOD", "HAK_LABEL"], as_index=False)["OSSZES_MENNYISEG"]
            .sum()
            .sort_values("OSSZES_MENNYISEG", ascending=False)
            .head(12)
            .sort_values("OSSZES_MENNYISEG", ascending=True)
        )

        if grouped.empty:
            self.draw_empty(canvas, f"{cat} bontás (HAK)")
            return

        full_labels = grouped["HAK_LABEL"].astype(str).tolist()
        short_labels = [shorten_label(v, 42) for v in full_labels]
        vals_raw = grouped["OSSZES_MENNYISEG"].tolist()
        vals, unit = convert_for_axis(vals_raw)
        total = sum(vals_raw) or 1
        percents = [v / total * 100 for v in vals_raw]

        ax = canvas.ax
        self.style_axes(ax)
        canvas.reset_hover()

        color = self.category_colors.get(cat, "#0ea5e9")
        bars = ax.barh(short_labels, vals, color=color, alpha=0.88)
        ax.set_xlabel(f"Mennyiség ({unit})", fontsize=9, color="#64748b")
        ax.set_ylabel("HAK kód / megnevezés", fontsize=9, color="#64748b")
        ax.xaxis.set_major_formatter(axis_number_formatter(0))
        ax.tick_params(axis="y", labelsize=8)
        ax.grid(axis="x", linestyle="--", alpha=0.15, color="#cbd5e1")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)

        for bar, pct in zip(bars, percents):
            ax.text(
                bar.get_width(), bar.get_y() + bar.get_height() / 2,
                f"  {format_hu_number(pct, 1)}%",
                va="center", ha="left", fontsize=8, color="#475569",
            )

        canvas.add_bar_targets(
            bars,
            [f"Kategória: {cat}\n{lbl}" for lbl in full_labels],
            vals_raw,
            horizontal=True,
            percents=percents,
            payloads=[{"kind": "hak", "value": hak} for hak in grouped["HULLADEKKOD"].astype(str).tolist()],
            show_quantity_line=True,
        )
        canvas.figure.subplots_adjust(left=0.42, right=0.97, top=0.94, bottom=0.12)
        canvas.draw_idle()

    # -------------------------------------------------------------------------
    # Tábla nézetek
    # -------------------------------------------------------------------------
    def refresh_data_year_tabs(self):
        while self.data_year_tabs.count():
            page = self.data_year_tabs.widget(0)
            self.data_year_tabs.removeTab(0)
            page.deleteLater()

        master = self.get_filtered_master()
        if master is None or master.empty:
            page = QWidget()
            QVBoxLayout(page).addWidget(QLabel("Nincs megjeleníthető adat."))
            self.data_year_tabs.addTab(page, "Nincs adat")
            return

        if "EV" not in master.columns:
            page = QWidget()
            QVBoxLayout(page).addWidget(QLabel("Az adatok nem tartalmaznak év mezőt."))
            self.data_year_tabs.addTab(page, "Nincs év")
            return

        years = sorted(master["EV"].dropna().astype(int).unique(), reverse=True)
        for year in [None] + years:
            sub = master if year is None else master[master["EV"].astype("Int64") == year]
            tab_name = "Összes" if year is None else str(year)

            out = sub.copy()
            if "HULLADEK_VESZELYES_10" in out.columns:
                out["Veszélyes"] = out["HULLADEK_VESZELYES_10"].apply(
                    lambda x: "Igen" if pd.to_numeric(x, errors="coerce") == 1 else "Nem"
                )
            else:
                out["Veszélyes"] = ""

            out = add_quantity_columns(out)

            preferred = [
                "CEGNEV", "EV", "ADATTIPUS_LABEL", "PRODUCT_GROUP",
                "HULLADEKKOD", "HULLADEK_MEGNEVEZES", "Veszélyes",
                "Mennyiség (kg)", "Mennyiség (t)",
            ]
            existing = [c for c in preferred if c in out.columns]
            if "Veszélyes" not in existing:
                existing.append("Veszélyes")
            out = out[existing].copy()
            out.columns = [
                "Cég" if c == "CEGNEV" else
                "Év" if c == "EV" else
                "Adattípus" if c == "ADATTIPUS_LABEL" else
                "Termékcsoport" if c == "PRODUCT_GROUP" else
                "Hulladékkód" if c == "HULLADEKKOD" else
                "Megnevezés" if c == "HULLADEK_MEGNEVEZES" else
                c
                for c in out.columns
            ]

            page = QWidget()
            page_l = QVBoxLayout(page)
            table = QTableView()
            self.configure_table(table)
            self._set_table_df(table, out)
            page_l.addWidget(table)
            self.data_year_tabs.addTab(page, tab_name)

    def refresh_ktj_tab(self):
        if self._is_busy():
            return
        df = self.get_filtered_raw()
        if df.empty:
            self._set_table_df(self.ktj_table, pd.DataFrame())
            return

        key = self._filters_key()

        def operation():
            try:
                return self._prepare_ktj_view(df.copy())
            finally:
                self.persistence.close()

        def completed(out):
            if key != self._filters_key():
                self._dirty_views.add(2)
                self.on_tab_changed(self.tabs.currentIndex())
                return
            self._set_table_df(self.ktj_table, out)
            self._dirty_views.discard(2)
            self.set_status("KTJ nézet frissítve.")

        self._run_task("KTJ metaadatok betöltése a háttérben...", operation, completed)

    def _prepare_ktj_view(self, df: pd.DataFrame) -> pd.DataFrame:
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
            out["PRODUCT_GROUP"] = out["HULLADEKKOD"].astype(str).apply(self.category_service.resolve_group)

        if "OSSZES_MENNYISEG" in out.columns:
            out["Összes mennyiség"] = out["OSSZES_MENNYISEG"].apply(lambda x: format_quantity_auto_unit(x, 2))
        elif "MENNYISEG" in out.columns:
            out["Összes mennyiség"] = out["MENNYISEG"].apply(lambda x: format_quantity_auto_unit(x, 2))
        else:
            out["Összes mennyiség"] = ""

        self.ktj_service.ensure_metadata_loaded(out)
        out = self.ktj_service.enrich_with_metadata(out)

        preferred_cols = [
            "CEGNEV", "KUJ", "KTJ", "KTJ megnevezés", "Besorolás",
            "Irányítószám", "Helység", "Cím", "Megye", "Régió",
            "EV", "ADATTIPUS_LABEL", "PRODUCT_GROUP", "HULLADEKKOD",
            "HULLADEK_MEGNEVEZES", "Veszélyes", "Összes mennyiség",
        ]
        out = out[[c for c in preferred_cols if c in out.columns]].copy()
        out.rename(columns={
            "CEGNEV": "Cég", "KUJ": "KÜJ", "EV": "Év",
            "ADATTIPUS_LABEL": "Adattípus", "PRODUCT_GROUP": "Termékcsoport",
            "HULLADEKKOD": "Hulladékkód", "HULLADEK_MEGNEVEZES": "Megnevezés",
        }, inplace=True)

        sort_cols = [c for c in ["KÜJ", "KTJ", "Cég", "Év", "Termékcsoport", "Hulladékkód"] if c in out.columns]
        if sort_cols:
            out = out.sort_values(sort_cols, ascending=True, na_position="last")

        return out.reset_index(drop=True)

    def refresh_missing_table(self):
        self._set_table_df(self.missing_table, self.missing_df)

    def _set_table_df(self, table: QTableView, df: pd.DataFrame):
        table.setSortingEnabled(False)
        previous = table.model()
        table.setModel(DataFrameModel(df if df is not None else pd.DataFrame(), table))
        if previous is not None:
            previous.deleteLater()
        table.resizeColumnsToContents()
        table.setSortingEnabled(True)

    # -------------------------------------------------------------------------
    # Összefoglaló / cégadatok
    # -------------------------------------------------------------------------
    def refresh_exec_summary(self):
        master = self.get_filtered_master()
        base = self.get_filtered_base_unique()

        if master is None or master.empty or base.empty or "OSSZES_MENNYISEG" not in base.columns:
            self.exec_text.setPlainText("Nincs megjeleníthető adat.")
            return

        qty = pd.to_numeric(base["OSSZES_MENNYISEG"], errors="coerce").fillna(0)
        total_qty = qty.sum()
        company_count = base["CEGNEV"].nunique() if "CEGNEV" in base.columns else 0
        waste_count = base["HULLADEKKOD"].nunique() if "HULLADEKKOD" in base.columns else 0
        category_count = master["PRODUCT_GROUP"].nunique() if "PRODUCT_GROUP" in master.columns else 0
        years = sorted(base["EV"].dropna().astype(int).unique()) if "EV" in base.columns else []

        if "HULLADEK_VESZELYES_10" in base.columns:
            dangerous_qty = qty[pd.to_numeric(base["HULLADEK_VESZELYES_10"], errors="coerce").fillna(0) == 1].sum()
        else:
            dangerous_qty = 0

        dangerous_ratio = (dangerous_qty / total_qty * 100) if total_qty else 0

        lines = [
            f"• Összes mennyiség: {format_quantity_auto_unit(total_qty, 2)}",
            f"• Cégek száma: {company_count}",
            f"• Hulladékkódok száma: {waste_count}",
            f"• Termékcsoportok száma: {category_count}",
            f"• Veszélyességi arány: {format_hu_number(dangerous_ratio, 2)} %",
            (f"• Vizsgált évek: {years[0]} – {years[-1]}" if len(years) > 1 else
             (f"• Vizsgált év: {years[0]}" if years else "• Vizsgált évek: nincs adat")),
            "",
            "ℹ️ A termékcsoport összesítések átfedhetnek,\n"
            "  mert egy kód több kategóriába is tartozhat.",
        ]
        self.exec_text.setPlainText("\n".join(lines))

    def refresh_company_info(self):
        selected = self.filter_company.checked_texts()
        if len(selected) != 1:
            self.company_info.setPlainText("Válassz ki pontosan 1 céget.")
            return

        company = selected[0]
        if self.source_df.empty or "CEGNEV" not in self.source_df.columns:
            self.company_info.setPlainText("Nincs forrás adat.")
            return

        sub = self.source_df[self.source_df["CEGNEV"].astype(str) == company]
        if sub.empty:
            self.company_info.setPlainText("Nincs forrás adat.")
            return

        row = sub.iloc[0].to_dict()
        address = split_address_parts(row.get("SZEKHELY"))

        lines = [
            f"Cég: {safe_str(row.get('CEGNEV'))}",
            f"KÜJ: {safe_str(row.get('KUJ'))}",
            f"Adószám: {safe_str(row.get('ADOSZAM'))}",
            f"Megye: {safe_str(address.get('Megye'))}",
            f"IRSZ: {safe_str(address.get('IRSZ'))}",
            f"Település: {safe_str(address.get('Település'))}",
            f"Cím: {safe_str(address.get('Cím'))}",
            f"Email: {safe_str(row.get('SRC_EMAIL'))}",
            f"Telefon: {format_phone(row.get('SRC_TELEFON'))}",
            f"Web: {safe_str(row.get('SRC_WEB'))}",
            f"Saját tőke: {safe_str(row.get('SRC_SAJAT_TOKE'))}",
            f"Nettó árbevétel: {safe_str(row.get('SRC_NETTO_ARBEV'))}",
        ]
        self.company_info.setPlainText("\n".join(lines))

    # -------------------------------------------------------------------------
    # Snapshot
    # -------------------------------------------------------------------------
    def save_snapshot(self):
        if self.source_df.empty and self.raw_df.empty:
            QMessageBox.warning(self, "Nincs adat", "Nincs mit menteni.")
            return

        note, ok = QInputDialog.getText(self, "Mentés", "Elnevezés:")
        if not ok:
            return

        def operation():
            try:
                sid = self.persistence.save_snapshot(
                    self.source_df, self.raw_df, self.missing_df,
                    source_file=self.current_file_path.strip(), note=note,
                )
                self.persistence.set_setting("desktop_active_snapshot", str(sid))
                return sid
            finally:
                self.persistence.close()

        self._run_task("Helyi mentés készítése...", operation,
                       lambda sid: self.set_status(f"Helyi mentés elkészült: #{sid}"))

    def _restore_last_snapshot(self):
        if self._is_busy() or not self.raw_df.empty or not self.source_df.empty:
            return
        try:
            snapshots = self.persistence.list_snapshots()
            if snapshots:
                active = self.persistence.get_setting("desktop_active_snapshot") or self.persistence.get_setting("web_active_snapshot")
                selected = next((row for row in snapshots if str(row[0]) == active), snapshots[0])
                self._load_snapshot(selected[0], selected[2])
        except Exception as exc:
            self.set_status(f"Mentésbetöltési hiba: {exc}")
            QMessageBox.critical(self, "Mentésbetöltési hiba", str(exc))

    def _load_snapshot(self, snapshot_id, filename):
        def operation():
            try:
                source, raw, missing = self.persistence.load_snapshot(snapshot_id)
                aggregates = aggregate_data(raw, source, self.category_service)
                self.persistence.set_setting("desktop_active_snapshot", str(snapshot_id))
                return source, raw, missing, aggregates
            finally:
                self.persistence.close()

        def completed(result):
            self.source_df, self.raw_df, self.missing_df, self.agg = result
            self._ktj_confirmed = False
            self._drill_category = None
            self.current_file_path = filename or ""
            self.current_file_label.setText(filename or "-")
            self.current_type_label.setText("Mentett adatok")
            self.filter_type.set_all_checked(False)
            self.populate_filters()
            self.apply_default_year_selection()
            self.refresh_all_views()
            self.set_status(f"Korábbi mentés betöltve: #{snapshot_id}")

        self._run_task("Korábbi mentés betöltése a háttérben...", operation, completed)

    def load_snapshot(self):
        snapshots = self.persistence.list_snapshots()
        if not snapshots:
            QMessageBox.information(self, "Nincs mentés", "Nincsenek mentett snapshotok.")
            return

        items = []
        for row in snapshots:
            sid = row[0]
            created = row[1]

            # Mentés neve (note)
            note = row[3] if len(row) > 3 else (row[2] if len(row) > 2 else "")

            created_txt = ""
            try:
                if hasattr(created, "strftime"):
                    created_txt = created.strftime("%Y.%m.%d %H:%M")
                else:
                    s = str(created).strip()
                    # "2026-05-15T14:08:24.498583" / "2026-05-15 14:08:24" stb.
                    s = s.replace("Z", "").replace("T", " ")
                    dt = pd.to_datetime(s, errors="coerce")
                    created_txt = dt.strftime("%Y.%m.%d %H:%M") if pd.notna(dt) else s
            except Exception:
                created_txt = str(created)

            items.append(f"{sid} | {created_txt} | {note or '-'}")

        selected, ok = QInputDialog.getItem(self, "Helyi mentések", "Korábbi mentések/nézetek:", items, 0, False)
        if not ok or not selected:
            return

        snapshot_id = int(selected.split("|")[0].strip())

        filename = next(row[2] for row in snapshots if row[0] == snapshot_id)
        self._load_snapshot(snapshot_id, filename)

    def delete_snapshot(self):
        snapshots = self.persistence.list_snapshots()
        if not snapshots:
            QMessageBox.information(self, "Nincs mentés", "Nincsenek mentett snapshotok.")
            return

        items = [f"{sid} | {created} | {src or '-'} | {note or '-'}" for sid, created, src, note in snapshots]
        selected, ok = QInputDialog.getItem(self, "Törlés", "Válassz snapshotot törléshez:", items, 0, False)
        if not ok or not selected:
            return

        snapshot_id = int(selected.split("|")[0].strip())

        reply = QMessageBox.question(
            self,
            "Megerősítés",
            f"Biztos törlöd a snapshotot? (ID: {snapshot_id})\nA művelet nem visszavonható.",
        )
        if reply != QMessageBox.Yes:
            return

        try:
            self.persistence.delete_snapshot(snapshot_id)
            self.set_status(f"Snapshot törölve: {snapshot_id}")
            QMessageBox.information(self, "Siker", f"Snapshot törölve: {snapshot_id}")
        except Exception as e:
            QMessageBox.critical(self, "Hiba", str(e))

    def _is_busy(self):
        return self._task_worker is not None or (
            getattr(self, "worker", None) is not None and self.worker.isRunning()
        )

    def _set_busy(self, busy):
        self.menuBar().setEnabled(not busy)
        self.progress.setVisible(busy)
        self.progress.setRange(0, 0 if busy else 100)
        if not busy:
            self.progress.setValue(0)

    def _run_task(self, message, operation, on_success):
        if self._is_busy():
            QMessageBox.warning(self, "Folyamatban", "Várd meg a folyamatban lévő műveletet.")
            return
        self._task_success = on_success
        self._task_result = None
        self._task_error = None
        self._task_worker = TaskWorker(operation, self)
        self._task_worker.succeeded.connect(self._task_succeeded)
        self._task_worker.failed.connect(self._task_failed)
        self._task_worker.finished.connect(self._task_finished)
        self._set_busy(True)
        self.set_status(message)
        self._task_worker.start()

    def _task_succeeded(self, result):
        self._task_result = result

    def _task_failed(self, error):
        self._task_error = error

    def _task_finished(self):
        callback, result, error = self._task_success, self._task_result, self._task_error
        self._task_worker.deleteLater()
        self._task_worker = None
        self._task_success = None
        self._set_busy(False)
        if error is not None:
            self._ktj_confirmed = False
            self._dirty_views.add(2)
            self.set_status(f"Hiba: {error}")
            QMessageBox.critical(self, "Műveleti hiba", error)
        else:
            callback(result)
            if not self._is_busy() and self.tabs.currentIndex() == 2:
                self.on_tab_changed(2)

    def _export_excel(self, filename, writer, *args, **kwargs):
        self._export_file(filename, "Excel (*.xlsx)", writer, *args, **kwargs)

    def _export_file(self, filename, file_filter, writer, *args, **kwargs):
        if self._is_busy():
            QMessageBox.warning(self, "Folyamatban", "Várd meg a folyamatban lévő műveletet.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Export mentése", str(OUTPUT_DIR / filename), file_filter,
        )
        if not path:
            return
        suffix = Path(filename).suffix
        if not path.lower().endswith(suffix):
            path += suffix

        def operation():
            try:
                return atomic_export(path, writer, *args, **kwargs)
            finally:
                self.persistence.close()

        self._run_task("Export készítése a háttérben...", operation, self._export_done)

    def export_portable(self, kind):
        if self.source_df.empty and self.raw_df.empty and self.missing_df.empty:
            QMessageBox.warning(self, "Nincs adat", "Nincs exportálható adatállomány.")
            return
        if kind in {"pkg", "json", "dataset"}:
            tables = (self.source_df.copy(), self.raw_df.copy(), self.missing_df.copy())
            if kind == "pkg":
                state = self.capture_state()

                def writer(path, source, raw, missing):
                    write_package(
                        path, source, raw, missing, state, self.ktj_service.export_metadata(raw),
                        enforce_limits=False,
                    )

                self._export_file("meh-megoszthato.okirpkg", "MÉH PKG (*.okirpkg)", writer, *tables)
            elif kind == "json":
                self._export_file(
                    "meh-adatok.json", "MÉH JSON (*.json)",
                    lambda path, *frames: Path(path).write_bytes(dataset_bytes(*frames)), *tables,
                )
            else:
                self._export_excel("meh-teljes-adatallomany.xlsx", write_dataset_excel, *tables)
        else:
            raise ValueError("Ismeretlen export.")

    def export_configured(self):
        if self._is_busy():
            QMessageBox.warning(self, "Folyamatban", "Várd meg a folyamatban lévő műveletet.")
            return
        if self.source_df.empty and self.raw_df.empty and self.missing_df.empty:
            QMessageBox.warning(self, "Nincs adat", "Nincs exportálható adatállomány.")
            return
        try:
            counties = available_counties(self.source_df, self.raw_df, self.missing_df)
        except ValueError as error:
            QMessageBox.critical(self, "Megyeadatok hibája", str(error))
            return
        description = self.filter_description()
        dialog = ExcelExportDialog(counties, description, self)
        if dialog.exec() != QDialog.Accepted:
            return
        options = dialog.options()
        if options.separate_counties:
            path = QFileDialog.getExistingDirectory(self, "Excel-fájlok célmappája", str(OUTPUT_DIR))
        else:
            path, _ = QFileDialog.getSaveFileName(
                self, "Excel mentése", str(OUTPUT_DIR / "meh-egyedi.xlsx"), "Excel (*.xlsx)",
            )
            if path and not path.lower().endswith(".xlsx"):
                path += ".xlsx"
        if not path:
            return
        source, raw, missing, master = (
            self.source_df.copy(), self.raw_df.copy(), self.missing_df.copy(), self.agg.master.copy(),
        )
        filters = self.current_filters()
        groups = self.category_service.all_groups()
        self._run_task(
            "Excel készítése a háttérben…",
            lambda: export_configured_excel(
                path, source, raw, missing, master, options, filters, groups, description,
            ),
            self._export_done,
        )

    def _export_done(self, path):
        self.set_status(f"Export elkészült: {path}", 100, 100)
        QMessageBox.information(self, "Export elkészült", path)

    def filter_description(self):
        return (
            f"Cégek: {', '.join(self.filter_company.checked_texts()) or 'Összes'}; "
            f"Évek: {', '.join(self.filter_year.checked_texts()) or 'Összes'}; "
            f"Termékcsoportok: {', '.join(self.filter_category.checked_texts()) or 'Összes'}; "
            f"HAK: {self.filter_code.text().strip() or 'Összes'}; "
            f"Adattípusok: {', '.join(self.filter_type.checked_texts()) or 'Összes'}"
        )

    def export_hak_matrix(self):
        master = self.get_filtered_master()
        if master.empty:
            QMessageBox.warning(self, "Nincs adat", "Nincs a szűrésnek megfelelő exportálható adat.")
            return
        self._export_excel("ceg_ev_hak_matrix.xlsx", write_hak_matrix_excel,
                           master.copy(), self.filter_description(), by_year=True)

    def export_missing_excel(self):
        self._export_excel("kimaradt_lekeresek.xlsx", write_missing_excel, self.missing_df.copy())

    # -------------------------------------------------------------------------
    # Excel exportok
    # -------------------------------------------------------------------------
    def export_summary_excel(self):
        master = self.get_filtered_master()
        if master is None or master.empty:
            QMessageBox.warning(self, "Nincs adat", "Nincs exportálható adat.")
            return

        groups = [g for g in self.category_service.all_groups() if g != "Nem besorolható"] + ["Nem besorolható"]
        self._export_excel("okir_summary_export.xlsx", write_summary_excel,
                           master.copy(), self.source_df.copy(), groups)

    def export_detail_excel(self):
        master = self.get_filtered_master()
        if master is None or master.empty:
            QMessageBox.warning(self, "Nincs adat", "Nincs exportálható adat.")
            return

        self._export_excel("okir_detail_export.xlsx", write_detail_excel,
                           master.copy(), self.source_df.copy())

    def export_ktj_excel(self):
        df = self.get_filtered_raw()
        if df.empty:
            QMessageBox.warning(self, "Nincs adat", "Nincs exportálható adat.")
            return

        self._export_excel("okir_ktj_export.xlsx", write_ktj_excel,
                           df.copy(), self.ktj_service, self.category_service, load_metadata=False)

    # -------------------------------------------------------------------------
    # Cache
    # -------------------------------------------------------------------------
    def clear_cache(self):
        reply = QMessageBox.question(self, "Megerősítés", "Biztos törlöd a cache teljes tartalmát?")
        if reply == QMessageBox.Yes:
            self.cache.clear_all()
            self.set_status("A gyorsítótár törölve lett.")
            QMessageBox.information(self, "Kész", "A gyorsítótár törölve lett.")

    def import_json(self):
        self._import_portable(False)

    def import_shared_package(self):
        self._import_portable(True)

    def _import_portable(self, package):
        if self._is_busy():
            QMessageBox.warning(self, "Folyamatban", "Várd meg a folyamatban lévő lekérést.")
            return
        file_filter = "MÉH PKG (*.okirpkg *.pkg *.zip)" if package else "MÉH JSON (*.json)"
        if package:
            paths, _ = QFileDialog.getOpenFileNames(self, "PKG mentések betöltése", "", file_filter)
        else:
            path, _ = QFileDialog.getOpenFileName(self, "Mentés betöltése", "", file_filter)
            paths = [path] if path else []
        if not paths:
            return
        path_label = "; ".join(paths)
        allow_legacy = False
        if package:
            try:
                legacy = False
                for path in paths:
                    with zipfile.ZipFile(path) as archive:
                        legacy |= "dataset.json" not in archive.namelist()
            except (OSError, ValueError, zipfile.BadZipFile) as exc:
                QMessageBox.critical(self, "PKG betöltési hiba", f"{Path(path).name}: {exc}")
                return
            if legacy:
                answer = QMessageBox.question(
                    self, "Régi PKG betöltése",
                    "A kijelölt fájlok között régi, pickle-alapú csomag is van. "
                    "Csak saját vagy megbízható forrásból származó fájlokat tölts be. "
                    "A korlátozott olvasó nem támogat minden pandas-típust. "
                    "Folytatod a betöltést?",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
                )
                if answer != QMessageBox.Yes:
                    return
                allow_legacy = True
        def operation():
            state, metadata = {}, {}
            if package:
                source, raw, missing, state, metadata = read_packages(
                    paths, allow_legacy=allow_legacy, enforce_limits=False,
                )
            else:
                source, raw, missing = read_dataset(Path(paths[0]).read_bytes())
            if source.empty and raw.empty and missing.empty:
                raise ValueError("A fájl nem tartalmaz feldolgozható adatot.")
            aggregates = aggregate_data(raw, source, self.category_service)
            try:
                for ktj, payload in metadata.items():
                    self.persistence.save_ktj_meta(ktj, payload)
                sid = self.persistence.save_snapshot(
                    source, raw, missing, path_label, "PKG import" if package else "JSON import",
                )
                self.persistence.set_setting("desktop_active_snapshot", str(sid))
            finally:
                self.persistence.close()
            return source, raw, missing, aggregates, state, metadata

        def completed(result):
            source, raw, missing, aggregates, state, metadata = result
            self.ktj_service.cache.update(metadata)
            self.current_file_path = path_label
            self.current_file_label.setText(path_label)
            years = raw["EV"].dropna() if not raw.empty else pd.Series(dtype=float)
            self.on_fetch_finished(source, raw, missing, int(years.max()) if not years.empty else None, aggregates)
            if len(paths) > 1:
                self.reset_filters()
            elif state:
                self.restore_state({**state, "current_file_path": path_label})

        self._run_task("Mentés betöltése...", operation, completed)

    def export_report(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Riport mentése", str(OUTPUT_DIR / "meh-riport.html"), "HTML (*.html)",
        )
        if not path:
            return
        try:
            master = self.get_filtered_master()
            Path(path).write_text(report_html(master, self.missing_df, self.filter_description()), encoding="utf-8")
            self.set_status(f"Export elkészült: {path}")
        except Exception as exc:
            QMessageBox.critical(self, "Export hiba", str(exc))

    def closeEvent(self, event):
        if self._is_busy():
            QMessageBox.warning(self, "Folyamatban", "Bezárás előtt várd meg a művelet befejezését.")
            event.ignore()
            return
        self.persistence.close()
        self.cache.close()
        super().closeEvent(event)