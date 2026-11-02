import math
from pathlib import Path
from typing import Any, Callable, List, Optional

from PySide6.QtCore import QModelIndex, Qt, Signal
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QStandardItem, QStandardItemModel
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSizePolicy,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

from core.constants import DEFAULT_QUERY_TYPE_LABEL
from core.logger import logger
from core.utils import format_hu_number, format_quantity_auto_unit
from services.configurable_export_service import EXCEL_SECTIONS, ExcelExportOptions


# ─────────────────────────────────────────────────────────────────────────────
class LogDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Műveleti napló")
        self.resize(900, 620)
        layout = QVBoxLayout(self)
        self.text = QTextEdit()
        self.text.setReadOnly(True)
        self.text.setPlainText(logger.text())
        layout.addWidget(self.text)
        row = QHBoxLayout()
        row.addStretch()
        rb = QPushButton("Frissítés"); rb.clicked.connect(self._refresh); row.addWidget(rb)
        cb = QPushButton("Bezárás");  cb.clicked.connect(self.accept);   row.addWidget(cb)
        layout.addLayout(row)

    def _refresh(self):
        self.text.setPlainText(logger.text())


# ─────────────────────────────────────────────────────────────────────────────
class MultiSelectComboBox(QComboBox):
    selectionChanged = Signal()

    def __init__(self, placeholder: str = "Összes", parent=None):
        super().__init__(parent)
        self.setEditable(True)
        self.lineEdit().setReadOnly(True)
        self._placeholder = placeholder
        self.lineEdit().setText(placeholder)
        self._model = QStandardItemModel(self)
        self.setModel(self._model)
        self.view().clicked.connect(self._on_clicked)
        self._update_summary()

    def clear_items(self):
        self._model.clear()
        self._update_summary()

    def add_check_item(self, text: str, checked: bool = False, user_data: Any = None):
        item = QStandardItem(text)
        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
        item.setData(Qt.Checked if checked else Qt.Unchecked, Qt.CheckStateRole)
        item.setData(user_data if user_data is not None else text, Qt.UserRole)
        self._model.appendRow(item)
        self._update_summary()

    def add_check_items(self, items):
        for text, value in items:
            item = QStandardItem(text)
            item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
            item.setData(Qt.Unchecked, Qt.CheckStateRole)
            item.setData(value, Qt.UserRole)
            self._model.appendRow(item)
        self._update_summary()

    def checked_values(self) -> List[Any]:
        return [
            self._model.item(i).data(Qt.UserRole)
            for i in range(self._model.rowCount())
            if self._model.item(i).data(Qt.CheckStateRole) == Qt.Checked
        ]

    def checked_texts(self) -> List[str]:
        return [
            self._model.item(i).text()
            for i in range(self._model.rowCount())
            if self._model.item(i).data(Qt.CheckStateRole) == Qt.Checked
        ]

    def set_all_checked(self, checked: bool):
        state = Qt.Checked if checked else Qt.Unchecked
        for i in range(self._model.rowCount()):
            self._model.item(i).setData(state, Qt.CheckStateRole)
        self._update_summary()
        self.selectionChanged.emit()

    def set_checked_by_user_data(self, allowed):
        allowed_set = set(allowed)
        for i in range(self._model.rowCount()):
            item = self._model.item(i)
            item.setData(
                Qt.Checked if item.data(Qt.UserRole) in allowed_set else Qt.Unchecked,
                Qt.CheckStateRole,
            )
        self._update_summary()
        self.selectionChanged.emit()

    def _on_clicked(self, index: QModelIndex):
        item = self._model.itemFromIndex(index)
        if not item:
            return
        item.setData(
            Qt.Unchecked if item.data(Qt.CheckStateRole) == Qt.Checked else Qt.Checked,
            Qt.CheckStateRole,
        )
        self._update_summary()
        self.selectionChanged.emit()
        self.showPopup()

    def _update_summary(self):
        checked = self.checked_texts()
        if not checked:
            self.lineEdit().setText(self._placeholder)
        elif len(checked) == 1:
            self.lineEdit().setText(checked[0])
        else:
            self.lineEdit().setText(f"{len(checked)} kiválasztva")


# ─────────────────────────────────────────────────────────────────────────────
class ExcelExportDialog(QDialog):
    def __init__(self, counties, description, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Excel készítése")
        self.resize(560, 600)
        layout = QVBoxLayout(self)
        layout.setSpacing(14)
        title = QLabel("Mit tartalmazzon az Excel?")
        title.setObjectName("appTitle")
        layout.addWidget(title)
        group = QGroupBox("Munkalapok")
        sheets = QVBoxLayout(group)
        self.section_checks = {}
        for key, label in EXCEL_SECTIONS.items():
            check = QCheckBox(label)
            check.setChecked(key == "matrix")
            check.toggled.connect(self._update_validity)
            sheets.addWidget(check)
            self.section_checks[key] = check
        layout.addWidget(group)
        self.use_filters = QCheckBox("A főablak aktuális szűrésének használata")
        self.use_filters.setChecked(True)
        layout.addWidget(self.use_filters)
        scope = QLabel(description)
        scope.setWordWrap(True)
        scope.setObjectName("mutedLabel")
        layout.addWidget(scope)
        form = QFormLayout()
        self.counties = MultiSelectComboBox("Összes megye")
        self.counties.add_check_items((county, county) for county in counties)
        form.addRow("Megyék", self.counties)
        layout.addLayout(form)
        self.separate_counties = QCheckBox("Megyénként külön Excel-fájl")
        layout.addWidget(self.separate_counties)
        note = QLabel(
            "A megye a cég székhelye alapján kerül meghatározásra. A nem azonosítható "
            "sorok az Ismeretlen megye csoportba kerülnek.\n"
            "A forráscégekre és hibákra csak a cég- és megyeszűrés vonatkozik.\n"
            "Külön fájloknál a kiválasztott célmappában új exportmappa készül; "
            "korábbi fájlokat nem írunk felül."
        )
        note.setWordWrap(True)
        note.setObjectName("mutedLabel")
        layout.addWidget(note)
        layout.addStretch()
        self.buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        self.buttons.button(QDialogButtonBox.Save).setText("Excel készítése")
        self.buttons.button(QDialogButtonBox.Save).setObjectName("primaryButton")
        self.buttons.button(QDialogButtonBox.Cancel).setText("Mégse")
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)
        self._update_validity()

    def _update_validity(self):
        if hasattr(self, "buttons"):
            self.buttons.button(QDialogButtonBox.Save).setEnabled(
                any(check.isChecked() for check in self.section_checks.values())
            )

    def options(self):
        return ExcelExportOptions(
            sections=tuple(key for key, check in self.section_checks.items() if check.isChecked()),
            use_filters=self.use_filters.isChecked(),
            counties=tuple(self.counties.checked_values()),
            separate_counties=self.separate_counties.isChecked(),
        )


class FileDropWidget(QFrame):
    fileDropped = Signal(str)

    def __init__(self, compact: bool = False, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setObjectName("fileDropWidget")
        self._drag_active = False

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(6)

        self.title = QLabel("Excel fájl behúzása")
        self.title.setObjectName("fileDropTitle")
        self.title.setAlignment(Qt.AlignCenter)

        self.subtitle = QLabel("Húzd ide az .xlsx / .xls fájlt" if compact
                               else "Húzd ide az .xlsx vagy .xls fájlt,\nvagy válaszd ki tallózással.")
        self.subtitle.setObjectName("fileDropSubtitle")
        self.subtitle.setAlignment(Qt.AlignCenter)

        layout.addStretch()
        layout.addWidget(self.title)
        layout.addWidget(self.subtitle)
        layout.addStretch()
        self.setMinimumHeight(90 if compact else 160)
        self._refresh_style()

    def _refresh_style(self):
        self.setProperty("dragActive", self._drag_active)
        self.style().unpolish(self); self.style().polish(self); self.update()

    def _is_valid_file(self, path: str) -> bool:
        return Path(path).suffix.lower() in {".xlsx", ".xls"}

    def dragEnterEvent(self, event: QDragEnterEvent):
        urls = event.mimeData().urls() if event.mimeData().hasUrls() else []
        if urls and self._is_valid_file(urls[0].toLocalFile()):
            self._drag_active = True; self._refresh_style(); event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self._drag_active = False; self._refresh_style(); super().dragLeaveEvent(event)

    def dropEvent(self, event: QDropEvent):
        self._drag_active = False; self._refresh_style()
        urls = event.mimeData().urls() if event.mimeData().hasUrls() else []
        if urls:
            path = urls[0].toLocalFile()
            if self._is_valid_file(path):
                self.fileDropped.emit(path); event.acceptProposedAction(); return
        event.ignore()


# ─────────────────────────────────────────────────────────────────────────────
class ImportDatabaseDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Új adatbázis tallózása")
        self.resize(620, 380)
        self.selected_file = ""

        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        self.drop_widget = FileDropWidget(compact=False)
        self.drop_widget.fileDropped.connect(self.set_selected_file)
        layout.addWidget(self.drop_widget)

        file_row = QHBoxLayout()
        self.file_label = QLabel("Nincs fájl kiválasztva")
        self.file_label.setWordWrap(True)
        browse_btn = QPushButton("Tallózás")
        browse_btn.clicked.connect(self.browse_file)
        file_row.addWidget(self.file_label, 1); file_row.addWidget(browse_btn)
        layout.addLayout(file_row)

        form = QFormLayout()
        self.query_type_combo = QComboBox()
        self.query_type_combo.addItems(["Keletkezés", "Átvétel", "Átadás", "Mindhárom"])
        self.query_type_combo.setCurrentText(DEFAULT_QUERY_TYPE_LABEL)
        self.cache_chk = QCheckBox("Gyorsítótár használata")
        self.cache_chk.setChecked(True)
        form.addRow("Lekérdezés típusa:", self.query_type_combo)
        form.addRow("Beállítás:",         self.cache_chk)
        layout.addLayout(form)

        self.buttons    = QDialogButtonBox()
        self.start_btn  = self.buttons.addButton("Indítás", QDialogButtonBox.AcceptRole)
        self.cancel_btn = self.buttons.addButton("Mégse",   QDialogButtonBox.RejectRole)
        self.start_btn.clicked.connect(self.validate_and_accept)
        self.cancel_btn.clicked.connect(self.reject)
        layout.addWidget(self.buttons)

    def set_selected_file(self, path: str):
        self.selected_file = path
        self.file_label.setText(path)

    def browse_file(self):
        from PySide6.QtWidgets import QFileDialog
        path, _ = QFileDialog.getOpenFileName(self, "Excel kiválasztása", "", "Excel Files (*.xlsx *.xls)")
        if path:
            self.set_selected_file(path)

    def validate_and_accept(self):
        if not self.selected_file:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.warning(self, "Hiányzó fájl", "Válassz ki vagy húzz be egy Excel fájlt.")
            return
        self.accept()

    def payload(self):
        return {
            "file_path":  self.selected_file,
            "query_type": self.query_type_combo.currentText(),
            "use_cache":  self.cache_chk.isChecked(),
        }


# ─────────────────────────────────────────────────────────────────────────────
class HoverAnnotation:
    def __init__(self, ax):
        self.annot = ax.annotate(
            "", xy=(0, 0), xytext=(12, 12), textcoords="offset points",
            bbox=dict(boxstyle="round,pad=0.45", fc="white", ec="#cbd5e1",
                      alpha=0.97, linewidth=1.2),
            arrowprops=dict(arrowstyle="->", color="#94a3b8", lw=0.9),
            fontsize=9, zorder=20,
        )
        self.annot.set_visible(False)

    def show(self, x, y, text: str, x_offset=12, y_offset=12):
        self.annot.xy = (x, y)
        self.annot.set_text(text)
        self.annot.set_position((x_offset, y_offset))
        self.annot.set_visible(True)

    def hide(self):
        self.annot.set_visible(False)


# ─────────────────────────────────────────────────────────────────────────────
class MplCanvas(FigureCanvas):
    doubleClicked  = Signal()
    clickedTarget  = Signal(object)   # payload dict

    def __init__(self, width=6, height=4, dpi=100):
        self.figure = Figure(figsize=(width, height), dpi=dpi, facecolor="white")
        self.ax = self.figure.add_subplot(111)
        super().__init__(self.figure)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.hover_targets     = []
        self.hover_annotation  = HoverAnnotation(self.ax)
        self._last_hover_key   = None
        self.mpl_connect("motion_notify_event", self.on_hover)
        self.mpl_connect("button_press_event",  self.on_click)
        self.mpl_connect("scroll_event",        self.on_scroll)

    def reset_hover(self):
        self.hover_targets    = []
        self.hover_annotation.hide()
        self.hover_annotation = HoverAnnotation(self.ax)
        self._last_hover_key  = None

    # ── click ────────────────────────────────────────────────────────────────
    def on_click(self, event):
        if getattr(event, "dblclick", False):
            self.doubleClicked.emit(); return
        if event.inaxes != self.ax:
            return
        # artists first
        for t in self.hover_targets:
            if t.get("kind") == "point":
                continue
            artist = t.get("artist")
            if artist is None:
                continue
            try:
                ok, _ = artist.contains(event)
            except Exception:
                ok = False
            if ok:
                self.clickedTarget.emit(t.get("payload")); return
        # points
        for t in self.hover_targets:
            if t.get("kind") != "point" or event.xdata is None:
                continue
            px, py = t["point"]
            if abs(event.xdata - px) <= t.get("radius_x", 0.35) and abs(event.ydata - py) <= t.get("radius_y", 1.0):
                self.clickedTarget.emit(t.get("payload")); return

    # ── hover targets ────────────────────────────────────────────────────────
    def add_bar_targets(self, bars, texts, values, horizontal=False,
                        percents=None, payloads=None, show_quantity_line=True):
        payloads = payloads or [None] * len(bars)
        for idx, bar in enumerate(bars):
            if horizontal:
                x = bar.get_width() + getattr(bar, "_x0", 0)
                y = bar.get_y() + bar.get_height() / 2
            else:
                x = bar.get_x() + bar.get_width() / 2
                y = bar.get_y() + bar.get_height()
            text = texts[idx]
            if percents is not None:
                text += f"\nArány: {format_hu_number(percents[idx], 2)} %"
            if show_quantity_line:
                text += f"\nMennyiség: {format_quantity_auto_unit(values[idx], 2)}"
            self.hover_targets.append({
                "artist": bar, "point": (x, y), "text": text,
                "kind": "bar", "payload": payloads[idx],
            })

    def add_point_targets(self, xs, ys, texts, payloads=None):
        payloads = payloads or [None] * len(xs)
        for i in range(len(xs)):
            self.hover_targets.append({
                "artist": None, "point": (xs[i], ys[i]),
                "text": texts[i], "kind": "point",
                "radius_x": 0.35,
                "radius_y": max(abs(ys[i]) * 0.08, 1),
                "payload": payloads[i],
            })

    def add_pie_targets(self, wedges, texts, values, percents=None, payloads=None):
        payloads = payloads or [None] * len(wedges)
        for idx, wedge in enumerate(wedges):
            angle = (wedge.theta1 + wedge.theta2) / 2
            x = math.cos(math.radians(angle)) * 0.7
            y = math.sin(math.radians(angle)) * 0.7
            text = texts[idx]
            if percents is not None:
                text += f"\nArány: {format_hu_number(percents[idx], 2)} %"
            text += f"\nMennyiség: {format_quantity_auto_unit(values[idx], 2)}"
            self.hover_targets.append({
                "artist": wedge, "point": (x, y), "text": text,
                "kind": "pie", "payload": payloads[idx],
            })

    # ── hover render ─────────────────────────────────────────────────────────
    def _tooltip_offsets(self, event):
        if event.x is None or event.y is None:
            return 12, 12
        xo = 12 if event.x < self.width()  * 0.70 else -230
        yo = 12 if event.y > self.height() * 0.25 else  45
        return xo, yo

    def on_hover(self, event):
        if event.inaxes != self.ax:
            if self.hover_annotation.annot.get_visible():
                self.hover_annotation.hide(); self._last_hover_key = None; self.draw_idle()
            return

        shown = False
        xo, yo = self._tooltip_offsets(event)
        hover_key = None

        for t in self.hover_targets:
            kind = t.get("kind", "bar")
            if kind == "point":
                if event.xdata is None or event.ydata is None:
                    continue
                px, py = t["point"]
                if abs(event.xdata - px) <= t.get("radius_x", 0.3) and abs(event.ydata - py) <= t.get("radius_y", 1.0):
                    self.hover_annotation.show(px, py, t["text"], xo, yo)
                    hover_key = ("pt", px, py); shown = True; break
            else:
                try:
                    ok, _ = t["artist"].contains(event)
                except Exception:
                    ok = False
                if ok:
                    px, py = t["point"]
                    self.hover_annotation.show(px, py, t["text"], xo, yo)
                    hover_key = (kind, px, py); shown = True; break

        if not shown:
            self.hover_annotation.hide(); hover_key = None

        if hover_key != self._last_hover_key:
            self._last_hover_key = hover_key; self.draw_idle()

    # ── scroll zoom ──────────────────────────────────────────────────────────
    def on_scroll(self, event):
        if event.inaxes != self.ax:
            return
        scale = 1 / 1.15 if event.button == "up" else 1.15
        xl = self.ax.get_xlim(); yl = self.ax.get_ylim()
        xd = event.xdata or (xl[0]+xl[1])/2
        yd = event.ydata or (yl[0]+yl[1])/2
        nw = (xl[1]-xl[0])*scale; nh = (yl[1]-yl[0])*scale
        rx = (xl[1]-xd)/(xl[1]-xl[0]) if xl[1]!=xl[0] else 0.5
        ry = (yl[1]-yd)/(yl[1]-yl[0]) if yl[1]!=yl[0] else 0.5
        self.ax.set_xlim([xd-nw*(1-rx), xd+nw*rx])
        self.ax.set_ylim([yd-nh*(1-ry), yd+nh*ry])
        self.draw_idle()


# ─────────────────────────────────────────────────────────────────────────────
class FullscreenChartDialog(QDialog):
    def __init__(self, title: str, render_callback: Callable[["MplCanvas"], None], parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(1400, 900)
        layout = QVBoxLayout(self)
        self.canvas = MplCanvas(width=12, height=8)
        layout.addWidget(self.canvas, 1)
        row = QHBoxLayout(); row.addStretch()
        cb = QPushButton("Bezárás"); cb.clicked.connect(self.accept); row.addWidget(cb)
        layout.addLayout(row)
        render_callback(self.canvas)


# ─────────────────────────────────────────────────────────────────────────────
class DashboardCard(QFrame):
    def __init__(self, title: str, subtitle: str = "", accent: str = "#3b82f6",
                 parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setObjectName("dashboardCard")
        self.setMinimumHeight(96)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(3)

        # colour accent stripe (top border)
        self.setStyleSheet(
            f"#dashboardCard {{ border-top: 3px solid {accent}; border-radius: 12px; }}"
        )

        self.title_label = QLabel(title)
        self.title_label.setObjectName("dashboardCardTitle")
        layout.addWidget(self.title_label)

        self.value_label = QLabel("-")
        self.value_label.setObjectName("dashboardCardValue")
        self.value_label.setWordWrap(False)
        layout.addWidget(self.value_label)

        self.sub_label = QLabel(subtitle)
        self.sub_label.setObjectName("dashboardCardSub")
        self.sub_label.setWordWrap(True)
        layout.addWidget(self.sub_label)

    def set_value(self, value: str):
        self.value_label.setText(value)


# ─────────────────────────────────────────────────────────────────────────────
class ChartPanel(QFrame):
    def __init__(self, title: str, expand_callback: Optional[Callable[[], None]] = None, parent=None):
        super().__init__(parent)
        self.expand_callback = expand_callback
        self.setObjectName("chartPanel")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 8)
        layout.setSpacing(4)

        header = QHBoxLayout()
        self.title_label = QLabel(title)
        self.title_label.setObjectName("chartPanelTitle")
        header.addWidget(self.title_label)
        header.addStretch()

        self.expand_btn = QPushButton("⤢")
        self.expand_btn.setFixedSize(30, 24)
        self.expand_btn.setToolTip("Nagyítás")
        self.expand_btn.clicked.connect(self.on_expand)
        header.addWidget(self.expand_btn)
        layout.addLayout(header)

        self.canvas = MplCanvas(width=8, height=5)
        self.canvas.doubleClicked.connect(self.on_expand)
        layout.addWidget(self.canvas, 1)

    def set_title(self, title: str):
        self.title_label.setText(title)

    def on_expand(self):
        if self.expand_callback:
            self.expand_callback()