import pandas as pd
from PySide6.QtCore import QAbstractTableModel, Qt


class DataFrameModel(QAbstractTableModel):
    """Render visible cells only instead of allocating a widget item per cell."""

    def __init__(self, frame, parent=None):
        super().__init__(parent)
        self.frame = frame.reset_index(drop=True)

    def rowCount(self, parent=None):
        return 0 if parent is not None and parent.isValid() else len(self.frame)

    def columnCount(self, parent=None):
        return 0 if parent is not None and parent.isValid() else len(self.frame.columns)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid() or role != Qt.DisplayRole:
            return None
        value = self.frame.iat[index.row(), index.column()]
        return "" if pd.isna(value) else str(value)

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole and orientation == Qt.Horizontal:
            return str(self.frame.columns[section])
        return None

    def sort(self, column, order=Qt.AscendingOrder):
        if not 0 <= column < len(self.frame.columns):
            return
        self.layoutAboutToBeChanged.emit()
        self.frame = self.frame.sort_values(
            self.frame.columns[column], ascending=order == Qt.AscendingOrder,
            kind="stable", na_position="last",
            key=lambda values: values if pd.api.types.is_numeric_dtype(values) else values.astype("string"),
        ).reset_index(drop=True)
        self.layoutChanged.emit()
