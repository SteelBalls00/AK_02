from PyQt5.QtCore import (
    QAbstractTableModel,
    Qt,
    QModelIndex
)
from PyQt5.QtGui import QFont


class TableModel(QAbstractTableModel):
    def __init__(self, table_data=None, parent=None):
        super().__init__(parent)

        self._columns = []
        self._rows = []
        self._total = []
        self.headers = []
        self.tooltips = []

        if table_data:
            self.set_table_data(table_data)

    # ---------- public API ----------

    def set_table_data(self, table_data):
        """
        Полное обновление данных таблицы
        """
        self.beginResetModel()

        self._columns = table_data.get("columns", [])
        self._rows = table_data.get("rows", [])
        self._total = table_data.get("total", [])

        self.headers = table_data["columns"]
        self.tooltips = table_data.get("tooltips", self.headers)
        self._data = table_data["rows"]

        self.endResetModel()

    # ---------- required overrides ----------

    def rowCount(self, parent=None):
        if self._total:
            return len(self._rows) + 1
        return len(self._rows)

    def columnCount(self, parent=QModelIndex()):
        return len(self._columns)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None

        row = index.row()
        col = index.column()

        # --- выбираем правильный источник данных
        if row < len(self._rows):
            row_data = self._rows[row]
            is_total_row = False
        else:
            row_data = self._total
            is_total_row = True

        # защита от выхода за границы
        if col >= len(row_data):
            return None

        value = row_data[col]

        if role == Qt.DisplayRole:
            return value

        # --- жирный шрифт для итога
        if role == Qt.FontRole and is_total_row:
            font = QFont()
            font.setBold(True)
            return font

        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation != Qt.Horizontal:
            return None

        if role == Qt.DisplayRole:
            return self.headers[section]

        if role == Qt.ToolTipRole:
            return self.tooltips[section]

        return None

    # ---------- sorting ----------

    def sort(self, column, order=Qt.AscendingOrder):
        if not self._rows:
            return

        reverse = order == Qt.DescendingOrder

        def sort_key(row):
            value = row[column]
            # если это число в строке — пробуем привести
            try:
                return int(value)
            except Exception:
                return str(value)

        self.layoutAboutToBeChanged.emit()
        self._rows.sort(key=sort_key, reverse=reverse)
        self.layoutChanged.emit()

    # ---------- helpers ----------

    def _is_total_row(self, row):
        return self._total and row == len(self._rows)

    def _get_value(self, row, col):
        if self._is_total_row(row):
            return self._total[col]
        return self._rows[row][col]
