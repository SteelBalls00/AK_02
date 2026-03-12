from PyQt5.QtWidgets import QTableView, QAbstractItemView
from PyQt5.QtCore import Qt


class FrozenTableView(QTableView):
    """
    QTableView с закреплённым первым столбцом (Excel-style).
    """

    def __init__(self, parent=None):
        super().__init__(parent)

        self.frozen_view = QTableView(self)

        self._init_frozen_view()

        # синхронизация скролла
        self.verticalScrollBar().valueChanged.connect(
            self.frozen_view.verticalScrollBar().setValue
        )
        self.frozen_view.verticalScrollBar().valueChanged.connect(
            self.verticalScrollBar().setValue
        )

        # обновление геометрии
        self.horizontalHeader().sectionResized.connect(self.update_frozen_geometry)
        self.verticalHeader().sectionResized.connect(self.update_frozen_geometry)

        self.verticalHeader().sectionResized.connect(self.sync_row_height)

    def sync_row_height(self, row, old_height, new_height):
        self.frozen_view.setRowHeight(row, new_height)

    def _init_frozen_view(self):

        self.frozen_view.setFocusPolicy(Qt.NoFocus)
        self.frozen_view.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.frozen_view.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        self.frozen_view.setEditTriggers(QAbstractItemView.NoEditTriggers)

        # одинаковый режим скролла
        self.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.frozen_view.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)

        # убираем лишнее
        self.frozen_view.verticalHeader().hide()
        self.frozen_view.setStyleSheet(
            "QTableView { border: none; background-color: palette(base); }"
        )

        # чтобы frozen был поверх viewport
        self.viewport().stackUnder(self.frozen_view)

    def setModel(self, model):
        super().setModel(model)

        self.frozen_view.setModel(model)

        # selection model можно ставить только после модели
        self.frozen_view.setSelectionModel(self.selectionModel())

        # скрываем все столбцы кроме первого
        for col in range(1, model.columnCount()):
            self.frozen_view.setColumnHidden(col, True)

        # одинаковая высота строк
        for row in range(model.rowCount()):
            height = self.rowHeight(row)
            self.frozen_view.setRowHeight(row, height)

        self.update_frozen_geometry()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.update_frozen_geometry()

    def scrollTo(self, index, hint=QAbstractItemView.EnsureVisible):
        if index.column() > 0:
            super().scrollTo(index, hint)

    def update_frozen_geometry(self):
        if not self.model():
            return

        width = self.columnWidth(0)
        self.frozen_view.setColumnWidth(0, width)

        self.frozen_view.setGeometry(
            self.verticalHeader().width(),
            self.frameWidth(),
            width,
            self.viewport().height() + self.horizontalHeader().height()
        )

    def setColumnWidth(self, column, width):
        super().setColumnWidth(column, width)

        if column == 0:
            self.update_frozen_geometry()