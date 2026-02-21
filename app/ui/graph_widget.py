from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QListWidget, QListWidgetItem,
    QComboBox, QCheckBox, QLabel,
    QDateEdit
)
from PyQt5.QtGui import QColor
from PyQt5.QtCore import Qt, pyqtSignal, QDate
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.cm as cm
from datetime import datetime


class ColorCheckItem(QWidget):

    def __init__(self, text, color):
        super().__init__()

        self.checkbox = QCheckBox()
        self.checkbox.setChecked(True)

        self.color_label = QLabel()
        self.color_label.setFixedSize(14, 14)
        self.color_label.setStyleSheet(
            f"background-color: rgb({int(color[0]*255)},"
            f"{int(color[1]*255)},"
            f"{int(color[2]*255)});"
            "border: 1px solid black;"
        )

        self.text_label = QLabel(text)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(2, 2, 2, 2)
        layout.setSpacing(6)

        layout.addWidget(self.checkbox)
        layout.addWidget(self.color_label)
        layout.addWidget(self.text_label)
        layout.addStretch()


class GraphWidget(QWidget):
    point_clicked = pyqtSignal(dict)  # Это сигнал, который будет передавать информацию о точке
    week_double_clicked = pyqtSignal(int)  # индекс недели

    def __init__(self, parent=None):
        super().__init__(parent)

        self.raw_data = None
        self.processor = None
        self.weeks = []
        self._week_dates = []
        self.judge_colors = {}
        self.category_colors = {}

        self._init_ui()

    # ---------------- UI ----------------

    def _init_ui(self):
        layout = QHBoxLayout(self)

        # ===== LEFT PANEL =====
        left_panel = QVBoxLayout()

        self.category_combo = QComboBox()
        self.category_combo.currentIndexChanged.connect(self.update_chart)

        left_panel.addWidget(QLabel("Показатель:"))
        left_panel.addWidget(self.category_combo)

        self.compare_mode = QCheckBox("Сравнение категорий")
        self.compare_mode.stateChanged.connect(self._toggle_compare_mode)
        left_panel.addWidget(self.compare_mode)

        left_panel.addWidget(QLabel("Судьи:"))

        self.judges_list = QListWidget()
        self.judges_list.itemChanged.connect(self.update_chart)
        left_panel.addWidget(self.judges_list)

        self.total_checkbox = QCheckBox("Показывать линию 'Всего'")
        self.total_checkbox.setChecked(True)
        self.total_checkbox.stateChanged.connect(self.update_chart)
        left_panel.addWidget(self.total_checkbox)

        # Диапазон дат
        left_panel.addWidget(QLabel("Период:"))

        self.date_from = QDateEdit()
        self.date_to = QDateEdit()

        self.date_from.setCalendarPopup(True)
        self.date_to.setCalendarPopup(True)

        self.date_from.dateChanged.connect(self.update_chart)
        self.date_to.dateChanged.connect(self.update_chart)

        left_panel.addWidget(self.date_from)
        left_panel.addWidget(self.date_to)



        self.categories_list = QListWidget()
        self.categories_list.itemChanged.connect(self.update_chart)
        self.categories_list.hide()

        left_panel.addWidget(self.categories_list)



        left_panel.addStretch()

        # ===== RIGHT PANEL =====
        self.figure = Figure()
        self.canvas = FigureCanvas(self.figure)

        layout.addLayout(left_panel, 2)
        layout.addWidget(self.canvas, 5)

        self.canvas.mpl_connect("pick_event", self.on_pick)

    # ---------------- PUBLIC API ----------------

    def set_data(self, raw_data, processor):
        self.raw_data = raw_data
        self.processor = processor

        all_weeks = sorted(
            raw_data.keys(),
            key=lambda w: datetime.strptime(w.split(" - ")[0], "%d.%m.%Y")
        )

        # берём только последние 20
        self.weeks = all_weeks[-20:]

        self._parse_week_dates()

        # 🚨 СНАЧАЛА создаём список имён
        judges = set()
        for week in self.raw_data.values():
            judges.update(week.keys())

        judges = sorted(judges)

        # 🚨 Назначаем цвета ДО создания виджетов
        self._assign_fixed_colors(judges)

        # Теперь создаём UI
        self._fill_categories()
        self.category_combo.currentIndexChanged.connect(self._fill_judges)
        self._fill_judges()

        self.update_chart()

    # ---------------- DATA PREP ----------------

    def _parse_week_dates(self):
        self._week_dates = []

        for w in self.weeks:
            start_str, end_str = w.split(" - ")
            start = datetime.strptime(start_str, "%d.%m.%Y")
            end = datetime.strptime(end_str, "%d.%m.%Y")
            self._week_dates.append((start, end))

        if not self._week_dates:
            return

        first = self._week_dates[0][0]
        last = self._week_dates[-1][1]

        # 🚨 ВАЖНО: блокируем сигналы
        self.date_from.blockSignals(True)
        self.date_to.blockSignals(True)

        self.date_from.setDate(QDate(first.year, first.month, first.day))
        self.date_to.setDate(QDate(last.year, last.month, last.day))

        self.date_from.blockSignals(False)
        self.date_to.blockSignals(False)

        print("WEEKS:", len(self.weeks))
        print("FILTERED:", len(self._get_filtered_weeks()))

    def _fill_categories(self):
        self.category_combo.clear()
        self.category_combo.addItems(self.processor.categories)

        self.categories_list.clear()

        for category in self.processor.categories:
            color = self.category_colors.get(category, (0.5, 0.5, 0.5))

            item = QListWidgetItem()
            widget = ColorCheckItem(category, color)

            item.setSizeHint(widget.sizeHint())
            self.categories_list.addItem(item)
            self.categories_list.setItemWidget(item, widget)

            widget.checkbox.stateChanged.connect(self.update_chart)

    def _fill_judges(self):
        self.judges_list.clear()

        category = self.category_combo.currentText()
        judges = set()

        # берём только отображаемые недели
        for week_key in self.weeks:
            week_data = self.raw_data.get(week_key, {})

            for judge, judge_data in week_data.items():
                cases = judge_data.get(category, [])
                if cases:  # есть дела по выбранной категории
                    judges.add(judge)

        judges = sorted(judges)

        for judge in judges:
            color = self.judge_colors.get(judge, (0.5, 0.5, 0.5))

            item = QListWidgetItem()
            widget = ColorCheckItem(judge, color)

            item.setSizeHint(widget.sizeHint())
            self.judges_list.addItem(item)
            self.judges_list.setItemWidget(item, widget)

            widget.checkbox.stateChanged.connect(self.update_chart)

    # ---------------- FILTER ----------------

    def _get_filtered_weeks(self):
        date_from = self.date_from.date().toPyDate()
        date_to = self.date_to.date().toPyDate()

        result = []

        for i, (start, end) in enumerate(self._week_dates):
            start_date = start.date()
            end_date = end.date()

            # проверяем пересечение диапазонов
            if start_date <= date_to and end_date >= date_from:
                result.append((i, self.weeks[i]))

        return result

    def _get_selected_judges(self):
        judges = []

        for i in range(self.judges_list.count()):
            item = self.judges_list.item(i)
            widget = self.judges_list.itemWidget(item)

            if widget.checkbox.isChecked():
                judges.append(widget.text_label.text())

        return judges

    # ---------------- BUILD SERIES ----------------

    def _build_series(self, category, judges, week_indexes):
        series = {}

        for judge in judges:
            values = []

            for i, week in week_indexes:
                week_data = self.raw_data[week]
                judge_data = week_data.get(judge, {})
                cases = judge_data.get(category, [])
                values.append(len(cases))

            series[judge] = values

        return series

    # ---------------- CHART ----------------

    def update_chart(self):
        if not self.raw_data:
            return

        self.figure.clear()
        ax = self.figure.add_subplot(111)

        week_indexes = self._get_filtered_weeks()
        if not week_indexes:
            self.canvas.draw()
            return

        judges = self._get_selected_judges()

        # ================== ОБЫЧНЫЙ РЕЖИМ ==================
        if not self.compare_mode.isChecked():

            if not judges:
                self.canvas.draw()
                return

            category = self.category_combo.currentText()
            series = self._build_series(category, judges, week_indexes)

            # 🔥 фильтрация судей с нулевыми значениями
            filtered_series = {
                judge: values
                for judge, values in series.items()
                if any(v > 0 for v in values)
            }

            for judge, values in filtered_series.items():
                ax.plot(
                    range(len(values)),
                    values,
                    marker="o",
                    label=judge,
                    color=self.judge_colors.get(judge, "gray"),
                    picker=6
                )

            # линия "Всего"
            if self.total_checkbox.isChecked():
                totals = [
                    sum(filtered_series[j][i] for j in filtered_series)
                    for i in range(len(next(iter(series.values()))))
                ]
                ax.plot(
                    range(len(totals)),
                    totals,
                    linestyle="--",
                    color="black",
                    picker=6
                )

            ax.set_title(category)

        # ================== СРАВНЕНИЕ КАТЕГОРИЙ ==================
        else:

            selected_categories = [
                self.categories_list.itemWidget(
                    self.categories_list.item(i)
                ).text_label.text()
                for i in range(self.categories_list.count())
                if self.categories_list.itemWidget(
                    self.categories_list.item(i)
                ).checkbox.isChecked()
            ]

            if not selected_categories:
                self.canvas.draw()
                return

            for category in selected_categories:
                values = []

                for _, week in week_indexes:
                    total = 0
                    week_data = self.raw_data[week]

                    for judge in judges:
                        cases = week_data.get(judge, {}).get(category, [])
                        total += len(cases)

                    values.append(total)

                ax.plot(
                    range(len(values)),
                    values,
                    marker="o",
                    label=category,
                    color=self.category_colors.get(category, "gray"),
                    picker=6
                )

            ax.set_title("Сравнение категорий")

        # ===== X ось =====
        ax.set_xticks(range(len(week_indexes)))
        ax.set_xticklabels(
            [w[-10:] for _, w in week_indexes],
            rotation=90
        )

        ax.grid(True)

        self.figure.tight_layout()
        self.canvas.draw()

    # ---------------- DOUBLE CLICK ----------------

    def on_click(self, event):

        # 1️⃣ Проверяем что клик внутри графика
        if not event.inaxes:
            return

        # 2️⃣ Только левая кнопка мыши
        if event.button != 1:
            return

        # 3️⃣ Координата X (индекс точки)
        x = round(event.xdata)

        week_indexes = self._get_filtered_weeks()

        # 4️⃣ Проверяем что индекс существует
        if 0 <= x < len(week_indexes):
            original_index = week_indexes[x][0]
            week_key = week_indexes[x][1]

            category = self.category_combo.currentText()
            judges = self._get_selected_judges()

            data = {
                "week_index": original_index,
                "week_key": week_key,
                "category": category,
                "judges": judges
            }

            self.point_clicked.emit(data)

    def on_pick(self, event):

        line = event.artist
        mouse_event = event.mouseevent

        ind = event.ind[0]

        week_indexes = self._get_filtered_weeks()

        if not (0 <= ind < len(week_indexes)):
            return

        _, week_key = week_indexes[ind]

        judge = line.get_label()
        category = self.category_combo.currentText()

        ydata = line.get_ydata()
        value = int(ydata[ind])

        data = {
            "week_key": week_key,
            "category": category,
            "judge": judge,
            "value": value,
            "double_click": mouse_event.dblclick
        }

        self.point_clicked.emit(data)

    def _assign_fixed_colors(self, judges):
        self.judge_colors.clear()
        self.category_colors.clear()

        # ---------- Судьи ----------
        cmap = cm.get_cmap("tab20")

        for idx, judge in enumerate(judges):
            self.judge_colors[judge] = cmap(idx % 20)

        # ---------- Категории ----------
        categories = list(self.processor.categories)
        cmap2 = cm.get_cmap("Set2")

        for idx, cat in enumerate(categories):
            self.category_colors[cat] = cmap2(idx % 8)

        # ---------- Категории ----------
        categories = list(self.processor.categories)

        cmap2 = cm.get_cmap("Set2")

        for idx, cat in enumerate(categories):
            self.category_colors[cat] = cmap2(idx % 8)

    def _toggle_compare_mode(self):
        if self.compare_mode.isChecked():
            self.category_combo.hide()
            self.judges_list.hide()
            self.total_checkbox.hide()
            self.categories_list.show()
        else:
            self.category_combo.show()
            self.judges_list.show()
            self.total_checkbox.show()
            self.categories_list.hide()

        self.update_chart()