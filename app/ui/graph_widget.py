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


class ClickableLabel(QLabel):
    clicked = pyqtSignal()

    def mousePressEvent(self, event):
        self.clicked.emit()
        super().mousePressEvent(event)


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

        self.text_label = ClickableLabel(text)
        self.text_label.clicked.connect(self._toggle_checkbox)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(2, 2, 2, 2)
        layout.setSpacing(6)

        layout.addWidget(self.checkbox)
        layout.addWidget(self.color_label)
        layout.addWidget(self.text_label)
        layout.addStretch()

    def _toggle_checkbox(self):
        self.checkbox.setChecked(not self.checkbox.isChecked())


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
        self._user_range_selected = False
        self._hover_annotation = None
        self._pan_start = None
        self._pan_xlim = None
        self._pan_ylim = None

        self._init_ui()

    # ---------------- UI ----------------

    def _init_ui(self):
        layout = QHBoxLayout(self)

        # ===== LEFT PANEL =====
        left_panel = QVBoxLayout()

        self.category_combo = QComboBox()
        self.category_combo.currentIndexChanged.connect(self.update_chart)

        left_panel.addWidget(self.category_combo)

        self.compare_mode = QCheckBox("Сравнение категорий")
        self.compare_mode.stateChanged.connect(self._toggle_compare_mode)
        left_panel.addWidget(self.compare_mode)

        self.judges_list = QListWidget()
        self.judges_list.itemChanged.connect(self.update_chart)
        left_panel.addWidget(self.judges_list, 1)

        # Диапазон дат
        left_panel.addWidget(QLabel("Период:"))

        self.date_from = QDateEdit()
        self.date_to = QDateEdit()

        self.date_from.setCalendarPopup(True)
        self.date_to.setCalendarPopup(True)

        self.date_from.dateChanged.connect(self._on_date_changed)
        self.date_to.dateChanged.connect(self._on_date_changed)

        left_panel.addWidget(self.date_from)
        left_panel.addWidget(self.date_to)

        self.categories_list = QListWidget()
        self.categories_list.itemChanged.connect(self.update_chart)
        self.categories_list.hide()

        left_panel.addWidget(self.categories_list, 1)

        left_panel.addStretch()

        # ===== RIGHT PANEL =====
        self.figure = Figure()
        self.canvas = FigureCanvas(self.figure)
        self.canvas.mpl_connect("motion_notify_event", self._on_hover)
        self.canvas.mpl_connect("scroll_event", self._on_scroll)
        self.canvas.mpl_connect("button_press_event", self._on_press)
        self.canvas.mpl_connect("button_release_event", self._on_release)
        self.canvas.mpl_connect("motion_notify_event", self._on_pan_motion)

        layout.addLayout(left_panel, 1)
        layout.addWidget(self.canvas, 5)

        self.canvas.mpl_connect("pick_event", self.on_pick)

    def _toggle_all_generic(self, state, target_list):

        checked = state == Qt.Checked

        for i in range(target_list.count()):
            item = target_list.item(i)
            widget = target_list.itemWidget(item)

            text = widget.text_label.text()

            if text.startswith("Показывать") or text.startswith("Отметить"):
                continue

            widget.checkbox.blockSignals(True)
            widget.checkbox.setChecked(checked)
            widget.checkbox.blockSignals(False)

        self.update_chart()

    def _add_bottom_controls(self, target_list, mode="judges"):

        # --- Линия "Всего" только для судей
        if mode == "judges":
            total_item = QListWidgetItem()
            total_widget = ColorCheckItem("Показывать линию 'Всего'", (0, 0, 0))

            total_widget.color_label.hide()
            total_widget.checkbox.setChecked(True)
            total_widget.checkbox.stateChanged.connect(self.update_chart)

            total_item.setSizeHint(total_widget.sizeHint())
            target_list.addItem(total_item)
            target_list.setItemWidget(total_item, total_widget)

            self._total_item_widget = total_widget

        # --- Отметить / снять все (и для судей и для категорий)
        select_item = QListWidgetItem()
        select_widget = ColorCheckItem("Отметить / снять все", (0, 0, 0))

        select_widget.color_label.hide()
        select_widget.checkbox.stateChanged.connect(
            lambda state: self._toggle_all_generic(state, target_list)
        )

        select_item.setSizeHint(select_widget.sizeHint())
        target_list.addItem(select_item)
        target_list.setItemWidget(select_item, select_widget)

    def _reset_zoom(self):

        if not hasattr(self, "ax"):
            return

        week_indexes = self._get_filtered_weeks()
        if not week_indexes:
            return

        # --- Сброс X ---
        self.ax.set_xlim(0, len(week_indexes) - 1)

        # --- Сброс Y ---
        y_values = []

        for line in self.ax.get_lines():
            y_values.extend(line.get_ydata())

        if y_values:
            y_min = min(y_values)
            y_max = max(y_values)

            # небольшой отступ сверху
            padding = (y_max - y_min) * 0.05 if y_max != y_min else 1

            self.ax.set_ylim(y_min - padding, y_max + padding)

        self.canvas.draw_idle()

    def _on_press(self, event):

        if event.inaxes != self.ax:
            return

        # ЛКМ — панорамирование
        if event.button == 1:
            self._pan_start = (event.xdata, event.ydata)
            self._pan_xlim = self.ax.get_xlim()
            self._pan_ylim = self.ax.get_ylim()

        # ПКМ — сброс масштаба
        if event.button == 3:
            self._reset_zoom()

    def _on_release(self, event):
        self._pan_start = None
        self._pan_xlim = None

    def _on_pan_motion(self, event):

        if self._pan_start is None:
            return

        if event.inaxes != self.ax:
            return

        dx = event.xdata - self._pan_start[0]
        dy = event.ydata - self._pan_start[1]

        x_min, x_max = self._pan_xlim
        y_min, y_max = self._pan_ylim

        self.ax.set_xlim(x_min - dx, x_max - dx)
        self.ax.set_ylim(y_min - dy, y_max - dy)

        self.canvas.draw_idle()

    def _on_scroll(self, event):

        if not hasattr(self, "ax"):
            return

        if event.inaxes != self.ax:
            return

        base_scale = 1.2

        if event.button == "up":
            scale_factor = 1 / base_scale
        elif event.button == "down":
            scale_factor = base_scale
        else:
            return

        x_min, x_max = self.ax.get_xlim()
        y_min, y_max = self.ax.get_ylim()

        xdata = event.xdata
        ydata = event.ydata

        if xdata is None or ydata is None:
            return

        new_width = (x_max - x_min) * scale_factor
        new_height = (y_max - y_min) * scale_factor

        rel_x = (x_max - xdata) / (x_max - x_min)
        rel_y = (y_max - ydata) / (y_max - y_min)

        self.ax.set_xlim(
            xdata - new_width * (1 - rel_x),
            xdata + new_width * rel_x
        )

        self.ax.set_ylim(
            ydata - new_height * (1 - rel_y),
            ydata + new_height * rel_y
        )

        self.canvas.draw_idle()

    def _on_hover(self, event):

        if not hasattr(self, "ax"):
            return

        if event.inaxes != self.ax:
            if self._hover_annotation:
                self._hover_annotation.set_visible(False)
                self.canvas.draw_idle()
            return

        found = False

        for line in self.ax.get_lines():

            contains, info = line.contains(event)
            if not contains:
                continue

            ind = info["ind"][0]
            x = line.get_xdata()[ind]
            y = line.get_ydata()[ind]

            week_indexes = self._get_filtered_weeks()
            if not (0 <= ind < len(week_indexes)):
                return

            _, week_key = week_indexes[ind]

            label = line.get_label()

            # красиво отображаем название линии
            if label == "__total__":
                display_name = "Всего"
            else:
                display_name = label

            text = (
                f"{display_name}\n"
                f"{week_key}\n"
                f"Значение: {int(y)}"
            )

            if self._hover_annotation is None:
                self._hover_annotation = self.ax.annotate(
                    text,
                    xy=(x, y),
                    xytext=(15, 15),
                    textcoords="offset points",
                    bbox=dict(boxstyle="round", fc="white", ec="black"),
                    arrowprops=dict(arrowstyle="->")
                )
            else:
                self._hover_annotation.xy = (x, y)
                self._hover_annotation.set_text(text)
                self._hover_annotation.set_visible(True)

            found = True
            break

        if not found and self._hover_annotation:
            self._hover_annotation.set_visible(False)

        self.canvas.draw_idle()

    def _toggle_all_judges(self, state):

        checked = state == Qt.Checked

        for i in range(self.judges_list.count()):
            item = self.judges_list.item(i)
            widget = self.judges_list.itemWidget(item)

            text = widget.text_label.text()

            # игнорируем служебные элементы
            if text.startswith("Показывать") or text.startswith("Отметить"):
                continue

            widget.checkbox.blockSignals(True)
            widget.checkbox.setChecked(checked)
            widget.checkbox.blockSignals(False)

        self.update_chart()

    def _on_date_changed(self):
        self._user_range_selected = True
        self._fill_judges()
        self.update_chart()

    def set_data(self, raw_data, processor):
        self.raw_data = raw_data
        self.processor = processor

        all_weeks = sorted(
            raw_data.keys(),
            key=lambda w: datetime.strptime(w.split(" - ")[0], "%d.%m.%Y")
        )

        self.weeks = all_weeks

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

        if not self._user_range_selected:
            # только при первом запуске
            if len(self._week_dates) >= 20:
                first = self._week_dates[-20][0]
            else:
                first = self._week_dates[0][0]

            last = self._week_dates[-1][1]

            self.date_from.blockSignals(True)
            self.date_to.blockSignals(True)

            self.date_from.setDate(QDate(first.year, first.month, first.day))
            self.date_to.setDate(QDate(last.year, last.month, last.day))

            self.date_from.blockSignals(False)
            self.date_to.blockSignals(False)

        # print("_parse_week_dates - WEEKS:", len(self.weeks))
        # print("_parse_week_dates - FILTERED:", len(self._get_filtered_weeks()))

    def _fill_categories(self):
        # 🔥 сохраняем текущее состояние
        previous_state = {}
        for i in range(self.categories_list.count()):
            item = self.categories_list.item(i)
            widget = self.categories_list.itemWidget(item)
            previous_state[widget.text_label.text()] = widget.checkbox.isChecked()

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

            # 🔥 восстановление состояния
            if category in previous_state:
                widget.checkbox.setChecked(previous_state[category])
            else:
                widget.checkbox.setChecked(True)

            widget.checkbox.stateChanged.connect(self.update_chart)

        self._update_select_all_state()
        self._add_bottom_controls(self.categories_list, mode="categories")
        self.update_chart()

    def _fill_judges(self):

        # 🔥 сохраняем текущее состояние
        previous_state = {}

        for i in range(self.judges_list.count()):
            item = self.judges_list.item(i)
            widget = self.judges_list.itemWidget(item)
            previous_state[widget.text_label.text()] = widget.checkbox.isChecked()

        self.judges_list.clear()

        category = self.category_combo.currentText()
        judges = set()

        week_indexes = self._get_filtered_weeks()

        for _, week_key in week_indexes:
            week_data = self.raw_data.get(week_key, {})

            for judge, judge_data in week_data.items():
                cases = judge_data.get(category, [])
                if cases:
                    judges.add(judge)

        judges = sorted(judges)

        for judge in judges:
            color = self.judge_colors.get(judge, (0.5, 0.5, 0.5))

            item = QListWidgetItem()
            widget = ColorCheckItem(judge, color)

            item.setSizeHint(widget.sizeHint())
            self.judges_list.addItem(item)
            self.judges_list.setItemWidget(item, widget)

            # 🔥 восстановление состояния
            if judge in previous_state:
                widget.checkbox.setChecked(previous_state[judge])
            else:
                widget.checkbox.setChecked(True)

            widget.checkbox.stateChanged.connect(self.update_chart)

        self._add_bottom_controls(self.judges_list, mode="judges")
        self.update_chart()

    def _update_select_all_state(self):

        # определяем активный список
        if self.compare_mode.isChecked():
            target_list = self.categories_list
        else:
            target_list = self.judges_list

        total = 0
        checked = 0
        select_all_widget = None

        for i in range(target_list.count()):
            item = target_list.item(i)
            widget = target_list.itemWidget(item)

            text = widget.text_label.text()

            # находим служебную галку
            if text.startswith("Отметить"):
                select_all_widget = widget
                continue

            # пропускаем "Показывать Всего"
            if text.startswith("Показывать"):
                continue

            total += 1
            if widget.checkbox.isChecked():
                checked += 1

        if not select_all_widget:
            return

        select_all_widget.checkbox.blockSignals(True)

        if total == 0:
            select_all_widget.checkbox.setCheckState(Qt.Unchecked)
        elif checked == total:
            select_all_widget.checkbox.setCheckState(Qt.Checked)
        elif checked == 0:
            select_all_widget.checkbox.setCheckState(Qt.Unchecked)
        else:
            select_all_widget.checkbox.setCheckState(Qt.PartiallyChecked)

        select_all_widget.checkbox.blockSignals(False)

    # ---------------- FILTER ----------------

    def _get_filtered_weeks(self):
        date_from = self.date_from.date().toPyDate()
        date_to = self.date_to.date().toPyDate()

        # 🔥 защита от перевёрнутого диапазона
        if date_from > date_to:
            self.date_from.blockSignals(True)
            self.date_to.blockSignals(True)

            self.date_from.setDate(QDate(date_to.year, date_to.month, date_to.day))
            self.date_to.setDate(QDate(date_from.year, date_from.month, date_from.day))

            self.date_from.blockSignals(False)
            self.date_to.blockSignals(False)

        result = []

        for i, (start, end) in enumerate(self._week_dates):
            start_date = start.date()
            end_date = end.date()

            if start_date <= date_to and end_date >= date_from:
                result.append((i, self.weeks[i]))

        # print("_get_filtered_weeks - FROM:", date_from, "TO:", date_to)
        # print("_get_filtered_weeks - result:", result)

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

        # очищаем фигуру
        self.figure.clear()

        # создаём новый axes и сохраняем его
        self.ax = self.figure.add_subplot(111)

        # 🔥 ВАЖНО: сбрасываем hover-аннотацию
        self._hover_annotation = None

        week_indexes = self._get_filtered_weeks()
        if not week_indexes:
            self.canvas.draw()
            return

        # =========================
        # ОБЫЧНЫЙ РЕЖИМ (СУДЬИ)
        # =========================
        if not self.compare_mode.isChecked():

            category = self.category_combo.currentText()

            # ---- ВСЕ судьи (для totals)
            all_judges = set()
            for _, week_key in week_indexes:
                week_data = self.raw_data.get(week_key, {})
                all_judges.update(week_data.keys())

            all_judges = sorted(all_judges)

            # ---- выбранные судьи (для отображения)
            selected_judges = self._get_selected_judges()

            # строим серии
            full_series = self._build_series(category, all_judges, week_indexes)
            display_series = self._build_series(category, selected_judges, week_indexes)

            # ---- линии судей
            for judge, values in display_series.items():

                if not any(values):
                    continue

                self.ax.plot(
                    range(len(values)),
                    values,
                    marker="o",
                    label=judge,
                    color=self.judge_colors.get(judge, "gray"),
                    picker=6
                )

            # ---- линия "Всего" (НЕ зависит от галочек)
            if (
                    not self.compare_mode.isChecked()
                    and hasattr(self, "_total_item_widget")
                    and self._total_item_widget.checkbox.isChecked()
                    and full_series
            ):

                totals = [
                    sum(full_series[j][i] for j in full_series)
                    for i in range(len(week_indexes))
                ]

                self.ax.plot(
                    range(len(totals)),
                    totals,
                    linestyle="--",
                    color="black",
                    label="__total__",
                    picker=6   # ← ВАЖНО
                )

                # 🔥 если нет выбранных судей — вручную задать Y-лимиты
                if not selected_judges:
                    y_min = min(totals)
                    y_max = max(totals)

                    padding = (y_max - y_min) * 0.05 if y_max != y_min else 1

                    self.ax.set_ylim(y_min - padding, y_max + padding)

            self.ax.set_title(category)

        # =========================
        # СРАВНЕНИЕ КАТЕГОРИЙ
        # =========================
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

                for _, week_key in week_indexes:
                    total = 0
                    week_data = self.raw_data.get(week_key, {})

                    for judge_data in week_data.values():
                        cases = judge_data.get(category, [])
                        total += len(cases)

                    values.append(total)

                if not any(values):
                    continue

                self.ax.plot(
                    range(len(values)),
                    values,
                    marker="o",
                    label=category,
                    color=self.category_colors.get(category, "gray"),
                    picker=6
                )

            self.ax.set_title("Сравнение категорий")

        # =========================
        # ОБЩИЕ НАСТРОЙКИ
        # =========================

        self.ax.set_xticks(range(len(week_indexes)))
        self.ax.set_xticklabels(
            [w[-10:] for _, w in week_indexes],
            rotation=90
        )

        self.ax.grid(True)
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

        mouse_event = event.mouseevent
        line = event.artist

        ind = event.ind[0]
        week_indexes = self._get_filtered_weeks()

        if not (0 <= ind < len(week_indexes)):
            return

        _, week_key = week_indexes[ind]
        category = self.category_combo.currentText()

        ydata = line.get_ydata()
        clicked_value = int(ydata[ind])

        week_data = self.raw_data.get(week_key, {})
        label = line.get_label()

        # ==========================
        # КЛИК ПО "ВСЕГО"
        # ==========================
        if label == "__total__":

            judges_with_counts = []

            for judge, judge_data in week_data.items():
                cases = judge_data.get(category, [])
                count = len(cases)
                if count > 0:
                    judges_with_counts.append((judge, count))

            # 🔥 сортировка по убыванию количества дел
            judges_with_counts.sort(key=lambda x: x[1], reverse=True)

            # берём только имена судей
            matched_judges = [j[0] for j in judges_with_counts]

            data = {
                "week_key": week_key,
                "category": category,
                "judges": matched_judges,
                "value": clicked_value,
                "double_click": mouse_event.dblclick,
                "is_total": True
            }

            self.point_clicked.emit(data)
            return

        # ==========================
        # КЛИК ПО ЛИНИИ СУДЬИ
        # ==========================
        # 🔥 Ищем ВСЕХ судей с таким же значением
        matched_judges = []

        for judge, judge_data in week_data.items():
            cases = judge_data.get(category, [])
            if len(cases) == clicked_value:
                matched_judges.append(judge)

        data = {
            "week_key": week_key,
            "category": category,
            "judges": matched_judges,
            "value": clicked_value,
            "double_click": mouse_event.dblclick,
            "is_total": False
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

    def _toggle_compare_mode(self):
        if self.compare_mode.isChecked():
            self.category_combo.hide()
            self.judges_list.hide()
            self.categories_list.show()
        else:
            self.category_combo.show()
            self.judges_list.show()
            self.categories_list.hide()

        self.update_chart()