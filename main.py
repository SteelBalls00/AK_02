# pyinstaller --onedir --noconsole --hidden-import=openpyxl --name="AK_v1.5" main.py

'''
- путь к базам в файле настроек
- графики и детализация к ним
- столбцы для бездвижа и возвратов
- закрепить первый столбец с судьями,в случае ширины таблицы за пределы экрана

поправить:
- бокс с выбором суда иногда появляется пустой при наличии 1 суда

глобальные правки:
- поправить или сделать новый апдейт
'''

import sys
import os
import re
from docx import Document
from datetime import datetime, date
from openpyxl import Workbook


from PyQt5.QtWidgets import QFrame, QToolButton
from PyQt5.QtCore import Qt, QDate, QEasingCurve
from PyQt5.QtWidgets import (
    QApplication, QMenu, QMainWindow, QWidget,
    QVBoxLayout, QComboBox, QMessageBox, QTableView,
    QRadioButton, QGroupBox, QHBoxLayout, QPushButton,
    QLabel, QHeaderView, QTextEdit, QSplitter,
    QCalendarWidget, QDialog,
)
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtCore import QSize, QPropertyAnimation
from PyQt5.QtWidgets import QGraphicsOpacityEffect


from app.constants.pkl_mapping import PKL_MAPPING
from app.repository.bases_repository import BasesRepository
from app.repository.statistics import StatisticsRepository
from app.factory.processor_factory import ProcessorFactory
from app.domain.pkl_selector import select_pkl_for_context
from app.ui.table_model import TableModel
from app.workers.data_load_worker import DataLoadWorker
from app.export.word_exporter import export_model_to_word


BASE_DIR = os.path.join(os.path.dirname(__file__), "bases")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Анализ судебной статистики")

        self.bases_repo = BasesRepository(BASE_DIR)
        self.stats_repo = StatisticsRepository()

        # ====== СОСТОЯНИЕ (ДО UI!) ======
        self.specialization = "GPK"
        self.instance = "first"

        self.current_pkl_path = None
        self.current_raw_data = None
        self.current_context = None

        self.week_index = 0
        self.max_week_index = 0
        self.current_week_key = None

        self.active_workers = []

        # ====== UI ======
        self._init_ui()
        self._load_courts()

    def _init_ui(self):
        self._ui_ready = False

        central = QWidget(self)
        main_layout = QVBoxLayout(central)

        # ================= Верхняя панель =================
        header_widget = QWidget()
        header_widget.setObjectName("panel")
        top_layout = QHBoxLayout(header_widget)
        top_layout.setContentsMargins(8, 8, 8, 8)

        # --- Переключение недель ---
        self.prev_week_btn = QPushButton("◀")
        self.next_week_btn = QPushButton("▶")
        self.week_label = QLabel("")
        self.week_label.setAlignment(Qt.AlignCenter)
        self.week_label.setCursor(Qt.PointingHandCursor)
        self.week_label.mousePressEvent = self.on_week_label_clicked

        self.week_label.setProperty("role", "week-label")
        self.prev_week_btn.setProperty("role", "week-nav")
        self.next_week_btn.setProperty("role", "week-nav")

        self.prev_week_btn.clicked.connect(self.prev_week)
        self.next_week_btn.clicked.connect(self.next_week)

        for btn in (self.prev_week_btn, self.next_week_btn):
            btn.setFixedSize(68, 48)

        week_box = QHBoxLayout()
        week_box.addWidget(self.prev_week_btn)
        week_box.addWidget(self.week_label)
        week_box.addWidget(self.next_week_btn)

        top_layout.addLayout(week_box)

        # --- Суд ---
        court_group = QGroupBox("Суд")
        court_layout = QVBoxLayout(court_group)

        self.court_combo = QComboBox()
        self.court_combo.currentTextChanged.connect(self.on_court_changed)

        court_layout.addWidget(self.court_combo)

        top_layout.addWidget(court_group)

        # --- Специализация ---
        spec_group = QGroupBox("Специализация")
        self.spec_layout = QHBoxLayout(spec_group)

        self.spec_buttons = {}

        specs = {
            "GPK": "ГПК",
            "KAS": "КАС",
            "AP": "АП",
            "AP1": "АП1",
            "U1": "УГ",
            "M_U1": "М.Уг",
            "M_AOS": "М.",
        }

        for code, label in specs.items():
            rb = QRadioButton(label)
            rb.setProperty("spec", code)
            rb.toggled.connect(self.on_context_changed)
            self.spec_buttons[code] = rb
            self.spec_layout.addWidget(rb)

        self.spec_buttons["GPK"].setChecked(True)
        top_layout.addWidget(spec_group)

        # --- Инстанция ---
        inst_group = QGroupBox("Инстанция")
        inst_layout = QVBoxLayout(inst_group)

        self.instance_buttons = {}

        first_btn = QRadioButton("1 инстанция")
        first_btn.instance_value = "first"

        appeal_btn = QRadioButton("Апелляция")
        appeal_btn.instance_value = "appeal"

        self.instance_buttons["first"] = first_btn
        self.instance_buttons["appeal"] = appeal_btn

        appeal_btn.toggled.connect(self.on_context_changed)
        first_btn.toggled.connect(self.on_context_changed)

        inst_layout.addWidget(first_btn)
        inst_layout.addWidget(appeal_btn)

        self.instance_buttons["first"].setChecked(True)
        top_layout.addWidget(inst_group)

        # --- Кнопка выгрузки в Word ---
        script_dir = os.path.dirname(os.path.abspath(__file__))
        word_icon_path = os.path.join(script_dir, "Word_png.png")

        self.word_export_btn = QPushButton()
        self.word_export_btn.setIcon(QIcon(word_icon_path))
        self.word_export_btn.setIconSize(QSize(86, 25))
        self.word_export_btn.clicked.connect(self.export_to_word)
        self.word_export_btn.setObjectName("export_to_word")

        top_layout.addWidget(self.word_export_btn)

        # растяжка, чтобы элементы не слипались
        top_layout.addStretch()

        # ================= переключатель темы =================
        self.theme_toggle_btn = QToolButton()
        self.theme_toggle_btn.setText("🌙 Тёмная тема")
        self.theme_toggle_btn.setCheckable(True)
        self.theme_toggle_btn.setChecked(False)  # стартуем с тёмной

        self.theme_toggle_btn.clicked.connect(self.toggle_theme)

        top_layout.addWidget(self.theme_toggle_btn)

        # ================= Черточка перед таблицей =================
        self.splitter = QSplitter(Qt.Vertical)

        # ================= Таблица =================
        self.table_view = QTableView()
        self.model = TableModel()
        self.table_view.setModel(self.model)
        self.table_view.setSortingEnabled(True)

        self.table_view.setAlternatingRowColors(True)
        self.table_view.setShowGrid(True)
        self.table_view.verticalHeader().setVisible(False)
        self.table_view.horizontalHeader().setStretchLastSection(True)

        '''
        Настройка высоты строк
        20–22 — очень компактно
        24 — комфортно
        28 — «воздушно»
        '''

        vertical_header = self.table_view.verticalHeader()
        vertical_header.setDefaultSectionSize(24)  # Настройка высоты строк

        self.table_opacity = QGraphicsOpacityEffect(self.table_view.viewport())
        self.table_view.viewport().setGraphicsEffect(self.table_opacity)
        self.table_opacity.setOpacity(1.0)  # ВАЖНО

        self.fade_anim = QPropertyAnimation(self.table_opacity, b"opacity")
        self.fade_anim.setDuration(150)
        self.fade_anim.setEasingCurve(QEasingCurve.InOutQuad)

        header = self.table_view.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeToContents)
        header.setDefaultAlignment(Qt.AlignCenter)

        # не обрезать текст троеточием
        header.setTextElideMode(Qt.ElideNone)  # Управляет обрезкой текста, если он не влезает

        # центрирование
        header.setDefaultAlignment(Qt.AlignCenter)  # Центрирует текст внутри ячеек заголовка

        # ширина под содержимое
        header.setSectionResizeMode(QHeaderView.ResizeToContents)  # ширина столбца = ширина самого широкого содержимого

        # даём место для многострочных заголовков
        header.setFixedHeight(70)

        self.table_view.setStyleSheet("""
        QHeaderView::section {
            padding: 6px;
            font-weight: bold;
        }
        """)

        # --- Область детализации ---
        self.details_view = QTextEdit()

        self.details_view.setReadOnly(True)
        self.details_view.setMinimumHeight(30)
        self.details_view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.details_view.customContextMenuRequested.connect(
            self.show_details_context_menu
        )
        self.details_view.setLineWrapMode(QTextEdit.WidgetWidth)
        self.details_view.setFont(QFont("Consolas", 10))

        self.details_view.setPlaceholderText(
            "Выберите ячейку таблицы, чтобы увидеть детализацию"
        )

        selection_model = self.table_view.selectionModel()
        selection_model.selectionChanged.connect(self.on_table_selection_changed)

        # ================= Разделитель =================
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)

        # ================= Сборка =================
        self.splitter.addWidget(self.table_view)
        self.splitter.addWidget(self.details_view)
        self.splitter.setStretchFactor(0, 8)  # таблица
        self.splitter.setStretchFactor(1, 4)  # детализация

        main_layout.addWidget(header_widget)
        main_layout.addWidget(separator)
        main_layout.addWidget(self.splitter)

        self.setCentralWidget(central)

        self._ui_ready = True

    def set_radio_visible(self, btn, visible: bool):
        if not visible and btn.isChecked():
            btn.setAutoExclusive(False)
            btn.setChecked(False)
            btn.setAutoExclusive(True)

        btn.setVisible(visible)

    def update_instance_buttons(self, court_name: str):
        instances = self.bases_repo.get_available_instances(
            court_name,
            self.specialization
        )

        for inst, btn in self.instance_buttons.items():
            btn.setEnabled(inst in instances)

        # защита от невалидного состояния
        if self.instance not in instances and instances:
            self.instance = next(iter(instances))
            self.instance_buttons[self.instance].setChecked(True)

    def update_specialization_buttons(self, court_name: str):
        available_specs = self.bases_repo.get_available_specializations(court_name, self.instance)

        for spec, btn in self.spec_buttons.items():
            self.set_radio_visible(btn, spec in available_specs)

        # гарантируем выбранную специализацию
        if self.specialization not in available_specs and available_specs:
            new_spec = next(iter(available_specs))
            self.spec_buttons[new_spec].setChecked(True)
            self.specialization = new_spec

    def animate_table_update(self, update_callback):
        """
        Полностью безопасное обновление таблицы:
        - без мигания
        - без микро-дёрганий
        - без призраков старых данных
        """

        # если анимация уже идёт — остановить
        if self.fade_anim.state() == QPropertyAnimation.Running:
            self.fade_anim.stop()

        # 1. МГНОВЕННО скрываем содержимое таблицы
        self.table_opacity.setOpacity(0.0)

        # 2. Полностью блокируем перерисовку
        self.table_view.setUpdatesEnabled(False)
        # 3. Применяем данные
        update_callback()
        # 4. Разрешаем перерисовку
        self.table_view.setUpdatesEnabled(True)

        # 5. Плавно показываем новую таблицу
        self.fade_anim.setStartValue(0.0)
        self.fade_anim.setEndValue(1.0)
        self.fade_anim.start()

    def toggle_theme(self, checked: bool):
        '''Переключение цвета темы'''
        app = QApplication.instance()

        if checked:
            app.setStyleSheet(DARK_STYLE)
            self.theme_toggle_btn.setText("🌞 Светлая тема")
        else:
            app.setStyleSheet(LIGHT_STYLE)
            self.theme_toggle_btn.setText("🌙 Тёмная тема")

    def select_week_by_date(self, selected_date: date):
        """
        Выбирает неделю, в которую попадает дата.
        Если такой нет — выбирает ближайшую.
        """
        weeks = list(self.current_raw_data.keys())

        parsed_weeks = []

        for idx, week_str in enumerate(weeks):
            try:
                start_str, end_str = week_str.split(" - ")
                start = datetime.strptime(start_str, "%d.%m.%Y").date()
                end = datetime.strptime(end_str, "%d.%m.%Y").date()
                parsed_weeks.append((idx, start, end))
            except Exception:
                continue

        if not parsed_weeks:
            return

        # 1️⃣ Пытаемся найти точное попадание
        for idx, start, end in parsed_weeks:
            if start <= selected_date <= end:
                self.week_index = idx
                self.reload_current_court()
                return

        # 2️⃣ Ищем ближайшую неделю
        def distance(week):
            _, start, end = week
            if selected_date < start:
                return (start - selected_date).days
            if selected_date > end:
                return (selected_date - end).days
            return 0

        closest = min(parsed_weeks, key=distance)
        self.week_index = closest[0]
        self.reload_current_court()

    def on_calendar_confirmed(self, calendar: QCalendarWidget, dialog: QDialog):
        qdate = calendar.selectedDate()
        selected_date = date(qdate.year(), qdate.month(), qdate.day())

        self.select_week_by_date(selected_date)

        dialog.accept()

    def on_week_label_clicked(self, event):
        if not self.current_raw_data:
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Выбор даты")
        dialog.setModal(True)

        layout = QVBoxLayout(dialog)

        calendar = QCalendarWidget()
        calendar.setGridVisible(True)
        calendar.setSelectedDate(QDate.currentDate())

        layout.addWidget(calendar)

        btn_ok = QPushButton("Выбрать")
        layout.addWidget(btn_ok)

        btn_ok.clicked.connect(lambda: self.on_calendar_confirmed(calendar, dialog))

        dialog.resize(300, 250)
        dialog.exec_()

    def parse_details_blocks(self):
        """
        Разбирает текст детализации на блоки:
        [
            {
                "header": ["Судья: ...", "Показатель: ..."],
                "items": ["2-2735/2025, ...", ...]
            },
            ...
        ]
        """
        text = self.details_view.toPlainText()
        lines = [l.rstrip() for l in text.splitlines()]

        blocks = []
        current_header = []
        current_items = []

        for line in lines:
            if not line:
                continue

            if line.startswith("Судья:") or line.startswith("Показатель:"):
                if current_items:
                    blocks.append({
                        "header": current_header,
                        "items": current_items
                    })
                    current_header = []
                    current_items = []

                current_header.append(line)

            elif line.strip().startswith("•"):
                current_items.append(line.replace("• ", "").strip())

        if current_items:
            blocks.append({
                "header": current_header,
                "items": current_items
            })

        return blocks

    def export_details_to_excel(self, only_numbers: bool):
        blocks = self.parse_details_blocks()

        wb = Workbook()
        ws = wb.active
        ws.title = "Детализация"

        row = 1

        for block in blocks:
            for header_line in block["header"]:
                ws.cell(row=row, column=1, value=header_line)
                row += 1

            for item in block["items"]:
                if only_numbers:
                    item = self.extract_case_number(item)
                ws.cell(row=row, column=1, value=item)
                row += 1

            row += 2  # пустая строка между блоками

        filename = f"details_{datetime.now():%d.%m.%Y.%H.%M.%S}.xlsx"
        wb.save(filename)
        os.startfile(filename)

    def export_details_to_word(self, only_numbers: bool):
        blocks = self.parse_details_blocks()

        document = Document()
        document.add_heading("Детализация", level=1)

        for block in blocks:
            # Заголовок блока
            for header_line in block["header"]:
                document.add_paragraph(header_line)

            document.add_paragraph("")  # отступ

            # Содержимое
            for item in block["items"]:
                if only_numbers:
                    item = self.extract_case_number(item)
                document.add_paragraph(item)

            document.add_page_break()

        filename = f"details_{datetime.now():%d.%m.%Y.%H.%M.%S}.docx"
        document.save(filename)
        os.startfile(filename)

    def get_details_lines(self):
        """
        Возвращает список строк детализации (без пустых)
        """
        text = self.details_view.toPlainText()
        return [line.strip() for line in text.splitlines() if line.strip()]

    def extract_case_number(self, line: str) -> str:
        """
        Извлекает номер дела до первой запятой
        """
        if "," in line:
            return line.split(",", 1)[0].strip().replace('• ', '')
        return line.strip()

    def copy_details_to_clipboard(self):
        blocks = self.parse_details_blocks()

        lines = []
        for block in blocks:
            lines.extend(block["header"])
            lines.extend(block["items"])
            lines.append("")

        QApplication.clipboard().setText("\n".join(lines))

    def show_details_context_menu(self, pos):
        menu = QMenu(self)

        # --- Копировать ---
        copy_action = menu.addAction("Скопировать в буфер обмена")
        copy_action.triggered.connect(self.copy_details_to_clipboard)

        menu.addSeparator()

        # --- Word ---
        word_menu = menu.addMenu("Передать в Word")
        word_only_numbers = word_menu.addAction("Только номера дел")
        word_full = word_menu.addAction("Номера дел со всей информацией")

        word_only_numbers.triggered.connect(
            lambda: self.export_details_to_word(only_numbers=True)
        )
        word_full.triggered.connect(
            lambda: self.export_details_to_word(only_numbers=False)
        )

        # --- Excel ---
        excel_menu = menu.addMenu("Передать в Excel")
        excel_only_numbers = excel_menu.addAction("Только номера дел")
        excel_full = excel_menu.addAction("Номера дел со всей информацией")

        excel_only_numbers.triggered.connect(
            lambda: self.export_details_to_excel(only_numbers=True)
        )
        excel_full.triggered.connect(
            lambda: self.export_details_to_excel(only_numbers=False)
        )

        menu.exec_(self.details_view.mapToGlobal(pos))



    def _format_details_block(self, judge, column, details):
        _PREFIX_RE = re.compile(r"\d\.\d{3}-")

        def normalize_case_line(raw: str) -> str:
            """
            Удаляет ТОЛЬКО префикс вида '2.123-' (цифра + точка + 3 цифры + дефис).
            Если такого шаблона нет — строка возвращается без изменений.
            """
            return _PREFIX_RE.sub("", raw, count=1)

        column = column.replace('\n', ' ')
        lines = [
            f"Судья: {judge}",
            f"Показатель: {column}",
        ]

        if not details:
            lines.append("Детализация отсутствует")
            return "\n".join(lines)

        # lines.append("Исходные значения:")

        for title, values in details:
            total = len(values) if values else 0
            lines.append(f"{title}: {total}")

            for v in values:
                v = normalize_case_line(v)
                lines.append(f"  • {v}")

        return "\n".join(lines)

    def on_table_selection_changed(self, selected, deselected):
        if not self.current_context:
            return

        indexes = self.table_view.selectionModel().selectedIndexes()
        if not indexes:
            self.details_view.clear()
            return

        blocks = []

        for index in indexes:
            row = index.row()
            col = index.column()

            # имя судьи — всегда первый столбец
            judge = self.model.data(self.model.index(row, 0))

            column_name = self.model.headerData(col, Qt.Horizontal)

            details = self.current_processor.get_cell_details(
                judge=judge,
                column=column_name,
                week_index=self.week_index,
            )

            blocks.append(self._format_details_block(
                judge, column_name, details
            ))

        self.details_view.setPlainText("\n\n".join(blocks))

    def _load_courts(self):
        courts = self.bases_repo.get_courts_with_any_pkls()

        self.court_combo.clear()
        self.court_combo.addItems(courts)

        # --- если суд только один ---
        if len(courts) <= 1:
            self.court_combo.hide()
        else:
            self.court_combo.show()

        # автоматически выбираем первый (или единственный)
        if courts:
            self.court_combo.setCurrentIndex(0)

    def reload_current_court(self):
        if not hasattr(self, "court_combo"):
            return

        court = self.court_combo.currentText()
        if court:
            self.on_court_changed(court)

    def on_context_changed(self):
        if not getattr(self, "_ui_ready", False):
            return

        # specialization
        for spec, btn in self.spec_buttons.items():
            if btn.isChecked():
                self.specialization = spec
                break

        # instance
        for inst, btn in self.instance_buttons.items():
            if btn.isChecked():
                self.instance = inst
                break

        self.reload_current_court()

    def on_court_changed(self, court_name):
        # Получаем доступные инстанции для суда
        available_instances = self.bases_repo.get_available_instances(court_name, self.specialization)

        # 🔑 Если текущая инстанция недоступна — переключаемся
        if self.instance not in available_instances:
            self.instance = "first"
            self.instance_buttons["first"].setChecked(True)

        # 1️⃣ Обновляем доступные specialization
        self.update_specialization_buttons(court_name)

        # 2️⃣ Обновляем доступные инстанции
        self.update_instance_buttons(court_name)

        pkl_files = self.bases_repo.get_pkl_files(court_name)

        # есть ли для выбранной специализации апел. база
        has_appeal = any(
            info.instance == "appeal" and info.specialization == self.specialization
            for name, info in PKL_MAPPING.items()
            if name in pkl_files
        )

        # и если есть, то ставим активной
        self.instance_buttons["appeal"].setEnabled(has_appeal)

        if not has_appeal and self.instance == "appeal":
            self.instance_buttons["first"].setChecked(True)

        pkl_name = select_pkl_for_context(
            pkl_files,
            specialization=self.specialization,
            instance=self.instance
        )

        if not pkl_name:
            QMessageBox.warning(
                self,
                "Нет данных",
                "Для выбранного суда нет подходящей базы"
            )
            self.model.set_table_data({})
            return

        pkl_path = self.bases_repo.get_pkl_path(court_name, pkl_name)

        # получаем количество недель
        # если путь тот же — просто обновляем таблицу
        if self.current_pkl_path == pkl_path and self.current_raw_data is not None:
            self.load_table_async()
            return

        # загружаем pkl ОДИН раз
        raw_data, context = self.stats_repo.load(pkl_path)

        self.current_raw_data = raw_data
        self.current_context = context
        self.current_pkl_path = pkl_path

        weeks = list(raw_data.keys())
        self.max_week_index = max(0, len(weeks) - 1)

        # --- Пытаемся сохранить текущую неделю ---
        if self.current_week_key in weeks:
            self.week_index = weeks.index(self.current_week_key)
        else:
            self.week_index = self.max_week_index

        self.load_table_async()

        # если вышли за границы — корректируем
        if self.week_index > self.max_week_index:
            self.week_index = self.max_week_index

        self.table_view.resizeColumnsToContents()

    def load_table_async(self):
        self.table_view.setEnabled(False)

        # 1. Получаем процессор из фабрики
        processor = ProcessorFactory.get(self.current_context)

        # 2. Сохраняем его для детализации
        self.current_processor = processor

        # 3. Запускаем воркер
        worker = DataLoadWorker(
            processor=processor,
            raw_data=self.current_raw_data,
            week_index=self.week_index
        )

        self.active_workers.append(worker)

        worker.finished.connect(
            lambda table_data, w=worker: self.on_data_loaded(table_data, w)
        )
        worker.error.connect(self.on_data_error)

        worker.start()

    def on_data_loaded(self, table_data, worker):
        def apply():
            self.model.set_table_data(table_data)

            # сортировка по судье
            self.table_view.sortByColumn(0, Qt.AscendingOrder)

            # неделя
            self.week_label.setText(table_data.get("week", ""))

            # UI
            self.table_view.setEnabled(True)

            self.current_week_key = table_data.get("week")

        self.animate_table_update(apply)

        if worker in self.active_workers:
            self.active_workers.remove(worker)

    def on_data_error(self, message, worker):
        QMessageBox.critical(self, "Ошибка загрузки", message)
        self.table_view.setEnabled(True)

        if worker in self.active_workers:
            self.active_workers.remove(worker)

    def prev_week(self):
        if self.week_index > 0:
            self.week_index -= 1
            self.reload_current_court()

    def next_week(self):
        if self.week_index < self.max_week_index:
            self.week_index += 1
            self.reload_current_court()

    def export_to_word(self):
        if self.model.rowCount() == 0:
            QMessageBox.information(self, "Нет данных", "Таблица пуста")
            return

        export_model_to_word(
            model=self.model,
            processor=self.current_processor,
            court=self.court_combo.currentText(),
            week=self.week_label.text()
        )

LIGHT_STYLE = """
QWidget {
    font-family: "Segoe UI";
    color: #2b2b2b;
}

/* --- Таблица --- */
QTableView {
    background-color: #ffffff;
    gridline-color: #dcdcdc;
    selection-background-color: #e6f0fa;
    selection-color: #000000;
    alternate-background-color: #fafafa;
}

QTableView::item:selected {
    background-color: #cfe3f6;
}

/* --- Кнопки --- */
QPushButton {
    background-color: #3a6ea5;
    color: white;
    border: none;
    padding: 6px 12px;
    border-radius: 4px;
}

QPushButton:hover {
    background-color: #4a82c0;
}

QPushButton:pressed {
    background-color: #2f5d8a;
}

QPushButton#export_to_word {
    background-color: transparent;
    min-width: 48px;
    min-height: 48px;
    padding: 0px;
}

QPushButton#export_to_word:hover {
    background-color: #5a96d5;
}

QPushButton[role="week-nav"] {
    background-color: #4a86c5;
    border: none;
    font-weight: bold;
    padding: 4px 8px;
    font-size: 20pt;
    font-weight: bold;
}

QPushButton[role="week-nav"]:hover {
    background-color: #4a86c5;
    font-size: 20pt;
    font-weight: bold;
}

QPushButton[role="week-nav"]:pressed {
    background-color: rgba(0, 0, 0, 0.15);
}

/* ================ RadioButton ================ */
/* --- Radio / Check --- */
QRadioButton, QCheckBox {
    spacing: 6px;
    font-weight: bold;
}
QRadioButton[spec="GPK"] { font-weight: bold; }

QRadioButton::indicator {
    width: 14px;
    height: 14px;
}

QRadioButton::indicator:disabled {
    background-color: #c0c0c0;
    width: 14px;
    height: 14px;
    border-radius: 7px;
}

/* ================== LABEL ================== */
QLabel {
    color: black;
}

QLabel[role="week-label"] {
    font-size: 20pt;
    font-weight: bold;
}

/* --- ComboBox --- */
QComboBox {
    background-color: #ffffff;
    border: 1px solid #cfcfcf;
    padding: 4px;
    border-radius: 4px;
}

/* --- Детализация --- */
QTextEdit {
    background-color: #fcfcfc;
    border: 1px solid #cfcfcf;
    border-radius: 4px;
    padding: 6px;
}

/* --- ToolButton (если появятся) --- */
QToolButton {
    background-color: transparent;
    border: none;
    padding: 4px;
}

QToolButton:hover {
    background-color: #e6f0fa;
}
"""

DARK_STYLE = """
/* ================== БАЗА ================== */
QWidget {
    background-color: #2b2b2b;
    color: #e6e6e6;
    font-family: "Segoe UI";
}

/* ================== ПАНЕЛИ ================== */
QFrame, QWidget#panel {
    background-color: #313335;
    border: 1px solid #444444;
    border-radius: 4px;
}

/* ================== LABEL ================== */
QLabel {
    color: #e6e6e6;
}

QLabel[role="week-label"] {
    font-size: 20pt;
    font-weight: bold;
}

/* ================== КНОПКИ ================== */
QPushButton {
    background-color: #4a86c5;
    color: #ffffff;
    border: none;
    padding: 6px 12px;
    border-radius: 4px;
}

QPushButton:hover {
    background-color: #5a96d5;
    font-size: 20pt;
    font-weight: bold;
}

QPushButton:pressed {
    background-color: #3a6ea5;
}

QPushButton[role="week-nav"] {
    background-color: #4a86c5;
    border: none;
    color: #a6c8ff;
    font-size: 20pt;
    font-weight: bold;
}

QPushButton[role="week-nav"]:hover {
    background-color: #4a86c5;
}

QPushButton[role="week-nav"]:pressed {
    background-color: rgba(255, 255, 255, 0.15);
}


/* ================== TOOL BUTTON ================== */
QToolButton {
    background-color: transparent;
    border: none;
    padding: 4px;
}

QToolButton:hover {
    background-color: #3d5a73;
}

/* ================== COMBOBOX ================== */
QComboBox {
    background-color: #2f3133;
    border: 1px solid #555555;
    padding: 4px;
    border-radius: 4px;
}

QComboBox QAbstractItemView {
    background-color: #2f3133;
    selection-background-color: #3d5a73;
}

/* ================== RADIO / CHECK ================== */
QRadioButton, QCheckBox {
    spacing: 6px;
    font-weight: bold;
}
QRadioButton:disabled {
    color: #777777;
}

QRadioButton::indicator:disabled {
    background-color: #555555;
    border: 1px solid #444444;
    width: 14px;
    height: 14px;
    border-radius: 7px;
}

/* ================== ТАБЛИЦА ================== */
QTableView {
    background-color: #2f3133;
    gridline-color: #444444;
    selection-background-color: #3d5a73;
    selection-color: #ffffff;
    alternate-background-color: #2a2c2e;
}

QTableView::item {
    padding: 4px;
}

QTableView::item:selected {
    background-color: #3d5a73;
}

/* ================== ЗАГОЛОВКИ ТАБЛИЦЫ ================== */
QHeaderView::section {
    background-color: #3a3d41;
    border: 1px solid #444444;
    padding: 6px;
    font-weight: bold;
}

/* ================== SCROLLBAR ================== */
QScrollBar:vertical {
    background: #2b2b2b;
    width: 12px;
}

QScrollBar::handle:vertical {
    background: #555555;
    min-height: 20px;
    border-radius: 6px;
}

QScrollBar::handle:vertical:hover {
    background: #666666;
}

/* ================== TEXT EDIT (детализация) ================== */
QTextEdit {
    background-color: #2f3133;
    border: 1px solid #444444;
    border-radius: 4px;
    padding: 6px;
}

/* ================== SPLITTER ================== */
QSplitter::handle {
    background-color: #444444;
}
"""



def main():
    app = QApplication(sys.argv)
    # app.setStyle("macOS")  # очень важно
    app.setStyleSheet(LIGHT_STYLE)
    window = MainWindow()
    window.resize(1200, 800)
    window.showMaximized()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
