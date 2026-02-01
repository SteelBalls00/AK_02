import sys
import os

from PyQt5.QtWidgets import QFrame
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QApplication, QToolButton, QMainWindow, QWidget,
    QVBoxLayout, QComboBox, QMessageBox, QTableView,
    QRadioButton, QGroupBox, QHBoxLayout, QPushButton,
    QLabel, QHeaderView, QTextEdit, QSplitter
)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QSize

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
        central = QWidget(self)
        main_layout = QVBoxLayout(central)

        # ================= Верхняя панель =================
        top_layout = QHBoxLayout()

        # --- Переключение недель ---
        self.prev_week_btn = QPushButton("◀")
        self.next_week_btn = QPushButton("▶")
        self.week_label = QLabel("")
        self.week_label.setAlignment(Qt.AlignCenter)

        self.prev_week_btn.clicked.connect(self.prev_week)
        self.next_week_btn.clicked.connect(self.next_week)

        week_box = QHBoxLayout()
        week_box.addWidget(self.prev_week_btn)
        week_box.addWidget(self.week_label)
        week_box.addWidget(self.next_week_btn)

        top_layout.addLayout(week_box)

        # --- Суд ---
        court_box = QVBoxLayout()
        self.court_combo = QComboBox()
        self.court_combo.currentTextChanged.connect(self.on_court_changed)

        court_box.addWidget(self.court_combo)

        top_layout.addLayout(court_box)

        # --- Специализация ---
        spec_group = QGroupBox("Специализация")
        spec_layout = QHBoxLayout(spec_group)

        self.spec_buttons = {}

        for spec in ["GPK", "KAS", "AP", "AP1", "U1", "M_U1"]:
            btn = QRadioButton(spec)
            btn.toggled.connect(self.on_context_changed)
            spec_layout.addWidget(btn)
            self.spec_buttons[spec] = btn

        self.spec_buttons["GPK"].setChecked(True)
        top_layout.addWidget(spec_group)

        # --- Инстанция ---
        inst_group = QGroupBox("Инстанция")
        inst_layout = QVBoxLayout(inst_group)

        self.instance_buttons = {}

        first_btn = QRadioButton("1 инстанция")
        first_btn.instance_value = "first"
        first_btn.toggled.connect(self.on_context_changed)

        appeal_btn = QRadioButton("Апелляция")
        appeal_btn.instance_value = "appeal"
        appeal_btn.toggled.connect(self.on_context_changed)

        self.instance_buttons["first"] = first_btn
        self.instance_buttons["appeal"] = appeal_btn

        inst_layout.addWidget(first_btn)
        inst_layout.addWidget(appeal_btn)

        self.instance_buttons["first"].setChecked(True)
        top_layout.addWidget(inst_group)

        # растяжка, чтобы элементы не слипались
        top_layout.addStretch()

        # --- Кнопка выгрузки в Word ---
        script_dir = os.path.dirname(os.path.abspath(__file__))
        word_icon_path = os.path.join(script_dir, "Word_png.png")

        self.word_export_btn = QPushButton()
        self.word_export_btn.setIcon(QIcon(word_icon_path))
        self.word_export_btn.setIconSize(QSize(86, 25))
        self.word_export_btn.clicked.connect(self.export_to_word)

        top_layout.addWidget(self.word_export_btn)

        # ================= Перетаскиватель =================
        self.splitter = QSplitter(Qt.Vertical)

        # ================= Таблица =================
        self.table_view = QTableView()
        self.model = TableModel()
        self.table_view.setModel(self.model)
        self.table_view.setSortingEnabled(True)

        header = self.table_view.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeToContents)
        header.setDefaultAlignment(Qt.AlignCenter)
        header = self.table_view.horizontalHeader()

        # не обрезать текст троеточием
        header.setTextElideMode(Qt.ElideNone)

        # центрирование
        header.setDefaultAlignment(Qt.AlignCenter)

        # ширина под содержимое
        header.setSectionResizeMode(QHeaderView.ResizeToContents)

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

        self.details_view.setPlaceholderText(
            "Выберите ячейку таблицы, чтобы увидеть детализацию"
        )

        selection_model = self.table_view.selectionModel()
        selection_model.selectionChanged.connect(self.on_table_selection_changed)

        # ================= Индикатор загрузки =================
        self.loading_label = QLabel("Загрузка данных...")
        self.loading_label.setAlignment(Qt.AlignCenter)
        self.loading_label.setVisible(False)

        # ================= Разделитель =================
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)

        # ================= Сборка =================
        self.splitter.addWidget(self.table_view)
        self.splitter.addWidget(self.details_view)
        self.splitter.setStretchFactor(0, 8)  # таблица
        self.splitter.setStretchFactor(1, 2)  # детализация

        main_layout.addLayout(top_layout)
        main_layout.addWidget(self.loading_label)
        main_layout.addWidget(separator)
        main_layout.addWidget(self.splitter)
        # main_layout.addWidget(self.table_view, stretch=6)
        # main_layout.addWidget(self.details_view, stretch=2)

        self.setCentralWidget(central)

    def _format_details_block(self, judge, column, details):
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
                v = v[6:]
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
        courts = self.bases_repo.get_courts()

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
        for spec, btn in self.spec_buttons.items():
            if btn.isChecked():
                self.specialization = spec
                break

        self.reload_current_court()

    def on_court_changed(self, court_name):
        pkl_files = self.bases_repo.get_pkl_files(court_name)

        has_appeal = any(
            info.instance == "appeal" and info.specialization == self.specialization
            for name, info in PKL_MAPPING.items()
            if name in pkl_files
        )

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
        self.loading_label.setVisible(True)
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

        worker.finished.connect(self.on_data_loaded)
        worker.error.connect(self.on_data_error)

        worker.start()

    def on_data_loaded(self, table_data):
        self.loading_label.setVisible(False)
        self.table_view.setEnabled(True)

        self.model.set_table_data(table_data)

        # сортировка по судье
        self.table_view.sortByColumn(0, Qt.AscendingOrder)

        self.week_label.setText(table_data.get("week", ""))
        self.loading_label.setVisible(False)
        self.table_view.setEnabled(True)

        self.active_workers.clear()

        self.current_week_key = table_data.get("week")

    def on_data_error(self, message, worker):
        QMessageBox.critical(self, "Ошибка загрузки", message)
        self.loading_label.setVisible(False)
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
            specialization=self.specialization,
            court=self.court_combo.currentText(),
            week=self.week_label.text()
        )

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.resize(1200, 800)
    window.showMaximized()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
