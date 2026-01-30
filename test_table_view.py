import sys
from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QTableView
)

from app.ui.table_model import TableModel
from app.repository.statistics import StatisticsRepository
from app.factory.processor_factory import ProcessorFactory


def main():
    app = QApplication(sys.argv)

    repo = StatisticsRepository()
    raw_data, context = repo.load(
        r"\\192.168.0.200\Minato\3\Благовещенский городской суд\result4_with_2.pkl"
    )

    processor = ProcessorFactory.get(context)
    table_data = processor.build(raw_data, week_index=0)

    model = TableModel(table_data)

    view = QTableView()
    view.setModel(model)
    view.setSortingEnabled(True)
    view.resizeColumnsToContents()

    window = QMainWindow()
    window.setCentralWidget(view)
    window.resize(1200, 700)
    window.show()

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
