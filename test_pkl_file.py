import sys
import pickle
from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QTreeWidget,
    QTreeWidgetItem,
    QFileDialog,
)
from PyQt5.QtCore import Qt


def load_pkl(path):
    with open(path, "rb") as f:
        return pickle.load(f)


def add_item(parent, key, value):
    """
    Рекурсивно добавляет элементы в дерево
    """
    if isinstance(value, dict):
        item = QTreeWidgetItem(parent, [str(key), "dict"])
        for k, v in value.items():
            add_item(item, k, v)

    elif isinstance(value, (list, tuple)):
        item = QTreeWidgetItem(parent, [str(key), f"{type(value).__name__} ({len(value)})"])
        for i, v in enumerate(value):
            add_item(item, f"[{i}]", v)

    else:
        QTreeWidgetItem(parent, [str(key), repr(value)])


class PklTreeViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PKL Tree Viewer")
        self.resize(800, 600)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Key", "Value / Type"])
        self.tree.setColumnWidth(0, 300)

        self.setCentralWidget(self.tree)

        self.open_file()

    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Открыть PKL файл",
            "",
            "Pickle files (*.pkl)"
        )
        if not path:
            return

        data = load_pkl(path)
        self.tree.clear()

        root = QTreeWidgetItem(self.tree, ["<root>", type(data).__name__])
        self.tree.addTopLevelItem(root)

        if isinstance(data, dict):
            for k, v in data.items():
                add_item(root, k, v)
        else:
            add_item(root, "data", data)

        self.tree.expandToDepth(1)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    viewer = PklTreeViewer()
    viewer.show()
    sys.exit(app.exec_())
