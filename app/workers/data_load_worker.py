from PyQt5.QtCore import QThread, pyqtSignal


class DataLoadWorker(QThread):
    finished = pyqtSignal(object)
    error = pyqtSignal(str)

    def __init__(self, processor, raw_data, week_index):
        super().__init__()
        self.processor = processor
        self.raw_data = raw_data
        self.week_index = week_index

    def run(self):
        try:
            table_data = self.processor.build(
                self.raw_data,
                self.week_index
            )
            self.finished.emit(table_data)
        except Exception as e:
            self.error.emit(str(e))

