from PyQt5.QtCore import QThread, pyqtSignal


class DataLoadWorker(QThread):
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, processor_factory, raw_data, context, week_index):
        super().__init__()
        self.processor_factory = processor_factory
        self.raw_data = raw_data
        self.context = context
        self.week_index = week_index

    def run(self):
        try:
            processor = self.processor_factory.get(self.context)
            table_data = processor.build(self.raw_data, self.week_index)
            self.finished.emit(table_data)
        except Exception as e:
            self.error.emit(str(e))
