# app/processors/base.py

class BaseProcessor:
    """
    Базовый класс для всех процессоров статистики.
    """

    def build(self, raw_data, week_index):
        """
        raw_data: dict из .pkl
        week_index: индекс недели
        return: dict с columns, rows, total
        """
        raise NotImplementedError
