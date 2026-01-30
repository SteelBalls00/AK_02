import os


class BasesRepository:
    def __init__(self, base_dir):
        self.base_dir = base_dir

    def get_courts(self):
        """
        Возвращает список названий судов (папок)
        """
        if not os.path.exists(self.base_dir):
            return []

        return sorted(
            name for name in os.listdir(self.base_dir)
            if os.path.isdir(os.path.join(self.base_dir, name))
        )

    def get_pkl_files(self, court_name):
        """
        Возвращает список pkl-файлов для выбранного суда
        """
        court_path = os.path.join(self.base_dir, court_name)
        if not os.path.exists(court_path):
            return []

        return [
            f for f in os.listdir(court_path)
            if f.endswith(".pkl")
        ]

    def get_pkl_path(self, court_name, pkl_name):
        return os.path.join(self.base_dir, court_name, pkl_name)
