import os
from typing import List, Set
from app.constants.pkl_mapping import PKL_MAPPING


class BasesRepository:
    def __init__(self, base_dir: str):
        self.base_dir = base_dir

    def get_courts(self):
        """Все папки судов (как есть на диске)"""
        if not os.path.exists(self.base_dir):
            return []

        return sorted([
            name for name in os.listdir(self.base_dir)
            if os.path.isdir(os.path.join(self.base_dir, name))
        ])

    def get_pkl_files(self, court_name: str) -> List[str]:
        """Список .pkl файлов для суда"""
        court_dir = os.path.join(self.base_dir, court_name)
        if not os.path.exists(court_dir):
            return []

        return [
            f for f in os.listdir(court_dir)
            if f.endswith(".pkl")
        ]

    # ===== НОВОЕ =====

    def get_courts_with_any_pkls(self) -> List[str]:
        """
        Только те суды, у которых реально есть хотя бы один .pkl
        """
        result = []

        for court in self.get_courts():
            if self.get_pkl_files(court):
                result.append(court)

        return sorted(result)

    def get_available_specializations(self, court_name, instance):
        """
        Возвращает доступные специализации
        для конкретного суда и инстанции
        """

        pkl_files = self.get_pkl_files(court_name)

        specs = {
            info.specialization
            for name, info in PKL_MAPPING.items()
            if (
                    name in pkl_files
                    and info.instance == instance
            )
        }

        return specs

    def get_available_instances(self, court_name: str, specialization: str) -> Set[str]:
        """
        Возвращает доступные инстанции (first / appeal)
        для суда + специализации
        """
        pkl_files = self.get_pkl_files(court_name)

        instances = {
            info.instance
            for name, info in PKL_MAPPING.items()
            if (
                name in pkl_files
                and info.specialization == specialization
            )
        }

        return instances

    def get_pkl_path(self, court_name: str, pkl_name: str) -> str:
        return os.path.join(self.base_dir, court_name, pkl_name)
