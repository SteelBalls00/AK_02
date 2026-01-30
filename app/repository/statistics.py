'''Repository — «откуда взять данные»

Задача:
открыть .pkl
ничего не анализировать
ничего не считать

На выходе:

raw_data        # словари из pkl
pkl_name        # result4_with_2.pkl

Repository:

не знает, что такое GPK
не знает, как строится таблица
он просто возвращает данные + контекст
'''

import os
import pickle
from app.constants.pkl_mapping import get_pkl_info
from app.domain.context import DataContext


class StatisticsRepository:

    def load(self, pkl_path):
        with open(pkl_path, "rb") as f:
            raw_data = pickle.load(f)

        pkl_name = os.path.basename(pkl_path)
        pkl_info = get_pkl_info(pkl_name)

        context = DataContext.from_pkl_info(pkl_info)

        return raw_data, context
