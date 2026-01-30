from app.processors.base import BaseProcessor


class GPKFirstRegionalProcessor(BaseProcessor):
    """
    ГПК — 1 инстанция — областной суд
    """

    def __init__(self):
        self.columns = [
            "Судья",
            "Рассмотрено за неделю",
            "Рассмотрено с начала года",
            "Принято за неделю",
            "Принято с начала года",
            "Остаток",
            "Свыше 1 года",
            "Свыше 2 лет",
        ]

        self.categories = [
            "Рассмотрено за неделю",
            "Рассмотрено с начала года",
            "Принято за неделю",
            "Принято с начала года",
            "Остаток",
            "Свыше одного года",
            "Свыше двух лет",
        ]

    def build(self, raw_data, week_index):
        weeks = list(raw_data.keys())
        week_key = weeks[week_index]
        week_data = raw_data[week_key]

        totals = self._init_counter()
        rows = []

        for judge, judge_data in week_data.items():
            row, totals = self._process_judge(judge, judge_data, totals)
            rows.append(row)

        return {
            "week": week_key,
            "columns": self.columns,
            "rows": rows,
            "total": self._build_total_row(totals),
        }

    def _init_counter(self):
        return {name: 0 for name in self.categories}

    def _process_judge(self, judge, judge_data, totals):
        values = self._init_counter()

        for category, cases in judge_data.items():
            if category in values:
                count = len(cases)
                values[category] = count
                totals[category] += count

        return self._format_row(judge, values), totals

    def _format_row(self, j, v):
        return [
            j,
            v["Рассмотрено за неделю"],
            v["Рассмотрено с начала года"],
            v["Принято за неделю"],
            v["Принято с начала года"],
            v["Остаток"],
            v["Свыше одного года"],
            v["Свыше двух лет"],
        ]

    def _build_total_row(self, t):
        return [
            "Всего",
            t["Рассмотрено за неделю"],
            t["Рассмотрено с начала года"],
            t["Принято за неделю"],
            t["Принято с начала года"],
            t["Остаток"],
            t["Свыше одного года"],
            t["Свыше двух лет"],
        ]
