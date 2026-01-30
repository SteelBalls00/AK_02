from app.processors.base import BaseProcessor


class GPKAppealRegionalProcessor(BaseProcessor):
    """
    ГПК — апелляционная инстанция — областной суд
    """

    def __init__(self):
        self.columns = [
            "Судья",
            "Поступило за неделю",
            "Поступило с начала года",
            "Рассмотрено за неделю",
            "Рассмотрено с начала года",
            "Остаток",
            "Свыше 3 месяцев",
            "Свыше 6 месяцев",
            "Свыше 1 года",
        ]

        self.categories = [
            "Поступило за неделю",
            "Поступило с начала года",
            "Рассмотрено за неделю",
            "Рассмотрено с начала года",
            "Остаток",
            "Свыше трех месяцев",
            "Свыше шести месяцев",
            "Свыше одного года",
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
            v["Поступило за неделю"],
            v["Поступило с начала года"],
            v["Рассмотрено за неделю"],
            v["Рассмотрено с начала года"],
            v["Остаток"],
            v["Свыше трех месяцев"],
            v["Свыше шести месяцев"],
            v["Свыше одного года"],
        ]

    def _build_total_row(self, t):
        return [
            "Всего",
            t["Поступило за неделю"],
            t["Поступило с начала года"],
            t["Рассмотрено за неделю"],
            t["Рассмотрено с начала года"],
            t["Остаток"],
            t["Свыше трех месяцев"],
            t["Свыше шести месяцев"],
            t["Свыше одного года"],
        ]
