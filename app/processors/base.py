class BaseProcessor:
    COLUMN_TO_CATEGORY = {}
    COLUMN_TO_INCLUDED_CATEGORY = {}

    def build(self, raw_data, week_index):
        self.raw_data = raw_data
        self.week_index = week_index
        return self._build_table()

    def validate_mapping(self, columns):
        missing = set(columns) - set(self.COLUMN_TO_CATEGORY)
        if missing:
            raise ValueError(f"Нет mapping для столбцов: {missing}")

    def get_cell_details(self, judge, column, week_index):
        base_category = self.COLUMN_TO_CATEGORY.get(column)
        if not base_category:
            return []

        weeks = list(self.raw_data.keys())
        week_key = weeks[week_index]
        week_data = self.raw_data.get(week_key, {})
        judge_data = week_data.get(judge)

        if not judge_data:
            return []

        result = []

        # Основная категория
        base_values = judge_data.get(base_category, [])
        if isinstance(base_values, list):
            result.append(("Всего", base_values))

        # В т.ч. рассмотренные в текущем году
        included_category = self.COLUMN_TO_INCLUDED_CATEGORY.get(column)
        if included_category:
            included_values = judge_data.get(included_category, [])
            if isinstance(included_values, list):
                result.append(
                    ("В т.ч. рассмотренные в текущем году", included_values)
                )

        return result