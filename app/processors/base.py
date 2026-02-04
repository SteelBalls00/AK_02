import re


class BaseProcessor:
    COLUMN_TO_CATEGORY = {}
    COLUMN_TO_INCLUDED_CATEGORY = {}

    columns = []
    categories = []

    word_template_key = None
    specialization = None  # ← ВАЖНО

    def get_specialization(self):
        if not self.specialization:
            raise NotImplementedError(
                f"{self.__class__.__name__} не задал specialization"
            )
        return self.specialization

    def build(self, raw_data, week_index):
        self.raw_data = raw_data
        self.week_index = week_index

        self.COLUMN_TO_CATEGORY = self.build_column_to_category(
            self.columns,
            self.categories
        )

        self.validate_mapping(self.columns)

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
                if column == "Без движения\nсейчас (за год)":
                    result.append(("Оставленные без движения в текущем году", included_values))
                else:
                    result.append(("Рассмотренные в текущем году", included_values))

        return result

    @staticmethod
    def normalize(text: str) -> str:
        return re.sub(r"\s+", " ", text.replace("\n", " ").lower()).strip()

    def build_column_to_category(self, columns, categories):
        mapping = {"Судья": None}

        normalized_categories = {
            self.normalize(cat): cat for cat in categories
        }

        for col in columns:
            if col == "Судья":
                continue

            norm_col = self.normalize(col)
            matched = None

            for norm_cat, original_cat in normalized_categories.items():
                if norm_cat in norm_col or norm_col in norm_cat:
                    matched = original_cat
                    break

            mapping[col] = matched

        return mapping