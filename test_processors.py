from app.repository.statistics import StatisticsRepository
from app.factory.processor_factory import ProcessorFactory


def test_pkl(pkl_path, week_index=0):
    repo = StatisticsRepository()
    raw_data, context = repo.load(pkl_path)

    processor = ProcessorFactory.get(context)
    table = processor.build(raw_data, week_index)

    print("\nНеделя:", table["week"])
    print("Колонки:")
    for col in table["columns"]:
        print(" -", col)

    print("\nПервые строки:")
    for row in table["rows"][:3]:
        print(row)

    print("\nИтого:")
    print(table["total"])


if __name__ == "__main__":
    test_pkl(
        r"\\192.168.0.200\Minato\3\Амурский областной суд\result4_with_33.pkl",
        week_index=0
    )
