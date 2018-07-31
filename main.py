from collections import defaultdict, Iterable

import pandas as pd
from pandas import DataFrame
from openpyxl import load_workbook, Workbook  # библиотеки для работы с эксель
from datetime import datetime, timedelta


def parse_pegas(file_name):
    df = pd.read_excel(file_name, sheet_name='Dannie')

    df = df.rename(columns={'Дата включения': 'date',
                            'Код филиала': 'branch_code',
                            'Причина приостановки уплаты взносов': 'cancellation_reason'})

    return df


def parse_drfs(file_name):
    df = pd.read_excel(file_name, sheet_name='Лист1')

    return df


def is_supporting_date(dt: datetime):
    supporting_months = [2, 5, 8, 11]
    if dt.month in supporting_months:
        if dt.day == 1:
            return True
    return False


def date_range(dt: datetime, days: int) -> Iterable[datetime]:
    for i in range(days):
        yield dt
        dt = dt + timedelta(days=1)


def save_to_excel(data: defaultdict) -> None:
    wb = Workbook()
    for branche_code, supporting_dates in data.items():

        ws = wb.create_sheet(title=str(branche_code))
        ws.merge_cells("A1:B1")
        ws["A1"].value = "Для файла Пегас"

        ws["A2"] = "Начало периода"
        ws["B2"] = "Конец периода"

        cell_number = 3
        for support_date in supporting_dates:
            for dt in date_range(support_date, 20):
                ws[f"A{cell_number}"].value = (dt - timedelta(days=60)).strftime("%Y-%m-%d")
                ws[f"B{cell_number}"].value = dt.strftime("%Y-%m-%d")
                cell_number += 1

            cell_number += 1

    wb.save("test.xlsx")


def analyze_pegas_data(data: DataFrame):
    pegas_branch_to_date_map = defaultdict(set)
    for branch_code, group in data.groupby(['branch_code']):
        for i, row in group.iterrows():
            pegas_branch_to_date_map[branch_code].add(row['date'])

    save_to_excel(pegas_branch_to_date_map)


if __name__ == "__main__":
    pegas_file_name = "dataset_pegas_test.xlsx"
    drfs_file_name = "dataset_drfs_test.xlsx"
    fact_file_name = "dataset_fact_test.xlsx"

    pegas_data = parse_pegas(pegas_file_name)
    analyze_pegas_data(pegas_data)

    # drfs_data = parse_drfs(drfs_file_name)

    exit()

    wb2 = load_workbook('dataset_fact_test.xlsx')

    ws2 = wb2['Fact']

    df = pd.DataFrame(ws2.values)  # Преобразуем Sheet в DataFrame

    print(df.truncate())

    data = ws2.values

    cols = next(data)[1:]

    cols_kvart = cols[4:len(cols):4].index  # назване колонок только с кварталами

    data = list(data)  # Данные становятся списком

    idx = [r[0] for r in data]  # заголовки все филиалы
