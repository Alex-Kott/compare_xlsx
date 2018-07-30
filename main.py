import pandas as pd
from pandas import DataFrame
from openpyxl import load_workbook  # библиотеки для работы с эксель


def parse_pegas(file_name):
    df = pd.read_excel(file_name, sheet_name='Dannie')

    df = df.rename(index=str, columns={'Дата включения': 'date',
                                       'Код филиала': 'branch_code',
                                       'Причина приостановки уплаты взносов': 'cancellation_reason'})

    return df


def parse_drfs(file_name):
    df = pd.read_excel(file_name, sheet_name='Лист1')

    return df


def analyze_pegas_data(data: DataFrame):
    for name, group in data.groupby(['branch_code']):
        # print(name)
        for i in group.sort_values(by='date'):
            pass
        print('___________________')


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
