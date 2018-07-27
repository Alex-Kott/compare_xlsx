import pandas as pd
from openpyxl import load_workbook  # библиотеки для работы с эксель
from openpyxl.utils import get_column_letter, column_index_from_string




def parse_pegas(file_name):
    pegas_wb = load_workbook(file_name)
    worksheet = pegas_wb['Dannie']
    df = pd.DataFrame(worksheet.values)

    df = df.rename(index=str, columns={0: 'date',
                                       1: 'branch_code',
                                       2: 'cancellation_reason'})

    return df.groupby(['branch_code', 'date']).size()



if __name__ == "__main__":
    pegas_file_name = "dataset_pegas_test.xlsx"
    drfs_file_name = "dataset_drfs_test.xlsx"
    fact_file_name = "dataset_fact_test.xlsx"


    parse_pegas(pegas_file_name)
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
