import re

import pandas
from pandas.core.frame import DataFrame


def get_barcode_data_from_xsl(file_name):
    data_from_excel = pandas.read_excel(file_name, header=None)
    for data in data_from_excel.values.tolist():
        print(data)
        if len(data) == 2:
            yield data


def find_col_with_phone_number(df: DataFrame) -> int:
    i = 0
    for col in df.iloc[0]:
        if re.search('^[+]*\\d{10,11}$', str(col)):
            return i
        i += 1
    return -1


def remove_duplicates(df: DataFrame, col) -> DataFrame:
    return df.drop_duplicates(subset=[col])


def remove_by_values(df: DataFrame, col, values: list) -> DataFrame:
    values = [int(x) for x in values]
    df = df[~df[col].isin(values)]
    return df


def save_data_to_file(df: DataFrame, result_file_name):
    df.to_excel(result_file_name)



def data_handling(file_name, result_file_name, values: list):
    df = pandas.read_excel(file_name, header=None)

    col_number = find_col_with_phone_number(df)
    if col_number < 0:
        raise ValueError(f'Phone not founded in file {file_name}')

    df = remove_duplicates(df, col_number)
    df = remove_by_values(df, col_number, values)

    save_data_to_file(df, result_file_name)


if __name__ == "__main__":
    fake_phones = ['9052072342', '9030266669']
    f = r"C:\Users\pavel.suhanov\PycharmProjects\add-barcode\phone_test.xlsx"
    res = r"C:\Users\pavel.suhanov\PycharmProjects\add-barcode\phone_test_result.xlsx"
    data_handling(f, res, fake_phones)
