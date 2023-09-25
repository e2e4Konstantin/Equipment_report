import os
import pandas as pd
from pandas import DataFrame

from file_features.message import output_message_exit
from data_frame_features.generate_column_names import generate_column_names


def read_data_frame(src_file_name: str, sheet_name: str) -> DataFrame | None:
    """
    Читает данные из Excel файла в pandas.DataFrame(), присваивает название столбцам и строкам как в Excel.
    Строки нумеруем с 1, столбцы: 'A', 'B'....
    Создает упакованный файл формата parquet. Если уже есть такой gzip файл читает из него.
    :param src_file_name: Файл с данными
    :param sheet_name: Имя таблицы
    :return: Экземпляр класса pandas.DataFrame()
    """

    file_path, file_name = os.path.split(src_file_name)
    parquet_file = os.path.join(file_path, f"{file_name.split('.')[0]}_{sheet_name}.gzip")
    df = DataFrame()
    if os.path.exists(parquet_file):
        try:
            df: DataFrame = pd.read_parquet(parquet_file)
            columns = df.columns
            df = df[columns].astype(pd.StringDtype())
        except IOError as err:
            output_message_exit(str(err), parquet_file)
    else:
        try:
            df: DataFrame = pd.read_excel(io=src_file_name, sheet_name=sheet_name)
            columns = df.columns
            df = df[columns].astype(pd.StringDtype())
            column_names = generate_column_names(len(columns)+1)
            df.rename(columns=dict(zip(columns, column_names)), inplace=True, errors="raise")
            df.index = range(1, len(df.index) + 1)
            df.to_parquet(parquet_file, compression='gzip')
        except IOError as err:
            output_message_exit(str(err), src_file_name)
    if not df.empty:
        return df
    return None
