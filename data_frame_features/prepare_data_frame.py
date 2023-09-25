from pandas import DataFrame

from data_frame_features.read_data_frame import read_data_frame
from data_frame_features.filter_data_frame import filter_data_frame
from file_features.message import output_message_exit


def prepare_data_frame(file_name: str, sheet_name: str, column_name: str, filter_target: str) -> DataFrame:
    """
    Читает данные из файла. Фильтрует по значению в столбце.
    :param file_name: Имя файла
    :param sheet_name: Название листа
    :param column_name: индекс столбца типа 'H'
    :param filter_target: значение фильтра
    :return: отфильтрованные данные
    """
    df: DataFrame = read_data_frame(file_name, sheet_name)
    if not df.empty and filter_target:
        df = filter_data_frame(df, column_name, filter_target)
    else:
        output_message_exit(f"в файле: {file_name!r}",f"нет данных на листе: {sheet_name!r}")
    return df
