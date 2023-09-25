from pandas import DataFrame
from file_features.message import output_message_exit, output_message


def filter_data_frame(src_df: DataFrame, column_name: str, target: str) -> DataFrame:
    """ Фильтрует данные DataFrame по столбцу column_name на полное соответствие цели target. """
    if not src_df.empty:
        if column_name in src_df.columns:
            # df = src_df[src_df[column_name].str.strip().str.fullmatch(pat=target)]

            df = src_df.loc[src_df[column_name].isin([target])]

            # print(f"цель {target!r} в df: {len(df.to_records(index=False).tolist())} позиций")
            return df
        else:
            output_message_exit(f"в DataFrame нет колонки с названием:", column_name)
    else:
        output_message(f"в DataFrame нет данных:", f"{src_df.info(verbose=False, memory_usage='deep', show_counts=True)}")
    return src_df


