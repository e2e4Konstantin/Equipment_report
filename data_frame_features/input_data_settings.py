from data_frame_features.read_data_frame import read_data_frame
from data_frame_features.info_data_frame import info_data_frame
from file_features.message import output_message_exit


class EquipmentsSourceData:
    sheets = ['Tables', 'Equipment', 'Attributes', 'Options']

    def __init__(self, input_file_name: str):
        self.file = input_file_name
        self.tables = read_data_frame(input_file_name, sheet_name=self.sheets[0])
        self.equipments = read_data_frame(input_file_name, sheet_name=self.sheets[1])
        self.attributes = read_data_frame(input_file_name, sheet_name=self.sheets[2])
        self.options = read_data_frame(input_file_name, sheet_name=self.sheets[3])
        self.check()

    def check(self):
        file_message = f"в файле: {self.file!r}"
        if self.tables.empty:
            output_message_exit(file_message, f"нет ни одной таблицы на листе {self.sheets[0]!r}.")
        if self.equipments.empty:
            output_message_exit(file_message, f"нет данных на листе {self.sheets[1]!r}.")
        if self.attributes.empty:
            output_message_exit(file_message, f"нет атрибутов на листе {self.sheets[2]!r}.")
        if self.options.empty:
            output_message_exit(file_message, f"нет параметров на листе {self.sheets[3]!r}.")

    def info(self):
        info_data_frame(self.tables)
        info_data_frame(self.equipments)
        info_data_frame(self.attributes)
        info_data_frame(self.options)
