import pandas as pd


class DataHandler:
    """
    Класс для работы с данными. Используется для преобразования словаря во фрейм pandas,
    сохранения полученного фрейма в файл excel и получения количества строк данных.
    """
    FINANCIAL_FORMAT_INDEX = 44  # индекс для создания формата ячейки excel типа "финансовый"
    NUMERICAL_FORMAT_INDEX = 2  # индекс для создания формата ячейки excel типа "числовой"

    def __init__(self, workbook, logger):
        self.logger = logger
        self.dataframe = pd.DataFrame()  # создаем пустой фрейм данных pandas
        self.ws = workbook.add_worksheet()  # создаем рабочую область в excel
        self.financial_format = workbook.add_format({'num_format': self.FINANCIAL_FORMAT_INDEX})  # формат записи excel
        self.numerical_format = workbook.add_format({'num_format': self.NUMERICAL_FORMAT_INDEX})  # формат записи excel

    def create_dataframe_from_dict(self, data_dict_courses):
        """
        преобразует словарь во фрейм pandas
        :param data_dict_courses: словарь типа {'usd': {'data': 'course'}, ...}
        """
        for currency, data_dict in data_dict_courses.items():
            # создаю фременный фрейм с двумя столбцами для валюты "currency"
            temp_df = pd.DataFrame(data_dict.items(), columns=[f'Дата {currency}', f'Курс {currency}'])
            # добавляю третий столбец с изменениями курса за сутки, с округлением до 4 знаков
            temp_df[f'Изменение {currency}'] = temp_df[f'Курс {currency}'].diff(periods=-1).round(4)
            # объединяю временный фрейм с итоговым
            self.dataframe = pd.concat([self.dataframe, temp_df], axis=1)

        self.dataframe['отношение eur/usd'] = (self.dataframe['Курс eur']/self.dataframe['Курс usd']).round(4)

    def _set_col_widths(self):
        """
        Функция для автоматической установки ширины столбцов в excel, используется внутри класса
        widths: содержит max длину строки в каждом столбце с учетом заголовка
        """
        widths = [max([len(str(s)) for s in self.dataframe[col].values] + [len(col)]) for col in self.dataframe.columns]
        [self.ws.set_column(i, i, width) for i, width in enumerate(widths)]  # устанавливаю ширину каждого столбца

    def _fill_columns_name_to_excel(self):
        """
        функция для записи заголовков в excel, используется внутри класса
        self.dataframe.columns: содержит заголовки данных
        """
        [self.ws.write(0, index, title) for index, title in enumerate(self.dataframe.columns)]

    def _fill_data_to_excel(self, row, col, format_):
        """
        Функция записывает значение по строке и столбцу в excel в выбранном формате,
        пропуская заголовок, используется внутри класса.
        :param row: индекс строки
        :param col: индекс столбца
        :param format_: формат ячейки для записи в excel
        """
        self.ws.write(row + 1, col, self.dataframe.iloc[row, col], format_)

    def save_dataframe_to_excel(self):
        """
        функция для сохранения фрейма данных в файл excel
        """
        try:
            self._set_col_widths()  # установить ширину
            self._fill_columns_name_to_excel()  # записать заголовки

            max_row, _ = self.dataframe.shape  # число строк во фрейме
            for row_ind in range(max_row):
                for col_ind, title in enumerate(self.dataframe.columns):
                    if title in ['Курс usd', 'Курс eur', 'Изменение eur', 'Изменение usd']:
                        self._fill_data_to_excel(row_ind, col_ind, self.financial_format)
                    else:
                        self._fill_data_to_excel(row_ind, col_ind, self.numerical_format)
        except Exception as err:
            self.logger.info(f'При сохранении файла возникла непредвиденная ошибка: {err}')

    def count_rows(self):
        """
        возвращает количество строк с учетом заголовка
        """
        return self.dataframe.shape[0] + 1
