from docxtpl import DocxTemplate, InlineImage
import datetime
from pathlib import Path
import time
import xlwings as xw

INIQ_TEXT = str(int(time.time()))


class Report:
    def __init__(self):
        try:
            wb = xw.Book('risk_analysis_prototype.xlsx')
            self.ws_DB = wb.sheets['DB']
            self.ws_MASS = wb.sheets['Масса ОВ']

        except FileNotFoundError:
            print('Файл не открыт risk_analysis_prototype.xlsx')

    def temp_explanatory_note(self):
        path_template = Path(__file__).parents[0]
        # создадим общий словарь для заполнения документа
        context = {}
        # 1. данные для словаря
        context['year'] = datetime.date.today().year
        # 1.1. данные общие
        context |= self.__get_data_in_DB_list()
        # 1.2. данные по оборудованию
        # Таблица с оборудованием
        context['dev_table'] = [x for x in self.__get_data_in_MASS_list() if x['Volume_pipe'] == 0]
        context['pipe_table'] = [x for x in self.__get_data_in_MASS_list() if x['Volume_pipe'] != 0]
        # 1.3. данные по количеству опасного вещества
        context['mass_sub_table'] = self.__get_data_in_MASS_list()
        context['sum_sub'] = round(sum([float(i['Quantity']) for i in context['mass_sub_table']]),2)

        doc = DocxTemplate(f'temp_rpz.docx')
        # Заполним документ из словаря
        doc.render(context)
        doc.save(f'{path_template}\\add_docs\\RPZ_{INIQ_TEXT}.docx')

    def __get_data_in_DB_list(self) -> dict:
        # получим данные с листа данных проекта
        data = self.ws_DB.range("A2:C22").value
        result_dict = {}
        for item in data:
            if item[1] == None: break
            result_dict[item[2]] = item[1]
        return result_dict

    def __get_data_in_MASS_list(self) -> list:
        # получим данные с листа c массами данных
        data = self.ws_MASS.range("A3:R22").value

        result_list = []
        for item in data:
            if item[0] == None: break
            devices = {'Completion': item[15],
                       'Locations': item[0],
                       'Pozition': item[1],
                       'Temperature': item[7],
                       'Volume': item[14],
                       'Pressure': item[6],
                       'Diameter': item[10],
                       'Length': item[9],
                       'Flow': item[11],
                       'Volume_pipe': round(item[8],1),
                       'Quantity': round(item[3],1),
                       'State': item[5]}
            result_list.append(devices)
        return result_list


if __name__ == '__main__':
    r = Report().temp_explanatory_note()
