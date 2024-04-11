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
            self.ws_STATISTICS = wb.sheets['Статистика аварий']
            self.ws_CALC = wb.sheets['Расчет']
            self.data_for_table = self.__get_data_in_CALC_list()

        except FileNotFoundError:
            print('Файл не открыт risk_analysis_prototype.xlsx')
        except:
            print('Ошибка со страницами')

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
        context['sum_sub'] = round(sum([float(i['Quantity']) for i in context['mass_sub_table']]), 2)
        # 1.4. данные по статистике аварий
        context['oil_tank_accident_table'] = self.__get_data_in_STATISTICS_list()[0]
        context['oil_pipeline_accident_table'] = self.__get_data_in_STATISTICS_list()[1]
        # 1.5. Таблица с описанием сценариев и частотами
        context['scenario_table'] = self.__get_data_for_scenario_table()
        # 1.6. Таблица с массами участвует в аварии и в поражающем факторе
        context['mass_table'] = self.__get_data_for_mass_table()
        # 1.7. Таблица с расчетами
        context['calc_table'] = self.__get_data_for_calc_table()
        # 1.8. Таблица с погибшими
        context['dead_table'] = self.__get_data_for_dead_table()

        doc = DocxTemplate(f'temp_rpz.docx')
        # Заполним документ из словаря
        doc.render(context)
        doc.save(f'{path_template}\\add_docs\\RPZ_{INIQ_TEXT}.docx')

    def __get_data_in_DB_list(self) -> dict:
        # получим данные с листа данных проекта
        data = self.ws_DB.range("A2:C300").value
        result_dict = {}
        for item in data:
            if item[1] == None: break
            result_dict[item[2]] = item[1]
        return result_dict

    def __get_data_in_MASS_list(self) -> list:
        # получим данные с листа c массами данных
        data = self.ws_MASS.range("A3:R300").value

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
                       'Volume_pipe': round(item[8], 1),
                       'Quantity': round(item[3], 1),
                       'State': item[5]}
            result_list.append(devices)
        return result_list

    def __get_data_in_STATISTICS_list(self) -> tuple:
        # получим данные с листа статистики
        data_tank = self.ws_STATISTICS.range("A2:F10").value
        data_pipe = self.ws_STATISTICS.range("H2:M10").value

        result_list_tank = []
        for item in data_tank:
            if item[0] == None: break
            accident = {'Num': item[0],
                        'Date': item[1],
                        'View': item[2],
                        'Description': item[3],
                        'Scale': item[4],
                        'Damage': item[5]}
            result_list_tank.append(accident)

        result_list_pipe = []
        for item in data_pipe:
            if item[0] == None: break
            accident = {'Num': item[0],
                        'Date': item[1],
                        'View': item[2],
                        'Description': item[3],
                        'Scale': item[4],
                        'Damage': item[5]}
            result_list_pipe.append(accident)

        return (result_list_tank, result_list_pipe)

    def __get_data_in_CALC_list(self):
        # получим данные с листа данных Расчет
        data = [i for i in self.ws_CALC.range("A2:BA1000").value if i[0] is not None]
        i = 1
        for k in data:
            k[0] = f'C{i}'
            i += 1
        return data

    def __get_data_for_scenario_table(self):
        data = self.data_for_table

        result_list = []
        for item in data:
            devices = {'Sc': item[0],
                       'Unit': item[1],
                       'Sc_text': item[2],
                       'Сhance': "{:.2e}".format(item[7])}
            result_list.append(devices)
        return result_list

    def __get_data_for_mass_table(self):
        data = self.data_for_table

        result_list = []
        for item in data:
            devices = {'Sc': item[0],
                       'Unit': item[1],
                       'M1': round(item[8], 3),
                       'M2': round(item[9], 3)}
            result_list.append(devices)
        return result_list

    def __get_data_for_calc_table(self):
        data = self.data_for_table
        result_list = []
        for item in data:
            devices = {'Sc': item[0],
                       'Unit': item[1],
                       'q_10': item[15],
                       'q_7': item[16],
                       'q_4': item[17],
                       'q_1': item[18],
                       'P_28': item[20],
                       'P_14': item[21],
                       'P_5': item[22],
                       'P_2': item[23],
                       'Lf': item[24],
                       'Df': item[25],
                       'Rnkpr': item[26],
                       'Rvsp': item[27],
                       'Lpt': item[28],
                       'Ppt': item[29],
                       'Q600': item[30],
                       'Q320': item[31],
                       'Q220': item[32],
                       'Q120': item[33]
                       }
            result_list.append(devices)
        return result_list

    def __get_data_for_dead_table(self):
        data = self.data_for_table
        result_list = []
        for item in data:
            devices = {'Sc': item[0],
                       'Unit': item[1],
                       'Men1': item[35],
                       'Men2': item[36]}
            result_list.append(devices)
        return result_list

if __name__ == '__main__':
    r = Report().temp_explanatory_note()
