import math

from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import datetime
from pathlib import Path
import xlwings as xw

from config import PATH_,INIQ_TEXT
from FN_FG import F_charts
from Risk_result import get_risk_result, get_components_sum_data




class Report:
    def __init__(self):
        try:
            wb = xw.Book('risk_analysis_prototype.xlsx')
            self.ws_DB = wb.sheets['DB']
            self.ws_MASS = wb.sheets['Масса ОВ']
            self.ws_STATISTICS = wb.sheets['Статистика аварий']
            self.ws_CALC = wb.sheets['Расчет']
            self.ws_CHART = wb.sheets['FN_FG']
            self.data_for_table = self.__get_data_in_CALC_list()


        except FileNotFoundError:
            print('Файл не открыт risk_analysis_prototype.xlsx')
        except:
            print('Ошибка со страницами')

    def temp_explanatory_note(self):
        doc = DocxTemplate(f'{PATH_}\\temp_rpz.docx')
        path_template = Path(__file__).parents[0]
        # создадим общий  словарь для заполнения документа
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
        # 1.9. Таблица с ущербом
        context['damage_table'] = self.__get_data_for_damage_table()
        # 1.10. Анализ риска (п.2.3. РПЗ)
        context['R1_min'] = "{:.2e}".format(min([i[7] for i in self.data_for_table]))
        context['R1_max'] = "{:.2e}".format(max([i[7] for i in self.data_for_table]))
        # НОВЫЙ РЕЗУЛЬТАТ РИСКА
        data = get_risk_result.Risk().risk_result()
        context['Risk'] = self.__get_risk_data_for_table(data)

        self.__get_fn_fg_chart()
        context['fn'] = InlineImage(doc, f'{path_template}\\fn.jpg', width=Mm(140))
        context['fg'] = InlineImage(doc, f'{path_template}\\fg.jpg', width=Mm(140))

        # 1.11. Выводы РПЗ (п.2.3. РПЗ) NEW!
        data = get_components_sum_data.Components_sum_data().components_result()
        context['Sum_data'] = self.__get_components_data_for_table(data)

        # оставляем параметры риска децибелы общими для объекта
        most_possible = self.__get_most_possible_scenario()
        most_dangerous = self.__get_most_dangerous_scenario()
        context['probability_end'] = "{:.2e}".format(most_possible[7])
        context['damage_end'] = round(most_dangerous[47], 2)
        _temp = sum([i[48] for i in self.data_for_table]) / self.ws_DB.range("B23").value
        context['RdB'] = round(10 * math.log10(_temp * pow(10, 6) / 195), 2)
        context['Rng'] = "{:.2e}".format((_temp) * pow(10, 6))
        context['Rng2'] = "{:.2e}".format((_temp) * pow(10, 5))

        # Заполним документ из словаря

        doc.render(context)
        doc.save(f'{path_template}\\add_docs\\RPZ_{INIQ_TEXT}.docx')

    def temp_declaration_note(self):
        doc = DocxTemplate(f'{PATH_}\\temp_dpb.docx')
        path_template = Path(__file__).parents[0]
        # создадим общий словарь для заполнения документа
        context = {}
        # 1. данные для словаря
        context['year'] = datetime.date.today().year
        # 1.1. данные общие
        context |= self.__get_data_in_DB_list()
        # 1.2. данные по количеству опасного вещества
        context['mass_sub_table'] = self.__get_data_in_MASS_list()
        context['sum_sub'] = round(sum([float(i['Quantity']) for i in context['mass_sub_table']]), 2)
        # 1.3. Заполнение данных по наиболее вероятному и опасному сценарию
        # НОВЫЙ РЕЗУЛЬТАТ РИСКА
        data = get_risk_result.Risk().risk_result()
        context['Risk'] = self.__get_risk_data_for_table(data)
        data = get_components_sum_data.Components_sum_data().components_result()
        context['Sum_data'] = self.__get_components_data_for_table(data)

        # 1.4. Анализ риска (п.2.3.2 ДПБ)
        context['R1_min'] = "{:.2e}".format(min([i[7] for i in self.data_for_table]))
        context['R1_max'] = "{:.2e}".format(max([i[7] for i in self.data_for_table]))


        # диаграммы отрисованы temp_explanatory_note
        context['fn'] = InlineImage(doc, f'{path_template}\\fn.jpg', width=Mm(140))
        context['fg'] = InlineImage(doc, f'{path_template}\\fg.jpg', width=Mm(140))
        # 1.11. Выводы ДПБ (п.4.1. РПЗ)
        # оставляем параметры риска децибелы общими для объекта
        most_possible = self.__get_most_possible_scenario()
        most_dangerous = self.__get_most_dangerous_scenario()
        context['probability_end'] = "{:.2e}".format(most_possible[7])
        context['damage_end'] = round(most_dangerous[47], 2)
        _temp = sum([i[48] for i in self.data_for_table]) / self.ws_DB.range("B23").value
        context['RdB'] = round(10 * math.log10(_temp * pow(10, 6) / 195), 2)
        context['Rng'] = "{:.2e}".format((_temp) * pow(10, 6))
        context['Rng2'] = "{:.2e}".format((_temp) * pow(10, 5))

        # Заполним документ из словаря
        doc.render(context)
        doc.save(f'{path_template}\\add_docs\\DPB_{INIQ_TEXT}.docx')

    def temp_info_note(self):
        doc = DocxTemplate(f'{PATH_}\\temp_ifl.docx')
        path_template = Path(__file__).parents[0]
        # создадим общий словарь для заполнения документа
        context = {}
        # 1. данные для словаря
        context['year'] = datetime.date.today().year
        # 1.1. данные общие
        context |= self.__get_data_in_DB_list()
        # диаграммы отрисованы temp_explanatory_note
        context['fn'] = InlineImage(doc, f'{path_template}\\fn.jpg', width=Mm(140))
        context['fg'] = InlineImage(doc, f'{path_template}\\fg.jpg', width=Mm(140))
        # 1.2. Выводы ИФЛ
        data = get_risk_result.Risk().risk_result()
        context['Risk'] = self.__get_risk_data_for_table(data)
        data = get_components_sum_data.Components_sum_data().components_result()
        context['Sum_data'] = self.__get_components_data_for_table(data)

        # оставляем параметры риска децибелы общими для объекта
        most_possible = self.__get_most_possible_scenario()
        most_dangerous = self.__get_most_dangerous_scenario()
        context['probability_end'] = "{:.2e}".format(most_possible[7])
        context['damage_end'] = round(most_dangerous[47], 2)
        _temp = sum([i[48] for i in self.data_for_table]) / self.ws_DB.range("B23").value
        context['RdB'] = round(10 * math.log10(_temp * pow(10, 6) / 195), 2)
        context['Rng'] = "{:.2e}".format((_temp) * pow(10, 6))
        context['Rng2'] = "{:.2e}".format((_temp) * pow(10, 5))

        # Заполним документ из словаря
        doc.render(context)
        doc.save(f'{path_template}\\add_docs\\IFL_{INIQ_TEXT}.docx')

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
            # print('Pozition', item[1], ' Quantity', item[3])
            devices = {'Completion': item[15],
                       'Locations': item[0],
                       'Pozition': item[1],
                       'Temperature': item[7],
                       'Volume': item[14],
                       'Pressure': item[6],
                       'Diameter': item[10],
                       'Length': item[9],
                       'Flow': item[11],
                       'Volume_pipe': round(item[8], 3),
                       'Quantity': round(item[3], 3),
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
                       'Men1': int(item[35]),
                       'Men2': int(item[36])}
            result_list.append(devices)
        return result_list

    def __get_data_for_damage_table(self):
        data = self.data_for_table
        result_list = []
        for item in data:
            devices = {'Sc': item[0],
                       'Unit': item[1],
                       'D1': round(item[42], 2),
                       'D2': round(item[43], 2),
                       'D3': round(item[44], 2),
                       'D4': round(item[45], 2),
                       'D5': round(item[46], 2),
                       'Sum': round(item[47], 2)}
            result_list.append(devices)
        return result_list

    def __get_risk_data_for_table(self, data):
        result_list = []
        for facility_name, risk_data in data.items():
            row = {
                'Facility': facility_name,
                'IR_death': f"{risk_data['Индивидуальный риск гибели']:.2e}",
                'IR_injury': f"{risk_data['Индивидуальный риск ранения']:.2e}",
                'CR_death': f"{risk_data['Коллективный риск гибели']:.2e}",
                'CR_injury': f"{risk_data['Коллективный риск ранения']:.2e}",
                'Max_total_damage': round(risk_data['Максимальный суммарный ущерб'], 2),
                'Max_eco_damage': round(risk_data['Максимальный экологический ущерб'], 2)
            }
            result_list.append(row)
        return result_list

    def __get_components_data_for_table(self, data):
        result_list = []
        for facility_name, scenarios in data.items():
            # Предполагаем, что первый сценарий - наиболее опасный, второй - наиболее вероятный
            for index, scenario in enumerate(scenarios):
                scenario_type = "Наиболее опасный" if index == 0 else "Наиболее вероятный"

                row = {
                    'Facility': facility_name,
                    'ScenarioType': scenario_type,
                    'ScenarioNum': scenario['Номер сценария'],
                    'Equipment': scenario['Оборудование'],
                    'Scenario': scenario['Сценарий'],
                    'Probability': f"{scenario['Вероятность']:.2e}",
                    'AccidentAmount': round(scenario['Количество в аварии'], 2),
                    'Deaths': int(scenario['Погибшие']),
                    'Injuries': int(scenario['Пострадавшие']),
                    'Damage': round(scenario['Ущерб'], 2),
                    'EcoDamage': round(scenario['Ущерб экол.'], 2)
                }
                result_list.append(row)
        return result_list

    def __get_fn_fg_chart(self):
        self.ws_CHART.range("A2:E2000").clear_contents()
        data_probability = [i[7] for i in self.data_for_table]
        data_damage = [i[47] for i in self.data_for_table]
        data_dead = [i[35] for i in self.data_for_table]
        self.ws_CHART.range(f"A2").options(transpose=True).value = data_probability
        self.ws_CHART.range(f"B2").options(transpose=True).value = data_dead
        self.ws_CHART.range(f"D2").options(transpose=True).value = data_probability
        self.ws_CHART.range(f"E2").options(transpose=True).value = data_damage
        chart = F_charts.FN_FG_chart()
        chart.fn_chart()
        chart.fg_chart()

    def __get_most_possible_scenario(self):
        data = self.data_for_table
        search = max([i[7] for i in self.data_for_table])
        for item in data:
            if item[7] == search:
                return item
        return [0 for _ in range(len(data[0]))]

    def __get_most_dangerous_scenario(self):
        data = self.data_for_table
        search = max([i[35] for i in self.data_for_table])
        for item in data:
            if item[35] == search:
                return item
        return [0 for _ in range(len(data[0]))]


if __name__ == '__main__':
    Report().temp_explanatory_note()
    Report().temp_declaration_note()
    Report().temp_info_note()