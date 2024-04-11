import math

from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import datetime
from pathlib import Path
import time
import xlwings as xw
import matplotlib.pyplot as plt
import numpy as np

INIQ_TEXT = str(int(time.time()))


class FN_FG_chart:
    '''
    Класс построения F/N  и  F/G диаграмм.

    Для построения диграммы должен быть открыт файл
    1) risk_analysis_prototype.xlsx
    2) на листе "FN_FG" заполненны таблицы значений

    '''

    def __init__(self):
        try:
            wb = xw.Book('risk_analysis_prototype.xlsx')
            self.ws = wb.sheets['FN_FG']

        except FileNotFoundError:
            print('Файл не открыт risk_analysis_prototype.xlsx')

    def _sum_data_for_fn(self, data: list):
        '''
        Функция вычисления суммирования вероятностей F при которой пострадало не менее N человек
        :param data: данные вида [[3.8e-08, 1],[5.8e-08, 2],[1.1e-08, 1]..]
        :return: данные вида: {1: 0.00018, 2: 0.012, 3: 6.9008e-06, 4: 3.8e-08, 2: 7.29e-05}
        '''
        uniq = set(sorted([i[1] for i in data]))
        result = dict(zip(uniq, [0] * len(uniq)))

        for item_data in data:
            for item_uniq in uniq:
                if item_data[1] >= item_uniq:
                    result[item_uniq] = result[item_uniq] + item_data[0]

        del result[0]  # удалить суммарную вероятность где пострадало 0 человек
        return result

    def _sum_data_for_fg(self, data: list):
        '''
        Функция вычисления суммирования вероятностей F при которой ущерб не менее G млн.руб
        :param data: данные вида [[3.8e-08, 1.2],[5.8e-08, 0.2],[1.1e-08, 12.4]..]
        :return: данные вида: {0.2: 0.00018, 1: 0.012, 3: 6.9008e-06, 5: 3.8e-08, 6.25: 7.29e-05}
        '''
        uniq = np.arange(0, max([i[1] for i in data]) + max([i[1] for i in data]) / 7, max([i[1] for i in data]) / 7)

        result = dict(zip(uniq, [0] * len(uniq)))

        for item_data in data:
            for item_uniq in uniq:
                if item_data[1] >= item_uniq:
                    result[item_uniq] = result[item_uniq] + item_data[0]

        del result[0]  # удалить суммарную вероятность где ущерб 0
        return result

    def fn_chart(self):
        """
        Построение FN диаграммы
        :return:
        """
        # получим данные с листа для построения диаграммы
        data = list(filter(lambda x: x != [None, None], self.ws.range("A2:B2000").value))
        if len(data) < 3:
            return print('Значений для диаграммы должно быть больше')
        else:
            sum_data = self._sum_data_for_fn(data)
            people, probability = list(sum_data.keys()), list(sum_data.values())
            # для сплошных линий
            chart_line_x = []
            chart_line_y = []
            for i in people:
                chart_line_x.extend([i - 1, i, i, i])
                chart_line_y.extend([probability[people.index(i)], probability[people.index(i)], None, None])
            # для пунктирных линий
            chart_dot_line_x = []
            chart_dot_line_y = []
            for i in people:
                if i == people[-1]:
                    chart_dot_line_x.extend([i, i])
                    chart_dot_line_y.extend([probability[people.index(i)], 0])
                    break
                chart_dot_line_x.extend([i, i])
                chart_dot_line_y.extend([probability[people.index(i)], probability[people.index(i) + 1]])

            # Отрисовка графика
            fig = plt.figure()
            plt.semilogy(chart_line_x, chart_line_y, color='b', linestyle='-', marker='.')
            plt.semilogy(chart_dot_line_x, chart_dot_line_y, color='b', linestyle='--', marker='.')
            plt.xticks(ticks=people)
            plt.title('F/N - диаграмма')
            plt.xlabel('Количество погибших, чел')
            plt.ylabel('Вероятность, 1/год')
            plt.grid(True)
            # self.ws.pictures.add(fig, name='FN', update=True, left=self.ws.range('H5').left,
            #                      top=self.ws.range('H5').top)
            # plt.show()
            plt.savefig(f'fn.jpg')

    def fg_chart(self):
        """
        Построение FG диаграммы
        :return:
        """
        # получим данные с листа для построения диаграммы
        data = list(filter(lambda x: x != [None, None], self.ws.range("D2:E2000").value))
        if len(data) < 3:
            return print('Значений для диаграммы должно быть больше')
        else:
            sum_data = self._sum_data_for_fg(data)
            damage, probability = list(sum_data.keys()), list(sum_data.values())
            # для сплошных линий
            chart_line_x = []
            chart_line_y = []
            for i in damage:
                if damage[0] == i:
                    chart_line_x.extend([0, i, i, i])
                    chart_line_y.extend([probability[damage.index(i)], probability[damage.index(i)], None, None])
                elif damage[-1] == i:
                    chart_line_x.extend([damage[damage.index(i) - 1], damage[damage.index(i) - 1], i, i])
                    chart_line_y.extend(
                        [probability[damage.index(i)], probability[damage.index(i)], probability[damage.index(i)],
                         probability[damage.index(i)]])
                    break
                else:
                    chart_line_x.extend([damage[damage.index(i) - 1], i, i, i])
                    chart_line_y.extend([probability[damage.index(i)], probability[damage.index(i)], None, None])

            # для пунктирных линий
            chart_dot_line_x = []
            chart_dot_line_y = []
            for i in damage:
                if i == damage[-1]:
                    chart_dot_line_x.extend([i, i])
                    chart_dot_line_y.extend([probability[damage.index(i)], probability[damage.index(i)]])
                    chart_dot_line_x.extend([i, i])
                    chart_dot_line_y.extend([probability[damage.index(i)], 0])
                    break
                chart_dot_line_x.extend([i, i])
                chart_dot_line_y.extend([probability[damage.index(i)], probability[damage.index(i) + 1]])

            # Отрисовка графика
            fig = plt.figure()
            plt.semilogy(chart_line_x, chart_line_y, color='r', linestyle='-', marker='.')
            plt.semilogy(chart_dot_line_x, chart_dot_line_y, color='r', linestyle='--', marker='.')
            # plt.xticks(ticks=damage)
            plt.title('F/G - диаграмма')
            plt.xlabel('Ущерб, млн.руб')
            plt.ylabel('Вероятность, 1/год')
            plt.grid(True)
            # self.ws.pictures.add(fig, name='FG', update=True, left=self.ws.range('Q5').left,
            #                      top=self.ws.range('Q5').top)
            # plt.show()
            plt.savefig(f'fg.jpg')


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
        doc = DocxTemplate(f'temp_rpz.docx')
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
        # 1.9. Таблица с ущербом
        context['damage_table'] = self.__get_data_for_damage_table()
        # 1.10. Анализ риска (п.2.3. РПЗ)
        context['R1_min'] = "{:.2e}".format(min([i[7] for i in self.data_for_table]))
        context['R1_max'] = "{:.2e}".format(max([i[7] for i in self.data_for_table]))
        context['R_koll_dead'] = "{:.2e}".format(sum([i[48] for i in self.data_for_table]))
        context['R_koll_injury'] = "{:.2e}".format(sum([i[49] for i in self.data_for_table]))
        context['R_ind_dead'] = "{:.2e}".format(
            sum([i[48] for i in self.data_for_table]) / self.ws_DB.range("B23").value)
        context['R_ind_injury'] = "{:.2e}".format(
            sum([i[49] for i in self.data_for_table]) / self.ws_DB.range("B23").value)
        context['R_ecol'] = round(max([i[46] for i in self.data_for_table]), 2)
        context['R_sum'] = round(max([i[47] for i in self.data_for_table]), 2)
        self.__get_fn_fg_chart()
        context['fn'] = InlineImage(doc, f'{path_template}\\fn.jpg', width=Mm(140))
        context['fg'] = InlineImage(doc, f'{path_template}\\fg.jpg', width=Mm(140))
        # 1.11. Выводы РПЗ (п.2.3. РПЗ)
        most_possible = self.__get_most_possible_scenario()
        context[
            'most_possible'] = f'''сценарий: {most_possible[0]}, оборудование: {most_possible[1]}, частота {"{:.2e}".format(most_possible[2])} 1/год, ущерб: {round(most_possible[3], 2)} млн.руб.'''
        most_dangerous = self.__get_most_dangerous_scenario()
        context[
            'most_dangerous'] = f'''сценарий: {most_dangerous[0]}, оборудование: {most_dangerous[1]}, частота {"{:.2e}".format(most_dangerous[2])} 1/год, ущерб: {round(most_dangerous[3], 2)} млн.руб.'''

        context['probability_end'] = "{:.2e}".format(most_possible[2])
        context['damage_end'] = round(most_dangerous[3], 2)
        _temp = sum([i[48] for i in self.data_for_table]) / self.ws_DB.range("B23").value
        context['RdB'] = round(10 * math.log10(_temp * pow(10, 5) / 195),2)
        context['Rng'] = "{:.2e}".format((_temp) * pow(10, 6))
        context['Rng2'] = "{:.2e}".format((_temp) * pow(10, 5))

        # Заполним документ из словаря

        doc.render(context)
        doc.save(f'{path_template}\\add_docs\\RPZ_{INIQ_TEXT}.docx')

    def temp_declaration_note(self):
        doc = DocxTemplate(f'temp_dpb.docx')
        path_template = Path(__file__).parents[0]
        # создадим общий словарь для заполнения документа
        context = {}
        # 1.1. данные по количеству опасного вещества
        context['mass_sub_table'] = self.__get_data_in_MASS_list()
        context['sum_sub'] = round(sum([float(i['Quantity']) for i in context['mass_sub_table']]), 2)

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

    def __get_fn_fg_chart(self):
        self.ws_CHART.range("A2:E2000").clear_contents()
        data_probability = [i[7] for i in self.data_for_table]
        data_damage = [i[47] for i in self.data_for_table]
        data_dead = [i[35] for i in self.data_for_table]
        self.ws_CHART.range(f"A2").options(transpose=True).value = data_probability
        self.ws_CHART.range(f"B2").options(transpose=True).value = data_dead
        self.ws_CHART.range(f"D2").options(transpose=True).value = data_probability
        self.ws_CHART.range(f"E2").options(transpose=True).value = data_damage
        chart = FN_FG_chart()
        chart.fn_chart()
        chart.fg_chart()

    def __get_most_possible_scenario(self):
        data = self.data_for_table
        search = max([i[7] for i in self.data_for_table])
        for item in data:
            if item[7] == search:
                return (item[0], item[1], item[7], item[47])
        return (0, 0, 0, 0)

    def __get_most_dangerous_scenario(self):
        data = self.data_for_table
        search = max([i[35] for i in self.data_for_table])
        for item in data:
            if item[35] == search:
                return (item[0], item[1], item[7], item[47])
        return (0, 0, 0, 0)

    def pss(self):
        data = self.data_for_table
        search = max([i[7] for i in self.data_for_table])
        print(search)
        print((any(e[7] == search for e in data)))


if __name__ == '__main__':
    r = Report().temp_explanatory_note()
    # r = Report().pss()
