import matplotlib.pyplot as plt
import xlwings as xw


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
        :return: данные вида: {1: 0.00018, 0: 0.012, 3: 6.9008e-06, 4: 3.8e-08, 2: 7.29e-05}
        '''
        uniq = set(sorted([i[1] for i in data]))
        result = dict(zip(uniq, [0] * len(uniq)))

        for item_data in data:
            for item_uniq in uniq:
                if item_data[1]>=item_uniq:
                    result[item_uniq] = result[item_uniq] + item_data[0]

        del result[0] # удалить суммарную вероятность где пострадало 0 человек
        return result

    def _sorted_two_lists(self, first_list: list, second_list: list):
        '''
        Функция синхронной сортировки двух списков, по первому списку

        :param first_list: [1, 3, 4, 2]
        :param second_list: [0.00018, 6.9008e-06, 3.8e-08, 7.292e-05]
        :return: ([1, 2, 3, 4], [0.00018, 7.292e-05, 6.9008e-06, 3.8e-08])
        '''
        # соединим два списка специальной функцией zip
        x = zip(first_list, second_list)
        # отсортируем, взяв первый элемент каждого списка как ключ
        xs = sorted(x, key=lambda tup: tup[0])
        #  извлечем
        soted_first_list = [x[0] for x in xs]
        soted_second_list = [x[1] for x in xs]

        return (soted_first_list, soted_second_list)

    def fn_chart(self):
        # получим данные с листа для построения диаграммы
        data = list(filter(lambda x: x != [None, None], self.ws.range("A2:B2000").value))
        if len(data) < 3:
            return print('Значений для диаграммы должно быть больше')
        else:
            sum_data = self._sum_data_for_fn(data)
            people, probability = self._sorted_two_lists(list(sum_data.keys()), list(sum_data.values()))
            print(people, probability)

            chart_line_x = []
            chart_line_y = []
            for i in people:
                chart_line_x.extend([i - 1, i, i, i])
                chart_line_y.extend([probability[people.index(i)], probability[people.index(i)], None, None])

            # print(chart_line_x, chart_line_y)

            chart_dot_line_x = []
            chart_dot_line_y = []
            for i in people:

                if people[-1] == i:
                    chart_dot_line_x.extend([i, i])
                    chart_dot_line_y.extend([probability[people.index(i)], probability[people.index(i)]])
                    break
                chart_dot_line_x.extend([i, i, None, None])
                chart_dot_line_y.extend(
                    [probability[people.index(i)], probability[people.index(i + 1)], probability[people.index(i + 1)],
                     probability[people.index(i + 1)]])

            # Отрисовка графика
            fig = plt.figure()
            plt.semilogy(chart_line_x, chart_line_y, color='b', linestyle='-', marker='.')
            plt.semilogy(chart_dot_line_x, chart_dot_line_y, color='b', linestyle='--', marker='.')
            plt.xticks(ticks=people)
            plt.title('F/N - диаграмма')
            plt.xlabel('Количество погибших, чел')
            plt.ylabel('Вероятность, 1/год')
            plt.grid(True)
            self.ws.pictures.add(fig, name='FN', update=True, left=self.ws.range('H5').left,
                                 top=self.ws.range('H5').top)
            # plt.show()
            # plt.savefig(f'fn.jpg')

    # def fg_chart(self, data: list):
    #     money = data[1]  # деньги
    #     probability = data[0]  # вероятности
    #     unq = money.copy()  # уникальные значение из money
    #     unq.sort()
    #
    #     sum_ver = []  # сумма вероятностей для >= каждого уникального занчения
    #
    #     for item_unq in unq:
    #         count = 0  # счетчик индекса для выбора индекса соответствующего probability
    #         res = 0  # результат сложения
    #         for iter in money:  # если значение из списка people
    #             if iter >= item_unq:  # больше чем item_unq из уникальных значений
    #                 res += probability[count]  # проссумировать вероятности
    #             count += 1
    #         sum_ver.append(res)
    #
    #     # Координаты для сплошных линий
    #     y = []
    #     for i in sum_ver:
    #         y.extend([i, i, None])
    #     y = y[:-4]
    #
    #     x = []
    #     for i in range(len(unq) - 1):
    #         x.extend([unq[i], unq[i + 1], None])
    #     x.extend([unq[-1], unq[-1], None])
    #     x = x[:-4]
    #
    #     # Координаты для пунктирных линий
    #     y1 = []
    #     for i in range(len(sum_ver) - 1):
    #         y1.extend([sum_ver[i], sum_ver[i + 1], None])
    #     y1 = y1[:-4]
    #
    #     x1 = []
    #     for i in range(len(unq) - 1):
    #         x1.extend([unq[i + 1], unq[i + 1], None])
    #     x1 = x1[:-4]
    #
    #     # Построение графика
    #     fig, ax = plt.subplots()
    #
    #     ax.plot(x, y, 'ro', linewidth=2, linestyle='-')
    #
    #     ax.plot(x1, y1, color='r', linewidth=2, linestyle='--')
    #     ax.set_yscale("log")
    #
    #     #  Прежде чем рисовать вспомогательные линии
    #     #  необходимо включить второстепенные деления
    #     #  осей:
    #     ax.minorticks_on()
    #
    #     #  Определяем внешний вид линий основной сетки:
    #     ax.grid(which='major',
    #             color='k',
    #             linewidth=0.5)
    #
    #     #  Определяем внешний вид линий вспомогательной
    #     #  сетки:
    #     ax.grid(which='minor',
    #             color='k',
    #             linestyle=':')
    #
    #     ax.set_title('F/G - диаграмма')
    #     ax.set_xlabel('Ущерб, млн.руб')
    #     ax.set_ylabel('Вероятность, 1/год')
    #
    #     fig.set_figwidth(12)
    #     fig.set_figheight(8)
    #
    #     # plt.show()
    #     plt.savefig(f'{self.path}\\fg.jpg')


if __name__ == '__main__':
    chart = FN_FG_chart()
    chart.fn_chart()
    #
    # data1 = [
    #     [3e-2, 3e-2, 3e-3, 3e-5, 3e-5],
    #     [1, 2.05, 1.1, 3.25, 5.64]
    # ]
    #
    # chart.fg_chart(data1)
