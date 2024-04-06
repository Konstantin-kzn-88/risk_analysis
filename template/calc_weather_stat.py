import xlwings as xw
import matplotlib.pyplot as plt


class Weather_calc:
    '''
    Класс расчута статистики погоды
    '''

    def __init__(self):
        try:
            wb = xw.Book('risk_analysis_prototype.xlsx')
            self.ws_WEATHER = wb.sheets['Погода 2023 (Казань)']


        except FileNotFoundError:
            print('Файл не открыт risk_analysis_prototype.xlsx')

    def calc_statistics(self):
        # получим данные с листа статистики
        data = [i for i in self.ws_WEATHER.range("A3:H500").value if i[0] is not None]
        t_10 = t_20 = t_39 = 0  # температуры 10,20 и 39 град
        v_1 = v_2 = v_3 = 0  # скорости ветра 1,2 и 3 м/с

        for i in data:
            if i[1] > -40 and i[1] < 20:
                t_10 += 1
            elif i[1] > 20 and i[1] < 38:
                t_20 += 1
            elif i[1] > 38 and i[1] < 60:
                t_39 += 1

            if i[5] <= 1:
                v_1 += 1
            elif i[5] >=2 and i[1] < 3:
                v_2 += 1
            elif i[5] >= 3:
                v_3 += 1

        print(t_10/len(data),t_20/len(data),t_39/len(data))
        print(v_1/len(data), v_2/len(data), v_3/len(data))

        index_wind_velosity = ['<=1 м/с', '2 м/с', '>=3 м/с']
        index_temperature = ['<=10 град.С', '20 град.С', 't_max град.С']
        temperature = [t_10/len(data),t_20/len(data),t_39/len(data)]
        wind_velosity = [v_1/len(data), v_2/len(data), v_3/len(data)]
        # Построение графиков
        plt.figure(figsize=(9, 9))
        plt.subplot(2, 1, 1)
        plt.bar(index_wind_velosity, wind_velosity, color='c')  # построение графика
        plt.title('Вероятность возникновения температуры и скорости ветра')  # заголовок
        plt.ylabel('Вероятность скорости ветра, м/с', fontsize=14)  # ось ординат
        plt.subplot(2, 1, 2)
        plt.bar(index_temperature, temperature, color='b')  # построение графика
        plt.ylabel('Вероятность температуры, град.С', fontsize=14)  # ось ординат
        plt.show()


    def show_statistics_graph(self):
        # получим данные с листа статистики

        data = [i for i in self.ws_WEATHER.range("A3:H500").value if i[0] is not None]

        index = [i[0] for i in data]
        temperature = [i[1] for i in data]
        wind_velosity = [i[5] for i in data]
        # Построение графиков
        plt.figure(figsize=(9, 9))
        plt.subplot(2, 1, 1)
        plt.bar(index, temperature, color='g')  # построение графика
        plt.title('Распределение температуры и скорости ветра')  # заголовок
        plt.ylabel('Температура, град.С', fontsize=14)  # ось ординат
        plt.subplot(2, 1, 2)
        plt.bar(index, wind_velosity, color='b')  # построение графика
        plt.xlabel('Месяцы', fontsize=14)  # ось абсцисс
        plt.ylabel('Скорость ветра, м/с', fontsize=14)  # ось ординат
        plt.show()


if __name__ == '__main__':
    r = Weather_calc().calc_statistics()
