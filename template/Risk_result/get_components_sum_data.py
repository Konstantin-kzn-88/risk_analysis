import xlwings as xw
import pprint

pp = pprint.PrettyPrinter(indent=4)


class Components_sum_data:


    def __init__(self):
        try:
            self.wb = xw.books.active
            self.ws = self.wb.sheets['Расчет']

        except FileNotFoundError:
            print('Файл не отрыт, либо не листа "Сценарии"')




    def get_more_danger(self, index):
        scenario = {
            'Оборудование': xw.Range(f'B{index-6}').value,
            'Сценарий': xw.Range(f'C{index-6}').value,
            'Номер сценария': xw.Range(f'A{index - 6}').value,
            'Количество в аварии': xw.Range(f'I{index-6}').value,
            'Вероятность': xw.Range(f'H{index-6}').value,
            'Погибшие': xw.Range(f'AJ{index-6}').value,
            'Пострадавшие': xw.Range(f'AK{index-6}').value,
            'Ущерб экол.': xw.Range(f'AU{index-6}').value,
            'Ущерб': xw.Range(f'AV{index-6}').value,
        }

        for i in range(index-6, index+2):
            if xw.Range(f'H{i}').value is None:
                break
            # print(xw.Range(f'B{i}').value, xw.Range(f'AJ{i}').value)
            if int(xw.Range(f'AJ{i}').value) > scenario['Погибшие']:
                scenario = {
                    'Оборудование': xw.Range(f'B{i}').value,
                    'Сценарий': xw.Range(f'C{i}').value,
                    'Номер сценария': xw.Range(f'A{i}').value,
                    'Количество в аварии': xw.Range(f'I{i}').value,
                    'Вероятность': xw.Range(f'H{i}').value,
                    'Погибшие': xw.Range(f'AJ{i}').value,
                    'Пострадавшие': xw.Range(f'AK{i}').value,
                    'Ущерб экол.': xw.Range(f'AU{i}').value,
                    'Ущерб': xw.Range(f'AV{i}').value,
                }
            else:
                continue
        return scenario


    def get_more_probability(self, index):
        scenario = {
            'Оборудование': xw.Range(f'B{index-6}').value,
            'Сценарий': xw.Range(f'C{index-6}').value,
            'Номер сценария': xw.Range(f'A{index-6}').value,
            'Количество в аварии': xw.Range(f'I{index-6}').value,
            'Вероятность': xw.Range(f'H{index-6}').value,
            'Погибшие': xw.Range(f'AJ{index-6}').value,
            'Пострадавшие': xw.Range(f'AK{index-6}').value,
            'Ущерб экол.': xw.Range(f'AU{index-6}').value,
            'Ущерб': xw.Range(f'AV{index-6}').value,
        }
        for i in range(index-6, index+2):
            if xw.Range(f'H{i}').value is None:
                break
            # print(xw.Range(f'B{i}').value, xw.Range(f'H{i}').value)

            if float(xw.Range(f'H{i}').value) > scenario['Вероятность']:
                scenario = {
                    'Оборудование': xw.Range(f'B{i}').value,
                    'Сценарий': xw.Range(f'C{i}').value,
                    'Номер сценария': xw.Range(f'A{i}').value,
                    'Количество в аварии': xw.Range(f'I{i}').value,
                    'Вероятность': xw.Range(f'H{i}').value,
                    'Погибшие': xw.Range(f'AJ{i}').value,
                    'Пострадавшие': xw.Range(f'AK{i}').value,
                    'Ущерб экол.': xw.Range(f'AU{i}').value,
                    'Ущерб': xw.Range(f'AV{i}').value,
                }
            else:
                continue
        return scenario

    def components_result(self):

        # сгенерируем все индексы в которых могут находится составляющие
        all_index_of_component = [i for i in range(8, 1000, 10)]


        components = {}
        #
        # цикл расчета всех деревьев
        for index in all_index_of_component:
            # проверка что всё оборудование посчитано
            if xw.Range(f'L{index}').value == None:
                print("Расчет окончен", f'L{index}')
                break

        #     определяем наиболее опасный и наиболее вероятный в блоке сценарий
            if xw.Range(f'L{index}').value not in components:
                components[xw.Range(f'L{index}').value] = [self.get_more_danger(index), self.get_more_probability(index)]
            else:
                dead_man = self.get_more_danger(index)['Погибшие']
                probability = self.get_more_probability(index)['Вероятность']

                if dead_man > components[xw.Range(f'L{index}').value][0]['Погибшие']:
                    components[xw.Range(f'L{index}').value][0] = self.get_more_danger(index)

                if probability > components[xw.Range(f'L{index}').value][1]['Вероятность']:
                    components[xw.Range(f'L{index}').value][1] = self.get_more_probability(index)

        # pp.pprint(components)

        return components


if __name__ == '__main__':
    components= Components_sum_data().components_result()
    pp.pprint(components)

