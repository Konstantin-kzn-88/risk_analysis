import xlwings as xw
import pprint

pp = pprint.PrettyPrinter(indent=4)

class Risk:

    def __init__(self):
        try:
            self.wb = xw.books.active
            self.ws = self.wb.sheets['Расчет']

        except FileNotFoundError:
            print('Файл не отрыт, либо не листа "Сценарии"')


    def get_sum_risk(self, index):
        sum_risk = {
            'Коллективный риск гибели': sum(
                cell.value for cell in self.ws.range(f'AW{index - 6}:AW{index + 2}') if cell.value is not None),
            'Коллективный риск ранения': sum(
                cell.value for cell in self.ws.range(f'AX{index - 6}:AX{index + 2}') if cell.value is not None),
            'Индивидуальный риск гибели': sum(
                cell.value for cell in self.ws.range(f'AZ{index - 6}:AZ{index + 2}') if cell.value is not None),
            'Индивидуальный риск ранения': sum(
                cell.value for cell in self.ws.range(f'BA{index - 6}:BA{index + 2}') if cell.value is not None),
            'Максимальный суммарный ущерб': max(
                cell.value for cell in self.ws.range(f'AV{index - 6}:AV{index + 2}') if cell.value is not None),
            'Максимальный экологический ущерб': max(
                cell.value for cell in self.ws.range(f'AU{index - 6}:AU{index + 2}') if cell.value is not None),
        }
        return sum_risk


    def risk_result(self):
        # сгенерируем все индексы в которых могут находится составляющие
        all_index_of_component = [i for i in range(8, 1000, 10)]
        components = {}
        #
        # цикл расчета всех деревьев
        for index in all_index_of_component:
            # проверка что всё оборудование посчитано
            if xw.Range(f'L{index}').value == None:
                break

            if xw.Range(f'L{index}').value not in components:
                components[xw.Range(f'L{index}').value] = self.get_sum_risk(index)
            else:
                risk = self.get_sum_risk(index)


                components[xw.Range(f'L{index}').value]['Коллективный риск гибели'] = components[xw.Range(f'L{index}').value][
                                                                                          'Коллективный риск гибели'] + risk[
                                                                                          'Коллективный риск гибели']
                components[xw.Range(f'L{index}').value]['Коллективный риск ранения'] = components[xw.Range(f'L{index}').value][
                                                                                           'Коллективный риск ранения'] + risk[
                                                                                           'Коллективный риск ранения']
                components[xw.Range(f'L{index}').value]['Индивидуальный риск гибели'] = components[xw.Range(f'L{index}').value][
                                                                                            'Индивидуальный риск гибели'] + \
                                                                                        risk[
                                                                                            'Индивидуальный риск гибели']
                components[xw.Range(f'L{index}').value]['Индивидуальный риск гибели'] = components[xw.Range(f'L{index}').value][
                                                                                            'Индивидуальный риск гибели'] + \
                                                                                        risk[
                                                                                            'Индивидуальный риск гибели']
                components[xw.Range(f'L{index}').value]['Максимальный суммарный ущерб'] = max(
                    components[xw.Range(f'L{index}').value]['Максимальный суммарный ущерб'],
                    risk['Максимальный суммарный ущерб'])

                components[xw.Range(f'L{index}').value]['Максимальный экологический ущерб'] = max(
                    components[xw.Range(f'L{index}').value]['Максимальный экологический ущерб'],
                    risk['Максимальный экологический ущерб'])


        # pp.pprint(components)

        return components


if __name__ == '__main__':
    components= Risk().risk_result()
    pp.pprint(components)