import xlwings as xw

try:
    wb = xw.books.active
    ws = wb.sheets['Расчет']

except FileNotFoundError:
    print('Файл не отрыт, либо не листа "Сценарии"')

# сгенерируем все индексы в которых могут находится сценарии
all_index_of_scenario = [i for i in range(2, 1000, 1)]

num = 1

for index in all_index_of_scenario:
    # проверка что всё оборудование посчитано
    if xw.Range(f'A{index}').value == None:
        continue
    else:
        xw.Range(f'A{index}').value = f'C' + str(num)
    num += 1
