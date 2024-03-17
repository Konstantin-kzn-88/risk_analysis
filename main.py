import xlwings as xw
from calculation_methods import calc_strait_fire, calc_tvs_explosion, calc_jet_fire, calc_lower_concentration, \
    calc_fireball

try:
    wb = xw.books.active
    ws = wb.sheets['Расчет']

except FileNotFoundError:
    print('Файл не отрыт, либо не листа "Сценарии"')

# сгенерируем все индексы в которых могут находится деревья
all_index_of_type_tree = [i for i in range(7, 1000, 10)]

# цикл расчета всех деревьев
for index in all_index_of_type_tree:
    # проверка что всё оборудование посчитано
    if xw.Range(f'L{index}').value == None:
        print("Расчет окончен")
        break

    # Получили исходные данные для расчета
    get_data_for_calc = {
        'Площадь, м2': xw.Range(f'L{index - 5}').value,
        'Расход газ, кг/с': xw.Range(f'L{index - 4}').value,
        'Расход жидкость, кг/с': xw.Range(f'L{index - 3}').value,
        'Теплота сгорания, кДж/кг': xw.Range(f'L{index - 2}').value,
        'НКПР, об.%': xw.Range(f'L{index - 1}').value,
        'Тип дерева': xw.Range(f'L{index}').value,
    }

    if get_data_for_calc['Тип дерева'] == 12:
        # ПОЖАР
        # 1. Получить экземпляр класса пожара
        fire_unit = calc_strait_fire.Strait_fire()
        # 2. Получить зоны классифицированные
        zone_array = fire_unit.termal_class_zone(get_data_for_calc['Площадь, м2'], 0.06, 100, 20, 1)
        # 3. Запишем решение
        xw.Range(f'P{index - 5}').value = zone_array[0]
        xw.Range(f'Q{index - 5}').value = zone_array[1]
        xw.Range(f'R{index - 5}').value = zone_array[2]
        xw.Range(f'S{index - 5}').value = zone_array[3]

        # ВЗРЫВ
        # 1. Получим класс для расчета взрыва
        explosion_unit = calc_tvs_explosion.Explosion()
        # 2. Получим зоны классифицированные
        zone_cls_array = explosion_unit.explosion_class_zone(3, 3, xw.Range(f'J{index - 4}').value * 1000, 45320, 7, 1)
        # 3. Запишем решение
        xw.Range(f'T{index - 4}').value = zone_cls_array[2]
        xw.Range(f'U{index - 4}').value = zone_cls_array[3]
        xw.Range(f'V{index - 4}').value = zone_cls_array[4]
        xw.Range(f'W{index - 4}').value = zone_cls_array[5]

        # ФАКЕЛ
        # 1. Получим класс для расчета
        jetfire_unit = calc_jet_fire.Torch()
        # 2а. Получим зоны жидкостной факел
        lenght, width = jetfire_unit.jetfire_size(get_data_for_calc['Расход жидкость, кг/с'], 2)
        # 3. Запишем решение
        xw.Range(f'X{index - 2}').value = lenght
        xw.Range(f'Y{index - 2}').value = width
        # 2б. Получим зоны газовый факел
        lenght, width = jetfire_unit.jetfire_size(get_data_for_calc['Расход газ, кг/с'], 0)
        # 3. Запишем решение
        xw.Range(f'X{index}').value = lenght
        xw.Range(f'Y{index}').value = width

        # ПОЖАР-ВСПЫШКА
        # 1. Получим класс для расчета
        lclp_unit = calc_lower_concentration.LCLP()
        # 2. Получим зоны
        zone_lclp = lclp_unit.lower_concentration_limit(xw.Range(f'J{index + 1}').value * 1000, 100, 30,
                                                        get_data_for_calc['НКПР, об.%'])
        # 3. Запишем решение
        xw.Range(f'Z{index + 1}').value = zone_lclp[0]
        xw.Range(f'AA{index + 1}').value = zone_lclp[1]

        # ОГНЕННЫЙ ШАР
        # 1. Получим класс для расчета шара
        fireball_unit = calc_fireball.Fireball()
        # 2. Получим зоны классифицированные
        zone_cls_array = fireball_unit.termal_class_zone(xw.Range(f'J{index + 3}').value * 1000, 350)

        # 3. Запишем решение
        xw.Range(f'AD{index + 3}').value = zone_cls_array[0]
        xw.Range(f'AE{index + 3}').value = zone_cls_array[1]
        xw.Range(f'AF{index + 3}').value = zone_cls_array[2]
        xw.Range(f'AG{index + 3}').value = zone_cls_array[3]

    if get_data_for_calc['Тип дерева'] == 14:
        # ПОЖАР
        # 1. Получить экземпляр класса пожара
        fire_unit = calc_strait_fire.Strait_fire()
        # 2. Получить зоны классифицированные
        zone_array = fire_unit.termal_class_zone(get_data_for_calc['Площадь, м2'], 0.06, 100, 20, 1)
        # 3. Запишем решение
        xw.Range(f'P{index - 5}').value = zone_array[0]
        xw.Range(f'Q{index - 5}').value = zone_array[1]
        xw.Range(f'R{index - 5}').value = zone_array[2]
        xw.Range(f'S{index - 5}').value = zone_array[3]

        # ВЗРЫВ
        # 1. Получим класс для расчета взрыва
        explosion_unit = calc_tvs_explosion.Explosion()
        # 2. Получим зоны классифицированные
        zone_cls_array = explosion_unit.explosion_class_zone(3, 3, xw.Range(f'J{index - 4}').value * 1000, 45320, 7, 1)
        # 3. Запишем решение
        xw.Range(f'T{index - 4}').value = zone_cls_array[2]
        xw.Range(f'U{index - 4}').value = zone_cls_array[3]
        xw.Range(f'V{index - 4}').value = zone_cls_array[4]
        xw.Range(f'W{index - 4}').value = zone_cls_array[5]

    if get_data_for_calc['Тип дерева'] == 11:
        # ФАКЕЛ
        # 1. Получим класс для расчета
        jetfire_unit = calc_jet_fire.Torch()
        # 2а. Получим зоны жидкостной факел
        lenght, width = jetfire_unit.jetfire_size(get_data_for_calc['Расход жидкость, кг/с'], 2)
        # 3. Запишем решение
        xw.Range(f'X{index - 5}').value = lenght
        xw.Range(f'Y{index - 5}').value = width

        # ВЗРЫВ
        # 1. Получим класс для расчета взрыва
        explosion_unit = calc_tvs_explosion.Explosion()
        # 2. Получим зоны классифицированные
        zone_cls_array = explosion_unit.explosion_class_zone(3, 3, xw.Range(f'J{index - 4}').value * 1000, 45320, 7, 1)
        # 3. Запишем решение
        xw.Range(f'T{index - 4}').value = zone_cls_array[2]
        xw.Range(f'U{index - 4}').value = zone_cls_array[3]
        xw.Range(f'V{index - 4}').value = zone_cls_array[4]
        xw.Range(f'W{index - 4}').value = zone_cls_array[5]

if __name__ == '__main__':
    print('Done')
