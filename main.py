import xlwings as xw
from calculation_methods import calc_strait_fire, calc_tvs_explosion, calc_jet_fire, calc_lower_concentration, \
    calc_fireball

DEGREE_OF_CLUTTER = 3 #степень загроможденности


def toxi_fake(r: int):
    return (r * 365, r * 689)


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
        print("Расчет окончен", f'L{index}')
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
    # print(get_data_for_calc)
    if get_data_for_calc['Тип дерева'] == 16:
        # ФАКЕЛ
        # 1. Получим класс для расчета
        jetfire_unit = calc_jet_fire.Torch()
        # 2а. Получим зоны полный жидкостной факел
        lenght, width = jetfire_unit.jetfire_size(get_data_for_calc['Расход жидкость, кг/с'], 2)
        # 3. Запишем решение
        xw.Range(f'Y{index - 5}').value = lenght
        xw.Range(f'Z{index - 5}').value = width
        # ВЗРЫВ
        # 1. Получим класс для расчета взрыва
        explosion_unit = calc_tvs_explosion.Explosion()
        # 2. Получим зоны классифицированные
        zone_cls_array = explosion_unit.explosion_class_zone(3, DEGREE_OF_CLUTTER, xw.Range(f'J{index - 4}').value * 1000, 45320, 7, 2)
        # 3. Запишем решение
        xw.Range(f'T{index - 4}').value = zone_cls_array[1]
        xw.Range(f'U{index - 4}').value = zone_cls_array[2]
        xw.Range(f'V{index - 4}').value = zone_cls_array[3]
        xw.Range(f'W{index - 4}').value = zone_cls_array[4]
        xw.Range(f'X{index - 4}').value = zone_cls_array[5]

        # ПОЖАР
        # 1. Получить экземпляр класса пожара
        fire_unit = calc_strait_fire.Strait_fire()
        # 2. Получить зоны классифицированные
        zone_array = fire_unit.termal_class_zone(get_data_for_calc['Площадь, м2'], 0.06, 100, 20, 1)
        # 3. Запишем решение
        xw.Range(f'P{index - 2}').value = zone_array[0]
        xw.Range(f'Q{index - 2}').value = zone_array[1]
        xw.Range(f'R{index - 2}').value = zone_array[2]
        xw.Range(f'S{index - 2}').value = zone_array[3]

        # ПОЖАР-ВСПЫШКА
        # 1. Получим класс для расчета
        lclp_unit = calc_lower_concentration.LCLP()
        # 2. Получим зоны
        zone_lclp = lclp_unit.lower_concentration_limit(xw.Range(f'J{index - 1}').value * 1000, 100, 30,
                                                        get_data_for_calc['НКПР, об.%'])
        # 3. Запишем решение
        xw.Range(f'AA{index - 1}').value = zone_lclp[0]
        xw.Range(f'AB{index - 1}').value = zone_lclp[1]



    if get_data_for_calc['Тип дерева'] == 3:
        # ПОЖАР
        print("Hi")
        # 1. Получить экземпляр класса пожара
        fire_unit = calc_strait_fire.Strait_fire()
        # 2. Получить зоны классифицированные
        zone_array = fire_unit.termal_class_zone(get_data_for_calc['Площадь, м2'], 0.06, 100, 20, 1)
        zone_array_2 = fire_unit.termal_class_zone(get_data_for_calc['Площадь, м2'] / 10, 0.06, 100, 20, 1)
        # 3. Запишем решение
        xw.Range(f'P{index - 5}').value = zone_array[0]
        xw.Range(f'Q{index - 5}').value = zone_array[1]
        xw.Range(f'R{index - 5}').value = zone_array[2]
        xw.Range(f'S{index - 5}').value = zone_array[3]

        xw.Range(f'P{index - 4}').value = zone_array[0]
        xw.Range(f'Q{index - 4}').value = zone_array[1]
        xw.Range(f'R{index - 4}').value = zone_array[2]
        xw.Range(f'S{index - 4}').value = zone_array[3]

        xw.Range(f'P{index - 2}').value = zone_array_2[0]
        xw.Range(f'Q{index - 2}').value = zone_array_2[1]
        xw.Range(f'R{index - 2}').value = zone_array_2[2]
        xw.Range(f'S{index - 2}').value = zone_array_2[3]

        xw.Range(f'P{index - 1}').value = zone_array_2[0]
        xw.Range(f'Q{index - 1}').value = zone_array_2[1]
        xw.Range(f'R{index - 1}').value = zone_array_2[2]
        xw.Range(f'S{index - 1}').value = zone_array_2[3]

    if get_data_for_calc['Тип дерева'] == 1:
        # ПОЖАР
        # 1. Получить экземпляр класса пожара
        fire_unit = calc_strait_fire.Strait_fire()
        # 2. Получить зоны классифицированные
        zone_array = fire_unit.termal_class_zone(get_data_for_calc['Площадь, м2'], 0.06, 100, 20, 1)
        zone_array_2 = fire_unit.termal_class_zone(get_data_for_calc['Площадь, м2'] / 10, 0.06, 100, 20, 1)
        # 3. Запишем решение
        xw.Range(f'P{index - 5}').value = zone_array[0]
        xw.Range(f'Q{index - 5}').value = zone_array[1]
        xw.Range(f'R{index - 5}').value = zone_array[2]
        xw.Range(f'S{index - 5}').value = zone_array[3]

        xw.Range(f'P{index - 2}').value = zone_array_2[0]
        xw.Range(f'Q{index - 2}').value = zone_array_2[1]
        xw.Range(f'R{index - 2}').value = zone_array_2[2]
        xw.Range(f'S{index - 2}').value = zone_array_2[3]

        # ВЗРЫВ
        # 1. Получим класс для расчета взрыва
        explosion_unit = calc_tvs_explosion.Explosion()
        # 2. Получим зоны классифицированные
        zone_cls_array = explosion_unit.explosion_class_zone(3, DEGREE_OF_CLUTTER, xw.Range(f'J{index - 4}').value * 1000, 45320, 7, 2)
        # 3. Запишем решение
        xw.Range(f'T{index - 4}').value = zone_cls_array[1]
        xw.Range(f'U{index - 4}').value = zone_cls_array[2]
        xw.Range(f'V{index - 4}').value = zone_cls_array[3]
        xw.Range(f'W{index - 4}').value = zone_cls_array[4]
        xw.Range(f'X{index - 4}').value = zone_cls_array[5]

        # ПОЖАР-ВСПЫШКА
        # 1. Получим класс для расчета
        lclp_unit = calc_lower_concentration.LCLP()
        # 2. Получим зоны
        zone_lclp = lclp_unit.lower_concentration_limit(xw.Range(f'J{index - 1}').value * 1000, 100, 30,
                                                        get_data_for_calc['НКПР, об.%'])
        # 3. Запишем решение
        xw.Range(f'AA{index - 1}').value = zone_lclp[0]
        xw.Range(f'AB{index - 1}').value = zone_lclp[1]

    if get_data_for_calc['Тип дерева'] == 4:
        # ФАКЕЛ
        # 1. Получим класс для расчета
        jetfire_unit = calc_jet_fire.Torch()
        # 2а. Получим зоны полный газовый факел
        lenght, width = jetfire_unit.jetfire_size(get_data_for_calc['Расход газ, кг/с'], 1)
        # 3. Запишем решение
        xw.Range(f'Y{index - 5}').value = lenght
        xw.Range(f'Z{index - 5}').value = width
        # 2б. Получим зоны частичный газовый факел
        lenght, width = jetfire_unit.jetfire_size(get_data_for_calc['Расход газ, кг/с'] / 3, 1)
        # 3. Запишем решение
        xw.Range(f'Y{index - 1}').value = lenght
        xw.Range(f'Z{index - 1}').value = width

        # ВЗРЫВ
        # 1. Получим класс для расчета взрыва
        explosion_unit = calc_tvs_explosion.Explosion()
        # 2. Получим зоны классифицированные
        zone_cls_array = explosion_unit.explosion_class_zone(3, DEGREE_OF_CLUTTER, xw.Range(f'J{index - 4}').value * 1000, 45320, 7, 2)
        zone_cls_array_2 = explosion_unit.explosion_class_zone(3, DEGREE_OF_CLUTTER, xw.Range(f'J{index}').value * 1000, 45320, 7, 2)
        # 3. Запишем решение
        xw.Range(f'T{index - 4}').value = zone_cls_array[1]
        xw.Range(f'U{index - 4}').value = zone_cls_array[2]
        xw.Range(f'V{index - 4}').value = zone_cls_array[3]
        xw.Range(f'W{index - 4}').value = zone_cls_array[4]
        xw.Range(f'X{index - 4}').value = zone_cls_array[5]

        xw.Range(f'T{index}').value = zone_cls_array_2[1]
        xw.Range(f'U{index}').value = zone_cls_array_2[2]
        xw.Range(f'V{index}').value = zone_cls_array_2[3]
        xw.Range(f'W{index}').value = zone_cls_array_2[4]
        xw.Range(f'X{index}').value = zone_cls_array_2[5]

        # ПОЖАР-ВСПЫШКА
        # 1. Получим класс для расчета
        lclp_unit = calc_lower_concentration.LCLP()
        # 2. Получим зоны
        zone_lclp = lclp_unit.lower_concentration_limit(xw.Range(f'J{index - 3}').value * 1000, 100, 30,
                                                        get_data_for_calc['НКПР, об.%'])
        zone_lclp_2 = lclp_unit.lower_concentration_limit(xw.Range(f'J{index + 1}').value * 1000, 100, 30,
                                                          get_data_for_calc['НКПР, об.%'])
        # 3. Запишем решение
        xw.Range(f'AA{index - 3}').value = zone_lclp[0]
        xw.Range(f'AB{index - 3}').value = zone_lclp[1]

        xw.Range(f'AA{index +1}').value = zone_lclp_2[0]
        xw.Range(f'AB{index +1}').value = zone_lclp_2[1]

    if get_data_for_calc['Тип дерева'] == 19:
        # ФАКЕЛ
        # 1. Получим класс для расчета
        jetfire_unit = calc_jet_fire.Torch()
        # 2а. Получим зоны полный жидкостной факел
        lenght, width = jetfire_unit.jetfire_size(get_data_for_calc['Расход жидкость, кг/с'], 2)
        # 3. Запишем решение
        xw.Range(f'Y{index - 5}').value = lenght
        xw.Range(f'Z{index - 5}').value = width
        # 2б. Получим зоны частичный жидкостной факел
        lenght, width = jetfire_unit.jetfire_size(get_data_for_calc['Расход жидкость, кг/с'] / 3, 2)
        # 3. Запишем решение
        xw.Range(f'Y{index - 1}').value = lenght
        xw.Range(f'Z{index - 1}').value = width

        # ВЗРЫВ
        # 1. Получим класс для расчета взрыва
        explosion_unit = calc_tvs_explosion.Explosion()
        # 2. Получим зоны классифицированные
        zone_cls_array = explosion_unit.explosion_class_zone(3, DEGREE_OF_CLUTTER, xw.Range(f'J{index - 4}').value * 1000, 45320, 7, 2)
        zone_cls_array_2 = explosion_unit.explosion_class_zone(3, DEGREE_OF_CLUTTER, xw.Range(f'J{index}').value * 1000, 45320, 7, 2)
        # 3. Запишем решение
        xw.Range(f'T{index - 4}').value = zone_cls_array[1]
        xw.Range(f'U{index - 4}').value = zone_cls_array[2]
        xw.Range(f'V{index - 4}').value = zone_cls_array[3]
        xw.Range(f'W{index - 4}').value = zone_cls_array[4]
        xw.Range(f'X{index - 4}').value = zone_cls_array[5]

        xw.Range(f'T{index}').value = zone_cls_array_2[1]
        xw.Range(f'U{index}').value = zone_cls_array_2[2]
        xw.Range(f'V{index}').value = zone_cls_array_2[3]
        xw.Range(f'W{index}').value = zone_cls_array_2[4]
        xw.Range(f'X{index}').value = zone_cls_array_2[5]

        # ОГНЕННЫЙ ШАР
        # 1. Получим класс для расчета шара
        fireball_unit = calc_fireball.Fireball()
        # 2. Получим зоны классифицированные
        zone_cls_array = fireball_unit.termal_class_zone(xw.Range(f'J{index - 3}').value * 1000, 350)

        # 3. Запишем решение
        xw.Range(f'AE{index - 3}').value = zone_cls_array[0]
        xw.Range(f'AF{index - 3}').value = zone_cls_array[1]
        xw.Range(f'AG{index - 3}').value = zone_cls_array[2]
        xw.Range(f'AH{index - 3}').value = zone_cls_array[3]

        # ПОЖАР-ВСПЫШКА
        # 1. Получим класс для расчета
        lclp_unit = calc_lower_concentration.LCLP()
        # 2. Получим зоны
        zone_lclp = lclp_unit.lower_concentration_limit(xw.Range(f'J{index + 1}').value * 1000, 100, 30,
                                                        get_data_for_calc['НКПР, об.%'])
        # 3. Запишем решение
        xw.Range(f'AA{index + 1}').value = zone_lclp[0]
        xw.Range(f'AB{index + 1}').value = zone_lclp[1]

    if get_data_for_calc['Тип дерева'] == 2:
        # ПОЖАР
        # 1. Получить экземпляр класса пожара
        fire_unit = calc_strait_fire.Strait_fire()
        # 2. Получить зоны классифицированные
        zone_array = fire_unit.termal_class_zone(get_data_for_calc['Площадь, м2'], 0.06, 100, 20, 1)
        zone_array_2 = fire_unit.termal_class_zone(get_data_for_calc['Площадь, м2'] / 10, 0.06, 100, 20, 1)
        # 3. Запишем решение
        xw.Range(f'P{index - 5}').value = zone_array[0]
        xw.Range(f'Q{index - 5}').value = zone_array[1]
        xw.Range(f'R{index - 5}').value = zone_array[2]
        xw.Range(f'S{index - 5}').value = zone_array[3]

        xw.Range(f'P{index - 2}').value = zone_array_2[0]
        xw.Range(f'Q{index - 2}').value = zone_array_2[1]
        xw.Range(f'R{index - 2}').value = zone_array_2[2]
        xw.Range(f'S{index - 2}').value = zone_array_2[3]

        # ВЗРЫВ
        # 1. Получим класс для расчета взрыва
        explosion_unit = calc_tvs_explosion.Explosion()
        # 2. Получим зоны классифицированные
        zone_cls_array = explosion_unit.explosion_class_zone(3, DEGREE_OF_CLUTTER, xw.Range(f'J{index - 4}').value * 1000, 45320, 7, 2)
        # 3. Запишем решение
        xw.Range(f'T{index - 4}').value = zone_cls_array[1]
        xw.Range(f'U{index - 4}').value = zone_cls_array[2]
        xw.Range(f'V{index - 4}').value = zone_cls_array[3]
        xw.Range(f'W{index - 4}').value = zone_cls_array[4]
        xw.Range(f'X{index - 4}').value = zone_cls_array[5]

        # ПОЖАР-ВСПЫШКА
        # 1. Получим класс для расчета
        lclp_unit = calc_lower_concentration.LCLP()
        # 2. Получим зоны
        zone_lclp = lclp_unit.lower_concentration_limit(xw.Range(f'J{index - 1}').value * 1000, 100, 30,
                                                        get_data_for_calc['НКПР, об.%'])
        # 3. Запишем решение
        xw.Range(f'AA{index - 1}').value = zone_lclp[0]
        xw.Range(f'AB{index - 1}').value = zone_lclp[1]

        #  ТОКСИ
        xw.Range(f'AC{index - 3}').value = round(toxi_fake(xw.Range(f'J{index - 3}').value)[0],1)
        xw.Range(f'AD{index - 3}').value = round(toxi_fake(xw.Range(f'J{index - 3}').value)[1],1)

        xw.Range(f'AC{index}').value = round(toxi_fake(xw.Range(f'J{index}').value)[0],1)
        xw.Range(f'AD{index}').value = round(toxi_fake(xw.Range(f'J{index}').value)[1],1)

    if get_data_for_calc['Тип дерева'] == 13:
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
        zone_cls_array = explosion_unit.explosion_class_zone(3, DEGREE_OF_CLUTTER, xw.Range(f'J{index - 4}').value * 1000, 45320, 7, 2)
        # 3. Запишем решение
        xw.Range(f'T{index - 4}').value = zone_cls_array[1]
        xw.Range(f'U{index - 4}').value = zone_cls_array[2]
        xw.Range(f'V{index - 4}').value = zone_cls_array[3]
        xw.Range(f'W{index - 4}').value = zone_cls_array[4]
        xw.Range(f'X{index - 4}').value = zone_cls_array[5]

        # ТОКСИ
        xw.Range(f'AC{index - 3}').value = round(toxi_fake(xw.Range(f'J{index - 3}').value)[0], 1)
        xw.Range(f'AD{index - 3}').value = round(toxi_fake(xw.Range(f'J{index - 3}').value)[1], 1)

        xw.Range(f'AC{index - 1}').value = round(toxi_fake(xw.Range(f'J{index - 1}').value)[0], 1)
        xw.Range(f'AD{index - 1}').value = round(toxi_fake(xw.Range(f'J{index - 1}').value)[1], 1)

        xw.Range(f'AC{index + 2}').value = round(toxi_fake(xw.Range(f'J{index + 2}').value)[0], 1)
        xw.Range(f'AD{index + 2}').value = round(toxi_fake(xw.Range(f'J{index + 2}').value)[1], 1)

        # ФАКЕЛ
        # 1. Получим класс для расчета
        jetfire_unit = calc_jet_fire.Torch()
        # 2а. Получим зоны полный жидкостной факел
        lenght, width = jetfire_unit.jetfire_size(get_data_for_calc['Расход жидкость, кг/с'], 2)
        # 3. Запишем решение
        xw.Range(f'Y{index - 2}').value = lenght
        xw.Range(f'Z{index - 2}').value = width
        # 2б. Получим зоны частичный жидкостной факел
        lenght, width = jetfire_unit.jetfire_size(get_data_for_calc['Расход газ, кг/с'] / 3, 1)
        # 3. Запишем решение
        xw.Range(f'Y{index}').value = lenght
        xw.Range(f'Z{index}').value = width

        # ПОЖАР-ВСПЫШКА
        # 1. Получим класс для расчета
        lclp_unit = calc_lower_concentration.LCLP()
        # 2. Получим зоны
        zone_lclp = lclp_unit.lower_concentration_limit(xw.Range(f'J{index + 1}').value * 1000, 100, 30,
                                                        get_data_for_calc['НКПР, об.%'])
        # 3. Запишем решение
        xw.Range(f'AA{index + 1}').value = zone_lclp[0]
        xw.Range(f'AB{index + 1}').value = zone_lclp[1]

        # ОГНЕННЫЙ ШАР+ПОЖАР
        # 1. Получим класс для расчета шара
        fireball_unit = calc_fireball.Fireball()
        # 2. Получим зоны классифицированные
        zone_cls_array = fireball_unit.termal_class_zone(xw.Range(f'J{index + 3}').value * 1000, 350)

        # 3. Запишем решение
        xw.Range(f'AE{index + 3}').value = zone_cls_array[0]
        xw.Range(f'AF{index + 3}').value = zone_cls_array[1]
        xw.Range(f'AG{index + 3}').value = zone_cls_array[2]
        xw.Range(f'AH{index + 3}').value = zone_cls_array[3]

        # 1. Получить экземпляр класса пожара
        fire_unit = calc_strait_fire.Strait_fire()
        # 2. Получить зоны классифицированные
        zone_array = fire_unit.termal_class_zone(get_data_for_calc['Площадь, м2'], 0.06, 100, 20, 1)
        # 3. Запишем решение
        xw.Range(f'P{index + 3}').value = zone_array[0]
        xw.Range(f'Q{index + 3}').value = zone_array[1]
        xw.Range(f'R{index + 3}').value = zone_array[2]
        xw.Range(f'S{index + 3}').value = zone_array[3]



    if get_data_for_calc['Тип дерева'] == 22:
        # ОГНЕННЫЙ ШАР
        # 1. Получим класс для расчета шара
        fireball_unit = calc_fireball.Fireball()
        # 2. Получим зоны классифицированные
        zone_cls_array = fireball_unit.termal_class_zone(xw.Range(f'J{index - 5}').value * 1000, 350)

        # 3. Запишем решение
        xw.Range(f'AE{index - 5}').value = zone_cls_array[0]
        xw.Range(f'AF{index - 5}').value = zone_cls_array[1]
        xw.Range(f'AG{index - 5}').value = zone_cls_array[2]
        xw.Range(f'AH{index - 5}').value = zone_cls_array[3]

        # ВЗРЫВ
        # 1. Получим класс для расчета взрыва
        explosion_unit = calc_tvs_explosion.Explosion()
        # 2. Получим зоны классифицированные
        zone_cls_array = explosion_unit.explosion_class_zone(3, DEGREE_OF_CLUTTER, xw.Range(f'J{index - 4}').value * 1000, 45320, 7, 2)
        # 3. Запишем решение
        xw.Range(f'T{index - 4}').value = zone_cls_array[1]
        xw.Range(f'U{index - 4}').value = zone_cls_array[2]
        xw.Range(f'V{index - 4}').value = zone_cls_array[3]
        xw.Range(f'W{index - 4}').value = zone_cls_array[4]
        xw.Range(f'X{index - 4}').value = zone_cls_array[5]

        # ТОКСИ
        xw.Range(f'AC{index - 3}').value = round(toxi_fake(xw.Range(f'J{index - 3}').value)[0], 1)
        xw.Range(f'AD{index - 3}').value = round(toxi_fake(xw.Range(f'J{index - 3}').value)[1], 1)

        xw.Range(f'AC{index - 1}').value = round(toxi_fake(xw.Range(f'J{index - 1}').value)[0], 1)
        xw.Range(f'AD{index - 1}').value = round(toxi_fake(xw.Range(f'J{index - 1}').value)[1], 1)

        xw.Range(f'AC{index + 2}').value = round(toxi_fake(xw.Range(f'J{index + 2}').value)[0], 1)
        xw.Range(f'AD{index + 2}').value = round(toxi_fake(xw.Range(f'J{index + 2}').value)[1], 1)

        # ФАКЕЛ
        # 1. Получим класс для расчета
        jetfire_unit = calc_jet_fire.Torch()
        # 2а. Получим зоны полный жидкостной факел
        lenght, width = jetfire_unit.jetfire_size(get_data_for_calc['Расход жидкость, кг/с'], 2)
        # 3. Запишем решение
        xw.Range(f'Y{index - 2}').value = lenght
        xw.Range(f'Z{index - 2}').value = width
        # 2б. Получим зоны частичный жидкостной факел
        lenght, width = jetfire_unit.jetfire_size(get_data_for_calc['Расход газ, кг/с'] / 3, 1)
        # 3. Запишем решение
        xw.Range(f'Y{index}').value = lenght
        xw.Range(f'Z{index}').value = width

        # ПОЖАР-ВСПЫШКА
        # 1. Получим класс для расчета
        lclp_unit = calc_lower_concentration.LCLP()
        # 2. Получим зоны
        zone_lclp = lclp_unit.lower_concentration_limit(xw.Range(f'J{index + 1}').value * 1000, 100, 30,
                                                        get_data_for_calc['НКПР, об.%'])
        # 3. Запишем решение
        xw.Range(f'AA{index + 1}').value = zone_lclp[0]
        xw.Range(f'AB{index + 1}').value = zone_lclp[1]

        # ОГНЕННЫЙ ШАР+ПОЖАР
        # 1. Получим класс для расчета шара
        fireball_unit = calc_fireball.Fireball()
        # 2. Получим зоны классифицированные
        zone_cls_array = fireball_unit.termal_class_zone(xw.Range(f'J{index + 3}').value * 1000, 350)

        # 3. Запишем решение
        xw.Range(f'AE{index + 3}').value = zone_cls_array[0]
        xw.Range(f'AF{index + 3}').value = zone_cls_array[1]
        xw.Range(f'AG{index + 3}').value = zone_cls_array[2]
        xw.Range(f'AH{index + 3}').value = zone_cls_array[3]

        # 1. Получить экземпляр класса пожара
        fire_unit = calc_strait_fire.Strait_fire()
        # 2. Получить зоны классифицированные
        zone_array = fire_unit.termal_class_zone(get_data_for_calc['Площадь, м2'], 0.06, 100, 20, 1)
        # 3. Запишем решение
        xw.Range(f'P{index + 3}').value = zone_array[0]
        xw.Range(f'Q{index + 3}').value = zone_array[1]
        xw.Range(f'R{index + 3}').value = zone_array[2]
        xw.Range(f'S{index + 3}').value = zone_array[3]


    if get_data_for_calc['Тип дерева'] == 8:
        # ПОЖАР
        # 1. Получить экземпляр класса пожара
        fire_unit = calc_strait_fire.Strait_fire()
        # 2. Получить зоны классифицированные
        zone_array = fire_unit.termal_class_zone(get_data_for_calc['Площадь, м2'], 0.06, 100, 20, 1)
        zone_array_2 = fire_unit.termal_class_zone(get_data_for_calc['Площадь, м2'] / 10, 0.06, 100, 20, 1)
        # 3. Запишем решение
        xw.Range(f'P{index - 5}').value = zone_array[0]
        xw.Range(f'Q{index - 5}').value = zone_array[1]
        xw.Range(f'R{index - 5}').value = zone_array[2]
        xw.Range(f'S{index - 5}').value = zone_array[3]

        xw.Range(f'P{index - 4}').value = zone_array[0]
        xw.Range(f'Q{index - 4}').value = zone_array[1]
        xw.Range(f'R{index - 4}').value = zone_array[2]
        xw.Range(f'S{index - 4}').value = zone_array[3]

        xw.Range(f'P{index - 2}').value = zone_array_2[0]
        xw.Range(f'Q{index - 2}').value = zone_array_2[1]
        xw.Range(f'R{index - 2}').value = zone_array_2[2]
        xw.Range(f'S{index - 2}').value = zone_array_2[3]

        xw.Range(f'P{index - 1}').value = zone_array_2[0]
        xw.Range(f'Q{index - 1}').value = zone_array_2[1]
        xw.Range(f'R{index - 1}').value = zone_array_2[2]
        xw.Range(f'S{index - 1}').value = zone_array_2[3]


    if get_data_for_calc['Тип дерева'] == 12:
        # ПОЖАР
        # 1. Получить экземпляр класса пожара
        fire_unit = calc_strait_fire.Strait_fire()
        # 2. Получить зоны классифицированные
        zone_array = fire_unit.termal_class_zone(get_data_for_calc['Площадь, м2'], 0.06, 100, 20, 1)
        # 3. Запишем решение
        # ПОЖАР
        xw.Range(f'P{index - 5}').value = zone_array[0]
        xw.Range(f'Q{index - 5}').value = zone_array[1]
        xw.Range(f'R{index - 5}').value = zone_array[2]
        xw.Range(f'S{index - 5}').value = zone_array[3]

        # ВЗРЫВ
        # 1. Получим класс для расчета взрыва
        explosion_unit = calc_tvs_explosion.Explosion()
        # 2. Получим зоны классифицированные
        zone_cls_array = explosion_unit.explosion_class_zone(3, DEGREE_OF_CLUTTER, xw.Range(f'J{index - 4}').value * 1000, 45320, 7, 2)
        # 3. Запишем решение
        xw.Range(f'T{index - 4}').value = zone_cls_array[1]
        xw.Range(f'U{index - 4}').value = zone_cls_array[2]
        xw.Range(f'V{index - 4}').value = zone_cls_array[3]
        xw.Range(f'W{index - 4}').value = zone_cls_array[4]
        xw.Range(f'X{index - 4}').value = zone_cls_array[5]

        # ФАКЕЛ
        # 1. Получим класс для расчета
        jetfire_unit = calc_jet_fire.Torch()
        # 2а. Получим зоны полный жидкостной факел
        lenght, width = jetfire_unit.jetfire_size(get_data_for_calc['Расход жидкость, кг/с'], 2)
        # 3. Запишем решение
        xw.Range(f'Y{index - 2}').value = lenght
        xw.Range(f'Z{index - 2}').value = width
        # 2б. Получим зоны частичный жидкостной факел
        lenght, width = jetfire_unit.jetfire_size(get_data_for_calc['Расход газ, кг/с'] / 3, 1)
        # 3. Запишем решение
        xw.Range(f'Y{index}').value = lenght
        xw.Range(f'Z{index}').value = width

        # ПОЖАР-ВСПЫШКА
        # 1. Получим класс для расчета
        lclp_unit = calc_lower_concentration.LCLP()
        # 2. Получим зоны
        zone_lclp = lclp_unit.lower_concentration_limit(xw.Range(f'J{index + 1}').value * 1000, 100, 30,
                                                        get_data_for_calc['НКПР, об.%'])
        # 3. Запишем решение
        xw.Range(f'AA{index + 1}').value = zone_lclp[0]
        xw.Range(f'AB{index + 1}').value = zone_lclp[1]

        # ОГНЕННЫЙ ШАР+ПОЖАР
        # 1. Получим класс для расчета шара
        fireball_unit = calc_fireball.Fireball()
        # 2. Получим зоны классифицированные
        zone_cls_array = fireball_unit.termal_class_zone(xw.Range(f'J{index + 3}').value * 1000, 350)

        # 3. Запишем решение
        xw.Range(f'AE{index + 3}').value = zone_cls_array[0]
        xw.Range(f'AF{index + 3}').value = zone_cls_array[1]
        xw.Range(f'AG{index + 3}').value = zone_cls_array[2]
        xw.Range(f'AH{index + 3}').value = zone_cls_array[3]

        # 1. Получить экземпляр класса пожара
        fire_unit = calc_strait_fire.Strait_fire()
        # 2. Получить зоны классифицированные
        zone_array = fire_unit.termal_class_zone(get_data_for_calc['Площадь, м2'], 0.06, 100, 20, 1)
        # 3. Запишем решение
        xw.Range(f'P{index + 3}').value = zone_array[0]
        xw.Range(f'Q{index + 3}').value = zone_array[1]
        xw.Range(f'R{index + 3}').value = zone_array[2]
        xw.Range(f'S{index + 3}').value = zone_array[3]



    if get_data_for_calc['Тип дерева'] == 21:
        # ОГНЕННЫЙ ШАР
        # 1. Получим класс для расчета шара
        fireball_unit = calc_fireball.Fireball()
        # 2. Получим зоны классифицированные
        zone_cls_array = fireball_unit.termal_class_zone(xw.Range(f'J{index - 5}').value * 1000, 350)

        # 3. Запишем решение
        xw.Range(f'AE{index - 5}').value = zone_cls_array[0]
        xw.Range(f'AF{index - 5}').value = zone_cls_array[1]
        xw.Range(f'AG{index - 5}').value = zone_cls_array[2]
        xw.Range(f'AH{index - 5}').value = zone_cls_array[3]

        # ВЗРЫВ
        # 1. Получим класс для расчета взрыва
        explosion_unit = calc_tvs_explosion.Explosion()
        # 2. Получим зоны классифицированные
        zone_cls_array = explosion_unit.explosion_class_zone(3, DEGREE_OF_CLUTTER, xw.Range(f'J{index - 4}').value * 1000, 45320, 7, 2)
        # 3. Запишем решение
        xw.Range(f'T{index - 4}').value = zone_cls_array[1]
        xw.Range(f'U{index - 4}').value = zone_cls_array[2]
        xw.Range(f'V{index - 4}').value = zone_cls_array[3]
        xw.Range(f'W{index - 4}').value = zone_cls_array[4]
        xw.Range(f'X{index - 4}').value = zone_cls_array[5]

        # ФАКЕЛ
        # 1. Получим класс для расчета
        jetfire_unit = calc_jet_fire.Torch()
        # 2а. Получим зоны полный жидкостной факел
        lenght, width = jetfire_unit.jetfire_size(get_data_for_calc['Расход жидкость, кг/с'], 2)
        # 3. Запишем решение
        xw.Range(f'Y{index - 2}').value = lenght
        xw.Range(f'Z{index - 2}').value = width
        # 2б. Получим зоны частичный жидкостной факел
        lenght, width = jetfire_unit.jetfire_size(get_data_for_calc['Расход газ, кг/с'] / 3, 1)
        # 3. Запишем решение
        xw.Range(f'Y{index}').value = lenght
        xw.Range(f'Z{index}').value = width

        # ПОЖАР-ВСПЫШКА
        # 1. Получим класс для расчета
        lclp_unit = calc_lower_concentration.LCLP()
        # 2. Получим зоны
        zone_lclp = lclp_unit.lower_concentration_limit(xw.Range(f'J{index + 1}').value * 1000, 100, 30,
                                                        get_data_for_calc['НКПР, об.%'])
        # 3. Запишем решение
        xw.Range(f'AA{index + 1}').value = zone_lclp[0]
        xw.Range(f'AB{index + 1}').value = zone_lclp[1]

        # ОГНЕННЫЙ ШАР+ПОЖАР
        # 1. Получим класс для расчета шара
        fireball_unit = calc_fireball.Fireball()
        # 2. Получим зоны классифицированные
        zone_cls_array = fireball_unit.termal_class_zone(xw.Range(f'J{index + 3}').value * 1000, 350)

        # 3. Запишем решение
        xw.Range(f'AE{index + 3}').value = zone_cls_array[0]
        xw.Range(f'AF{index + 3}').value = zone_cls_array[1]
        xw.Range(f'AG{index + 3}').value = zone_cls_array[2]
        xw.Range(f'AH{index + 3}').value = zone_cls_array[3]

        # 1. Получить экземпляр класса пожара
        fire_unit = calc_strait_fire.Strait_fire()
        # 2. Получить зоны классифицированные
        zone_array = fire_unit.termal_class_zone(get_data_for_calc['Площадь, м2'], 0.06, 100, 20, 1)
        # 3. Запишем решение
        xw.Range(f'P{index + 3}').value = zone_array[0]
        xw.Range(f'Q{index + 3}').value = zone_array[1]
        xw.Range(f'R{index + 3}').value = zone_array[2]
        xw.Range(f'S{index + 3}').value = zone_array[3]

    if get_data_for_calc['Тип дерева'] == 5:
        # ФАКЕЛ
        # 1. Получим класс для расчета
        jetfire_unit = calc_jet_fire.Torch()
        # 2а. Получим зоны полный газовый факел
        lenght, width = jetfire_unit.jetfire_size(get_data_for_calc['Расход газ, кг/с'], 1)
        # 3. Запишем решение
        xw.Range(f'Y{index - 5}').value = lenght
        xw.Range(f'Z{index - 5}').value = width
        # 2б. Получим зоны частичный газовый факел
        lenght, width = jetfire_unit.jetfire_size(get_data_for_calc['Расход газ, кг/с'] / 3, 1)
        # 3. Запишем решение
        xw.Range(f'Y{index - 1}').value = lenght
        xw.Range(f'Z{index - 1}').value = width

        # ВЗРЫВ
        # 1. Получим класс для расчета взрыва
        explosion_unit = calc_tvs_explosion.Explosion()
        # 2. Получим зоны классифицированные
        zone_cls_array = explosion_unit.explosion_class_zone(3, DEGREE_OF_CLUTTER, xw.Range(f'J{index - 4}').value * 1000, 45320, 7, 2)
        zone_cls_array_2 = explosion_unit.explosion_class_zone(3, DEGREE_OF_CLUTTER, xw.Range(f'J{index}').value * 1000, 45320, 7, 2)
        # 3. Запишем решение
        xw.Range(f'T{index - 4}').value = zone_cls_array[1]
        xw.Range(f'U{index - 4}').value = zone_cls_array[2]
        xw.Range(f'V{index - 4}').value = zone_cls_array[3]
        xw.Range(f'W{index - 4}').value = zone_cls_array[4]
        xw.Range(f'X{index - 4}').value = zone_cls_array[5]

        xw.Range(f'T{index}').value = zone_cls_array_2[1]
        xw.Range(f'U{index}').value = zone_cls_array_2[2]
        xw.Range(f'V{index}').value = zone_cls_array_2[3]
        xw.Range(f'W{index}').value = zone_cls_array_2[4]
        xw.Range(f'X{index}').value = zone_cls_array_2[5]

        # ПОЖАР-ВСПЫШКА
        # 1. Получим класс для расчета
        lclp_unit = calc_lower_concentration.LCLP()
        # 2. Получим зоны
        zone_lclp = lclp_unit.lower_concentration_limit(xw.Range(f'J{index - 3}').value * 1000, 100, 30,
                                                        get_data_for_calc['НКПР, об.%'])
        zone_lclp_2 = lclp_unit.lower_concentration_limit(xw.Range(f'J{index + 1}').value * 1000, 100, 30,
                                                          get_data_for_calc['НКПР, об.%'])
        # 3. Запишем решение
        xw.Range(f'AA{index - 3}').value = zone_lclp[0]
        xw.Range(f'AB{index - 3}').value = zone_lclp[1]

        xw.Range(f'AA{index + 1}').value = zone_lclp_2[0]
        xw.Range(f'AB{index + 1}').value = zone_lclp_2[1]

        # ТОКСИ
        xw.Range(f'AC{index - 2}').value = round(toxi_fake(xw.Range(f'J{index - 2}').value)[0], 1)
        xw.Range(f'AD{index - 2}').value = round(toxi_fake(xw.Range(f'J{index - 2}').value)[1], 1)

        xw.Range(f'AC{index + 2}').value = round(toxi_fake(xw.Range(f'J{index + 2}').value)[0], 1)
        xw.Range(f'AD{index + 2}').value = round(toxi_fake(xw.Range(f'J{index + 2}').value)[1], 1)

    if get_data_for_calc['Тип дерева'] == 23:
        # ПОЖАР
        # 1. Получить экземпляр класса пожара
        fire_unit = calc_strait_fire.Strait_fire()
        # 2. Получить зоны классифицированные
        zone_array = fire_unit.termal_class_zone(get_data_for_calc['Площадь, м2'], 0.06, 100, 20, 1)
        zone_array_2 = fire_unit.termal_class_zone(get_data_for_calc['Площадь, м2'] / 10, 0.06, 100, 20, 1)
        # 3. Запишем решение
        # ПОЖАР+ТОКСИ
        xw.Range(f'P{index - 5}').value = zone_array[0]
        xw.Range(f'Q{index - 5}').value = zone_array[1]
        xw.Range(f'R{index - 5}').value = zone_array[2]
        xw.Range(f'S{index - 5}').value = zone_array[3]

        xw.Range(f'P{index - 4}').value = zone_array[0]
        xw.Range(f'Q{index - 4}').value = zone_array[1]
        xw.Range(f'R{index - 4}').value = zone_array[2]
        xw.Range(f'S{index - 4}').value = zone_array[3]

        # ТОКСИ
        xw.Range(f'AC{index - 5}').value = round(toxi_fake(xw.Range(f'J{index - 5}').value)[0]*0.01, 1)
        xw.Range(f'AD{index - 5}').value = round(toxi_fake(xw.Range(f'J{index - 5}').value)[1]*0.01, 1)

        xw.Range(f'AC{index - 4}').value = round(toxi_fake(xw.Range(f'J{index - 4}').value)[0]*0.01, 1)
        xw.Range(f'AD{index - 4}').value = round(toxi_fake(xw.Range(f'J{index - 4}').value)[1]*0.01, 1)

        # ПОЖАР+ТОКСИ
        xw.Range(f'P{index - 2}').value = zone_array_2[0]
        xw.Range(f'Q{index - 2}').value = zone_array_2[1]
        xw.Range(f'R{index - 2}').value = zone_array_2[2]
        xw.Range(f'S{index - 2}').value = zone_array_2[3]

        xw.Range(f'P{index - 1}').value = zone_array_2[0]
        xw.Range(f'Q{index - 1}').value = zone_array_2[1]
        xw.Range(f'R{index - 1}').value = zone_array_2[2]
        xw.Range(f'S{index - 1}').value = zone_array_2[3]

        # ТОКСИ
        xw.Range(f'AC{index - 2}').value = round(toxi_fake(xw.Range(f'J{index - 2}').value)[0]*0.01, 1)
        xw.Range(f'AD{index - 2}').value = round(toxi_fake(xw.Range(f'J{index - 2}').value)[1]*0.01, 1)

        xw.Range(f'AC{index - 1}').value = round(toxi_fake(xw.Range(f'J{index - 1}').value)[0]*0.01, 1)
        xw.Range(f'AD{index - 1}').value = round(toxi_fake(xw.Range(f'J{index - 1}').value)[1]*0.01, 1)


    if get_data_for_calc['Тип дерева'] == 9:
        # ПОЖАР
        # 1. Получить экземпляр класса пожара
        fire_unit = calc_strait_fire.Strait_fire()
        # 2. Получить зоны классифицированные
        zone_array = fire_unit.termal_class_zone(get_data_for_calc['Площадь, м2'], 0.06, 100, 20, 1)
        zone_array_2 = fire_unit.termal_class_zone(get_data_for_calc['Площадь, м2'] / 10, 0.06, 100, 20, 1)
        # 3. Запишем решение
        xw.Range(f'P{index - 5}').value = zone_array[0]
        xw.Range(f'Q{index - 5}').value = zone_array[1]
        xw.Range(f'R{index - 5}').value = zone_array[2]
        xw.Range(f'S{index - 5}').value = zone_array[3]

        xw.Range(f'P{index - 2}').value = zone_array_2[0]
        xw.Range(f'Q{index - 2}').value = zone_array_2[1]
        xw.Range(f'R{index - 2}').value = zone_array_2[2]
        xw.Range(f'S{index - 2}').value = zone_array_2[3]

        # ВЗРЫВ
        # 1. Получим класс для расчета взрыва
        explosion_unit = calc_tvs_explosion.Explosion()
        # 2. Получим зоны классифицированные
        zone_cls_array = explosion_unit.explosion_class_zone(3, DEGREE_OF_CLUTTER, xw.Range(f'J{index - 4}').value * 1000, 45320, 7, 2)
        # 3. Запишем решение
        xw.Range(f'T{index - 4}').value = zone_cls_array[1]
        xw.Range(f'U{index - 4}').value = zone_cls_array[2]
        xw.Range(f'V{index - 4}').value = zone_cls_array[3]
        xw.Range(f'W{index - 4}').value = zone_cls_array[4]
        xw.Range(f'X{index - 4}').value = zone_cls_array[5]

        # ПОЖАР-ВСПЫШКА
        # 1. Получим класс для расчета
        lclp_unit = calc_lower_concentration.LCLP()
        # 2. Получим зоны
        zone_lclp = lclp_unit.lower_concentration_limit(xw.Range(f'J{index - 1}').value * 1000, 100, 30,
                                                        get_data_for_calc['НКПР, об.%'])
        # 3. Запишем решение
        xw.Range(f'AA{index - 1}').value = zone_lclp[0]
        xw.Range(f'AB{index - 1}').value = zone_lclp[1]

    if get_data_for_calc['Тип дерева'] == 10:
        # ПОЖАР
        # 1. Получить экземпляр класса пожара
        fire_unit = calc_strait_fire.Strait_fire()
        # 2. Получить зоны классифицированные
        zone_array = fire_unit.termal_class_zone(get_data_for_calc['Площадь, м2'], 0.06, 100, 20, 1)
        zone_array_2 = fire_unit.termal_class_zone(get_data_for_calc['Площадь, м2'] / 10, 0.06, 100, 20, 1)
        # 3. Запишем решение
        xw.Range(f'P{index - 5}').value = zone_array[0]
        xw.Range(f'Q{index - 5}').value = zone_array[1]
        xw.Range(f'R{index - 5}').value = zone_array[2]
        xw.Range(f'S{index - 5}').value = zone_array[3]

        xw.Range(f'P{index - 2}').value = zone_array_2[0]
        xw.Range(f'Q{index - 2}').value = zone_array_2[1]
        xw.Range(f'R{index - 2}').value = zone_array_2[2]
        xw.Range(f'S{index - 2}').value = zone_array_2[3]

        # ВЗРЫВ
        # 1. Получим класс для расчета взрыва
        explosion_unit = calc_tvs_explosion.Explosion()
        # 2. Получим зоны классифицированные
        zone_cls_array = explosion_unit.explosion_class_zone(3, DEGREE_OF_CLUTTER, xw.Range(f'J{index - 4}').value * 1000, 45320, 7, 2)
        # 3. Запишем решение
        xw.Range(f'T{index - 4}').value = zone_cls_array[1]
        xw.Range(f'U{index - 4}').value = zone_cls_array[2]
        xw.Range(f'V{index - 4}').value = zone_cls_array[3]
        xw.Range(f'W{index - 4}').value = zone_cls_array[4]
        xw.Range(f'X{index - 4}').value = zone_cls_array[5]

        # ПОЖАР-ВСПЫШКА
        # 1. Получим класс для расчета
        lclp_unit = calc_lower_concentration.LCLP()
        # 2. Получим зоны
        zone_lclp = lclp_unit.lower_concentration_limit(xw.Range(f'J{index - 1}').value * 1000, 100, 30,
                                                        get_data_for_calc['НКПР, об.%'])
        # 3. Запишем решение
        xw.Range(f'AA{index - 1}').value = zone_lclp[0]
        xw.Range(f'AB{index - 1}').value = zone_lclp[1]

        # ТОКСИ
        xw.Range(f'AC{index - 3}').value = round(toxi_fake(xw.Range(f'J{index - 3}').value)[0], 1)
        xw.Range(f'AD{index - 3}').value = round(toxi_fake(xw.Range(f'J{index - 3}').value)[1], 1)

        xw.Range(f'AC{index}').value = round(toxi_fake(xw.Range(f'J{index}').value)[0], 1)
        xw.Range(f'AD{index}').value = round(toxi_fake(xw.Range(f'J{index}').value)[1], 1)

    if get_data_for_calc['Тип дерева'] == 24:
        xw.Range(f'AI{index - 5}').value = get_data_for_calc['Площадь, м2'] * 1.2
        xw.Range(f'AI{index - 4}').value = get_data_for_calc['Площадь, м2'] / 6

if __name__ == '__main__':
    print('Done')
