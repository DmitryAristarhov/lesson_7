# -*- coding: utf-8 -*-
import csv
import json
from docxtpl import DocxTemplate
from timeit import timeit

# Задание
# LIGHT:
# 1) Вручную создать текстовый файл с данными (например, марка авто, модель авто, расход топлива, стоимость).
# 2) Создать doc шаблон, где будут использованы данные параметры.
# 3) Автоматически сгенерировать отчет о машине в формате doc (как в видео 7.2).
# 4) Создать csv файл с данными о машине.
# 5) Создать json файл с данными о машине.
# PRO:
# LIGHT +
# 6) Замерить время генерации отчета (время выполнения пункта 3).
# В каждый файл пунктов 4 и 5 добавить параметр: время, затраченное на генерацию отчета.

###################################
# Прочитаем исходный файл с данными
FILENAME = 'input_text.csv'
with open(FILENAME, encoding='cp1251') as file:
    shopping_list = [row for row in csv.DictReader(file)]

for row in shopping_list:
    print(row)
print()

# Сгенерируем отчет в docx
context = {
    'my_title': 'Список покупок',
    'my_list': [i['Название'] + ' - ' + i['Количество'] + ' ' + i['Единица измерения'] for i in shopping_list[:-1]]
}

# По идее нужно эти команды выполнить, но мы их во время замеров выволним)
# doc = DocxTemplate('template.docx')
# doc.render(context)
# doc.save('Список покупок(вымышленный).docx')

# Измерим время, затраченное на генерацию отчета
# Чтобы его измерить, нам нужна строка кода в строковой переменной.
code_string = "doc = DocxTemplate('template.docx'); doc.render(context); doc.save('Список покупок(вымышленный).docx')"
repeats = 1  # Если захотим сделать несколько замеров для точности)
time = timeit(code_string, number=repeats, globals=globals()) / repeats
print(time)

# Раз уж я текстовый файл создал сразу в csv, то, чтоб не скучно было, запишу его другим способом)
title = tuple(k for k in shopping_list[0])
data = [tuple(d.values()) for d in shopping_list]

# отсортировав, предварительно, например по магазинам.
data = sorted(data[:-1], key=lambda x: x[3]) + data[-1:]  # -1 чтобы 'Итого' не сортировать)

# Ну и время генерации отчета добавим
# Не знаю как правильно боротся локализацией. Эксель числа с точкой не понимает. Запятую хочет.
time_str = str(round(time, 4)).replace('.', ',')  # Делаем пока костыли)
data.append(('Время генерации отчета:', '', '', '', '', '{}'.format(time_str)))

# Записываем что получилось
with open('Список покупок(вымышленный).csv', 'w', newline='') as csv_file:
    writer = csv.writer(csv_file)
    writer.writerow(title)
    writer.writerows(data)

# Теперь json
# Переведем строки в числа, формат позволяет)
for i in shopping_list[:-1]:
    i['Количество'] = int(i['Количество'])
    i['Примерная стоимость за единицу'] = round(float(i['Примерная стоимость за единицу']), 2)
    i['Общая стоимость'] = round(float(i['Общая стоимость']), 2)

# Переделаем 'Итого'
shopping_list[-1] = {'Итого': round(float(shopping_list[-1]['Общая стоимость']), 2)}

# Добавим время генерации отчета
shopping_list.append({'Время генерации отчета': round(time, 4)})

# Записываем что получилось. Только добавим ensure_ascii=False что бы читать можно было в блокноте)
with open('Список покупок(вымышленный).json', 'w') as json_file:
    json.dump(shopping_list, json_file, ensure_ascii=False)
