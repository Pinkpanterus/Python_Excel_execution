#!/usr/bin/python
# -*- coding: utf8 -*-

from tkinter import filedialog
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import MyFunctions, MyClasses
import os


#Указываем пути к исходным файлам
rnpd_file_path = filedialog.askopenfilename(title="Выберите файл РНПД")
cfo140_file_path = filedialog.askopenfilename(title="Выберите файл 014ЦФО")

cfo140_new_file_path = cfo140_file_path.split('.')[0] + "_result" + ".xls"

#Проверяем что пути к файлам указаны
if(len(rnpd_file_path) == 0 or len(cfo140_file_path) == 0):
    print("Файлы не выбраны. Перезапустите программу и выберете файлы!")
    raise SystemExit

#Загружаем данные из файла cfo140
cfo140_file = openpyxl.open(cfo140_file_path, data_only=False, read_only=False)
working_sheet_cfo140_file = cfo140_file.active
MyFunctions.prepare_sheet(working_sheet_cfo140_file) #Убираем объединенные ячейки в файле 014ЦФО

#Начинаем обрабатывать файлы
rnpd_new_file_path = MyFunctions.convert_XLSB_to_XLSX_and_get_new_filepath_through_pandas(rnpd_file_path, 'Шаблон') # я так понял что могу только считывать данные из файла xlxb, а менять ничего не могу там, поэтому делаю копию в другом формате
rnpd_file_result = openpyxl.open(rnpd_new_file_path, data_only=False, read_only=False) #сюда буду записывать все изменения файла РНПД, отсюда сотрудник потом сможет перекопировать данные в файл .xlxb
rnpd_file_result_ws = rnpd_file_result.active
rnpd_file_result_ws.title = 'Шаблон'

matched_item_list = [] #сюда складываю все совпадения по индексу

# запоминаю заголовки таблиц, чтобы потом искать значение по заголовку
rnpd_file_column_names=[]
for row_index, row in enumerate(rnpd_file_result_ws):
    for cell_index, cell in enumerate(row):
        if (row_index == 0):
            rnpd_file_column_names.append(cell.value)
# print(rnpd_file_column_names)

cfo140_column_names = []
for row_index, row in enumerate(working_sheet_cfo140_file):
    for cell_index, cell in enumerate(row):
        if (row_index == 1):
            cfo140_column_names.append(cell.value)

for cfo140_row_index, cfo140_row in enumerate(working_sheet_cfo140_file):
        if (cfo140_row_index > 1):  # пропускаю первую пустую строку и заголовок в файле
            cfo140_row_key = str(cfo140_row[cfo140_column_names.index('Системный номер договора')].value) + '-' + str(
                cfo140_row[cfo140_column_names.index('Статья затрат')].value) + '-' + str(
                cfo140_row[cfo140_column_names.index('ШПП')].value) + '-' + str(
                cfo140_row[cfo140_column_names.index('Субъект')].value) + '-' + str(cfo140_row[cfo140_column_names.index('БП')].value)

            for rpnd_row_index, rnpd_row in enumerate(rnpd_file_result_ws):
                #print(f'ЦФО индекс: {cfo140_row_index}, РНПД индекс: {rpnd_row_index}')
                for cell_index, cell in enumerate(row):
                    if (rpnd_row_index > 0):
                        rnpd_row_key = str(rnpd_row[rnpd_file_column_names.index('№ Договора')].value.split()[0]) + '-' + str(rnpd_row[rnpd_file_column_names.index('СЗ')].value) + '-' + str(rnpd_row[rnpd_file_column_names.index('ШПЗ')].value) + '-' + str(rnpd_row[rnpd_file_column_names.index('Субъект')].value).rjust(6, "0") + '-' + str(rnpd_row[rnpd_file_column_names.index('Бизнес процесс')].value).rjust(6, "0")
                        if(cfo140_row_key == rnpd_row_key and cell_index == rnpd_file_column_names.index('Сумма строки распределения')):
                            old_summ = rnpd_row[rnpd_file_column_names.index('Сумма строки распределения')].value
                            new_summ = cfo140_row[cfo140_column_names.index('Сумма без НДС (руб)')].value
                            matched_item = MyClasses.MatchItem(cfo140_row_index, rpnd_row_index, old_summ, new_summ, cfo140_row_key)
                            matched_item_list.append(matched_item)
                            #Обрабатываю сопадения после цикла


#cfo140_sheet = cfo140_file.active
for cfo140_row_index, cfo140_row in enumerate(working_sheet_cfo140_file):
    if (cfo140_row_index > 1):  # пропускаю первую пустую строку и заголовок в файле
        cfo140_row_key = str(cfo140_row[cfo140_column_names.index('Системный номер договора')].value) + '-' + str(
            cfo140_row[cfo140_column_names.index('Статья затрат')].value) + '-' + str(
            cfo140_row[cfo140_column_names.index('ШПП')].value) + '-' + str(
            cfo140_row[cfo140_column_names.index('Субъект')].value) + '-' + str(
            cfo140_row[cfo140_column_names.index('БП')].value)

        row_key_matches_count = list(map(lambda x: x.key, matched_item_list)).count(cfo140_row_key)
        #print(f'Row key: {cfo140_row_key}, matches count: {matched_count}')
        if (row_key_matches_count > 1):
            rnpd_rows = list(filter(lambda x: cfo140_row_key in x.key, matched_item_list))
            rows_numbers = list(map(lambda x: x.rpnd_row_index, rnpd_rows)) #номера строк в файле РНПД

            cfo140_row[cfo140_column_names.index('Комментарий')].value = f"Найдено более одного сопадения в файле РНПД. Строки: {', '.join(map(str, rows_numbers))}!"
            cfo140_row[cfo140_column_names.index('Комментарий')].fill = PatternFill("solid", start_color="00FF6600")
            '''
            2)	Найдено более одного совпадения. Необходимо только записать Результат. В результате указать, что найдено более одного совпадения, указать номера строк из шаблона РНПД, строки подкрасить светло красноватым не ярким цветом.
            '''
        elif (row_key_matches_count == 1):
            matched_item = list(filter(lambda x: x.key == cfo140_row_key, matched_item_list))[0]

            if(matched_item.old_summ != matched_item.new_summ):
                rnpd_cell = rnpd_file_result_ws.cell(matched_item.rpnd_row_index +1,  rnpd_file_column_names.index('Сумма строки распределения') +1)
                rnpd_cell.value = matched_item.new_summ
                rnpd_cell.fill = PatternFill("solid", start_color="00FF6600")

                cfo140_row[cfo140_column_names.index('Комментарий')].value = f"В файле РНПД отредактирована сумма по строке: {matched_item.rpnd_row_index +1}. Старая сумма: {matched_item.old_summ}. Новая сумма {matched_item.new_summ}."
            else:
                cfo140_row[cfo140_column_names.index('Комментарий')].value = f"В файле РНПД сумма по строке: {matched_item.rpnd_row_index +1} совпадает."
            '''
            1)	Найдено одно совпадение. Необходимо в шаблоне РНПД на листе Шаблон поменять только сумму в колонке R, сумма строки распределения. Сумму указываем из реестра по 014 ЦФО колонка F, сумма без НДС. Результат записать в реестр по 014 ЦФО, колонку можно создать справа таблицы. В результате указать, что найдено совпадение, указать 
            '''
        else:
            BE = cfo140_row[cfo140_column_names.index('БЕ')].value
            CFO = cfo140_row[cfo140_column_names.index('ЦФО-получатель')].value
            KONTRAGENT = cfo140_row[cfo140_column_names.index('Наименование контрагента')].value
            DOGOVOR = cfo140_row[cfo140_column_names.index('Системный номер договора')].value
            SUMMA = cfo140_row[cfo140_column_names.index('Сумма без НДС (руб)')].value
            DATA = cfo140_row[cfo140_column_names.index('Дата фактического оказания услуг')].value.strftime("%d.%m.%Y")
            SHPP = cfo140_row[cfo140_column_names.index('ШПП')].value
            SZ = cfo140_row[cfo140_column_names.index('Статья затрат')].value
            PROGRAMMA = cfo140_row[cfo140_column_names.index('программа')].value
            SUBJECT =  cfo140_row[cfo140_column_names.index('Субъект')].value
            BP = cfo140_row[cfo140_column_names.index('БП')].value
            MVZ = cfo140_row[cfo140_column_names.index('МВЗ')].value

            row = [BE, CFO, '', KONTRAGENT, '', DOGOVOR, '', '', '', '', '', '', '', DATA, '', '', '', SUMMA, SHPP, SZ, '', PROGRAMMA, SUBJECT, '', BP, MVZ, '', '', '', '', '', '', '', '', '', '', '', '', '', '']

            rnpd_file_result_ws.append(row)
            rnpd_file_result.save(rnpd_new_file_path)

            for cell_index, cell in enumerate(rnpd_file_result_ws[rnpd_file_result_ws.max_row], start=1):
                rnpd_cell = rnpd_file_result_ws.cell(rnpd_file_result_ws.max_row, cell_index)
                rnpd_cell.fill = PatternFill("solid", start_color="00FFFF00")

            cfo140_row[cfo140_column_names.index('Комментарий')].value = f"В файле РНПД не найдено совпадений. Добавлена строка: {rnpd_file_result_ws.max_row}!"
            cfo140_row[cfo140_column_names.index('Комментарий')].fill = PatternFill("solid", start_color="00FFFF00")
            '''
            Не найдено ни одного совпадения. В шаблоне РНПД необходимо завести новую строку, перенести данные из реестра 014 ЦФО.  Соответствия столбцов указаны в таблице ниже.
            В результате указать, что не найдено совпадение, создана строка, указать номер строки из шаблона РНПД,  строку подкрасить желтым
            '''

            cfo140_file.save(cfo140_file_path.split('.')[0] + "_result" + ".xls")

rnpd_file_result.save(rnpd_new_file_path)
cfo140_file.save(cfo140_new_file_path)

#Запускаю сформированные файлы
os.startfile(rnpd_new_file_path)
os.startfile(cfo140_new_file_path)


