import tkinter
from tkinter import filedialog
import os
import pandas as pd


root = tkinter.Tk()
root.withdraw() #use to hide tkinter window

def prepare_sheet(sheet):
    merged_cells = list(map(str, sheet.merged_cells.ranges))  # Получаю список объединенных диапазонов
    # Разъединяю объединенные ячейки и дублирую запись
    for item in merged_cells:
        sheet.unmerge_cells(item)
        merged_cells_range = item.split(":")
        if merged_cells_range[0][0] == merged_cells_range[1][0]:
            letter = item.split(":").pop(0)[0]  # Символ столбца диапазона
            start = int(item.split(":").pop(0)[1:])  # Начало диапазона
            end = int(item.split(":").pop()[1:])  # Конец диапазона
            copy_cell = sheet[(letter + str(start))].value
            for n in range(start, end + 1):
                cell = letter + str(n)
                sheet[cell].value = copy_cell

def restore_rnpd_file_format(sheet):
    for row_index, row in enumerate(sheet):
        for cell_index, cell in enumerate(row):
            if (row_index > 0):
                row[6].value = str(row[6].value).replace(',', '.')
                if(len(str(row[21].value)) >1):
                    row[21].value = str(row[21].value).rjust(10, "0")
                row[22].value = str(row[22].value).rjust(6, "0")
                row[24].value = str(row[24].value).rjust(6, "0")
                row[25].value = str(row[25].value).rjust(3, "0")

def search_for_file_path(file_name):
    currdir = os.getcwd()
    temp_dir = filedialog.askdirectory(parent=root, initialdir=currdir, title=f'Please select a directory of file: {file_name}')
    if len(temp_dir) > 0:
        print ("You chose: %s" % temp_dir)
    return temp_dir


def convert_XLSB_to_XLSX_and_get_new_filepath_through_pandas(xlsb_file_path, sheet_name):
    rnpd_file = pd.read_excel(xlsb_file_path, engine='pyxlsb', sheet_name = sheet_name)
    xlsx_file_path = xlsb_file_path.split(".")[0] + "_result" + ".xlsx"
    rnpd_file.to_excel(xlsx_file_path, index = False)
    return xlsx_file_path


def get_cell_range(worksheet, min_row, max_row, min_column, max_column):
    cell_range=[worksheet.cell(row=i,column=j).value for i in range(min_row, max_row) for j in range(min_column, max_column)]
    return cell_range

