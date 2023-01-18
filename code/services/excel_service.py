import sys
from pathlib import Path
from win32com.client import constants, Dispatch

from constants import DATA_SHEET_NAME, PIVOT_SHEET_NAME, PIVOT_TABLE_NAME, BASE_DIR, OUTPUT_PATH


def create_sheet(wb, name):
    """Создает лист"""
    wb.Sheets.Add().Name = name
    return wb.Sheets(name)

def get_data_sheet(wb):
    """Возвращает Лист с данными"""
    return wb.Sheets(DATA_SHEET_NAME)

def get_sheets(wb):
    """Возвращает листы данных и свойдной таблицы"""
    return get_data_sheet(wb), create_sheet(wb, PIVOT_SHEET_NAME)

def clear_pts(ws):
    """Отчищает лист сводной таблицы"""
    for pt in ws.PivotTables():
        pt.TableRange2.Clear()

def insert_pt_fields(pt):
    """Формирует сводную таблицу"""
    # https://docs.microsoft.com/en-us/office/vba/api/excel.xlpivotfieldorientation
    field_rows = {}
    field_rows['Код'] = pt.PivotFields("Валюта (Код)")

    field_rows['Код'].Orientation = 1
    field_rows['Код'].Position = 1

    # https://docs.microsoft.com/en-us/office/vba/api/excel.xlconsolidationfunction
    field_values = {}
    field_values['Изменение'] = pt.PivotFields("Изменение")
    field_values['Изменение'].Orientation = 4
    field_values['Изменение'].Function = -4157
    field_values['Изменение'].NumberFormat = "0,000"
    field_values['Изменение'].Name = 'Суммарное изенение курса'
    field_values['Изменение'].Position = 1

    field_values['Изменение среднее'] = pt.PivotFields("Изменение")
    field_values['Изменение среднее'].Orientation = 4
    field_values['Изменение среднее'].Function = -4106
    field_values['Изменение среднее'].NumberFormat = "0,000"
    field_values['Изменение среднее'].Name = 'Среденее изенение курса'
    field_values['Изменение среднее'].Position = 2

    field_values['Курс'] = pt.PivotFields('Курс')
    field_values['Курс'].Orientation = 4
    field_values['Курс'].Function = -4106 
    field_values['Курс'].NumberFormat = "0,000"
    field_values['Курс'].Name = 'Средний курс'
    field_values['Курс'].Position = 3

    field_filters = {}
    field_filters['Дата'] = pt.PivotFields('Дата')
    
    field_filters['Дата'].Orientation = 3
    field_filters['Дата'].Position = 1

def create_pivot_table(filename):
    """Создает сводную таблицу"""
    c_path = BASE_DIR / OUTPUT_PATH / 'excel'
    filename =  f'{filename}.xlsx'

    print('INFO:    ', 'Создание сводной таблицы')
    excel = Dispatch('Excel.Application')
    # excel.Visible = True
    wb = excel.Workbooks.Open(c_path / filename)

    ws_data, ws_pivot = get_sheets(wb)

    clear_pts(ws_pivot)
    pt_cache = wb.PivotCaches().Create(1, SourceData=ws_data.UsedRange)
    pt = pt_cache.CreatePivotTable(ws_pivot.Range('A1'), PIVOT_TABLE_NAME)
    insert_pt_fields(pt)
    pt.ColumnGrand = False
    print('INFO:    ', 'Смежная таблица создана')
    wb.Close(SaveChanges=1)
    excel.Quit()
