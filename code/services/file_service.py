from constants import OUTPUT_PATH, DATA_SHEET_NAME, BASE_DIR
from utils import get_or_create_path

def write_csv(filename, path, df):
    """Создает csv файл"""
    filepath = get_or_create_path(path, 'csv')
    df.to_csv(filepath / f'{filename}.csv', index=False)

def write_xl(filename, path, df):
    """Создает excel файл"""
    filepath = get_or_create_path(path, 'excel')
    df.to_excel(filepath / f"{filename}.xlsx", sheet_name=DATA_SHEET_NAME, index=False)

def write_file(filename, df):
    """Создает файлы с данными"""
    output_path = get_or_create_path(BASE_DIR, OUTPUT_PATH)
    write_csv(filename, output_path, df)
    write_xl(filename, output_path, df)
