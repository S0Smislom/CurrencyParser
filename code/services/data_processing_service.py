
import pandas as pd
from datetime import datetime
from .file_service import write_file


def to_float(x):
    """Преобразует данные в float"""
    return float( '.'.join(x.split(',')) )

def prepare_data(df):
    """Преобразует данные"""
    df['Номинал'] = df['Номинал'].map(lambda x: int(x.split(' ')[-1]))
    df['Курс'] = df['Курс'].map(to_float)
    df['Изменение'] = df['Изменение'].map(to_float)
    df['Валюта (Код)'] = df.apply(lambda x: f"{x['Валюта']} ({x['Код']})" , axis=1)
    return df

def data_processing(head, data):
    """Обрабатывает и записывает данные в файл"""
    print('INFO:    ', 'Подготовка данных')
    df = pd.DataFrame(data, columns=head)
    df = prepare_data(df)
    filename = datetime.now().strftime("%m.%d.%Y")
    write_file(filename, df)
    return filename
