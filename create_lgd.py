import os
import shutil
import pandas as pd
from datetime import datetime
import openpyxl

from get_last_day_of_prev_month import get_last_day_of_previous_month, get_last_day_of_before_previous_month


def copy_files(source_folder, destination_folder):
    if not os.path.exists(source_folder):
        print(f"Исходная папка '{source_folder}' не существует.")
        return

    if not os.path.exists(destination_folder):
        os.mkdir(f"{last_day_previous_month}/LGD")

    files = os.listdir(source_folder)

    for file in files:
        source_path = os.path.join(source_folder, file)
        destination_path = os.path.join(destination_folder, file)
        shutil.copy2(source_path, destination_path)
        print(f"Файл '{file}' скопирован в '{destination_folder}'.")

def filter_and_append_sheet(source_file, destination_file, segment, sheet_name, unique_sheet_path):
    df = pd.read_excel(source_file)

    # Фильтруем данные
    filtered_df = df[(df['Сегмент'] == segment) & (df['Количество дней просрочки фактическое'] > 90)]

    # Загружаем данные из файла с уникальными записями
    unique_df = pd.read_excel(unique_sheet_path, header=None, names=['Контракт'])

    # Используем метод merge для объединения по столбцу "Контракт"
    result_df = pd.merge(filtered_df, unique_df, how='left', indicator=True).query('_merge == "left_only"').drop('_merge', axis=1)

    # Открываем существующий файл Excel
    with pd.ExcelWriter(destination_file, engine='openpyxl', mode='a') as writer:
        # Загружаем существующий лист
        writer.book = openpyxl.load_workbook(destination_file)

        # Загружаем существующий лист (если существует)
        writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)

        # Записываем новый лист
        result_df.to_excel(writer, sheet_name=sheet_name, index=False, columns=['Контракт'])
    df = pd.read_excel(source_file)

    filtered_df = df[(df['Сегмент'] == segment) & (df['Количество дней просрочки фактическое'] > 90)]

    unique_df = pd.read_excel(unique_sheet_path, header=None, names=['Контракт'])

    result_df = pd.merge(filtered_df, unique_df, how='left', indicator=True).query('_merge == "left_only"').drop('_merge', axis=1)

    with pd.ExcelWriter(destination_file, engine='xlsxwriter', mode='a') as writer:
        result_df.to_excel(writer, sheet_name=sheet_name, index=False, columns=['Контракт'])
    df = pd.read_excel(source_file)

    filtered_df = df[(df['Сегмент'] == segment) & (df['Количество дней просрочки фактическое'] > 90)]

    unique_df = pd.read_excel(unique_sheet_path, header=None, names=['Контракт'])

    result_df = pd.merge(filtered_df, unique_df, how='left', indicator=True).query('_merge == "left_only"').drop('_merge', axis=1)

    with pd.ExcelWriter(destination_file, engine='openpyxl', mode='a') as writer:
        result_df.to_excel(writer, sheet_name=sheet_name, index=False, columns=['Контракт'])

# Прошлый месяц
last_day_previous_month = get_last_day_of_previous_month()
# Позапрошлый месяц
last_day_before_prev_month = get_last_day_of_before_previous_month()

source_folder_path = f'./Провизия Арнур/{last_day_before_prev_month}/LGD'
destination_folder_path = f'./{last_day_previous_month}/LGD'

copy_files(source_folder_path, destination_folder_path)

source_excel_path = f'./{last_day_previous_month}/Открытые займы {last_day_previous_month}.xlsx'
destination_excel_path_agro = f'{destination_folder_path}/Агро.xlsx'
unique_sheet_path = f'{destination_folder_path}/Уникальные.xlsx'

sheet_name = datetime.strptime(last_day_previous_month, '%d.%m.%Y').strftime('%Y%m')
filter_and_append_sheet(source_excel_path, destination_excel_path_agro, "Агро", sheet_name, unique_sheet_path)

