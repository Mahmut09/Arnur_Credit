import pandas as pd
from math import isnan
# Нет колонки "реструктуризация", "Списания"

columns_to_delete = ["Unnamed: 0", "№ по гр.", "Филиал", "Источник финансирования", "Баланс по сборам", "Баланс по госпошлине", "Баланс предоплаты по ОС", "Баланс предоплаты по %", "Баланс по дисконту", "Баланс по регулярному резерву по ОС", "Баланс по регулярному резерву по процентам", "Баланс по резерву МСФО", "Резерв МСФО по процентам", "Исполнительная надпись", "Тип займа", "Статус контракта", "Группа", "Групповое соглашение", "Остановка начисления процентов", "Дата остановки начисления процентов", "Остановка начисления штрафовв", "Дата остановки начисления  штрафов", "Вид деятельности", "Подвид деятельности", "Подцель микрокредита", "Тип залога", "Рыночная стоимость", "Залоговая стоимость", "Ступень займа", "ИИН", "Рекомендации по скорингу (КДН)", "Дата последней оплаты процентов", "Статус Судебник", "Ставка резерва по контракту на % и штрафы", "ID контракта", "Специалист по займам", "Дата рождения", "Просрочка ОС", "Просрочка %", "ИП", "Группа кредитных продуктов", "Пол клиента", "Факт.дни просрочки ОС, %, штрафы, отсроч %", "Дата создания Линии кредитов", "Номер кошелька", "Подразделение Линии кредитов", "Дата доступности Линии кредитов", "Номер линии кредитов", "ID линии кредитов", "Текущий лимит", "Магазин Линии кредитов", "GUID клиента", "Подпись кредитной линии", "Сумма Лимита Линии кредитов", "ID кошелька", "Баланс ОС, %, штрафы, отсроч %", "Ставка резерва по контракту на ОС", "Количество дней просрочки по статусу", "Пользовательский статус"]
file_path = "./Провизия Арнур/Открытые займы 30.11.2023 НЕ редактированная .xlsx"
segment_file_path = "./Провизия Арнур/30.11.2023/Сегмент.xlsx"

def add_new_column(df, column_before_insert, column_name, value):
    df.insert(loc=column_before_insert + 1, column=column_name, value=value)

def insert_column_after_and_remove(df, existing_column_name, insert_after_column_name):
    if existing_column_name not in df.columns:
        return

    if insert_after_column_name not in df.columns:
        return

    insert_after_index = df.columns.get_loc(insert_after_column_name)

    new_column_values = df[existing_column_name].copy()

    df.drop(columns=[existing_column_name], errors='ignore', inplace=True)
    
    df.insert(loc=insert_after_index + 1, column=existing_column_name, value=new_column_values)

    return df

# Удаляем все колонки, где есть Списанные в колонке Регион
def remove_column_written_off(df):
    df['Регион'] = df['Регион'].astype(str)
    df = df[~df['Регион'].str.contains('Списанные', case=False)]

    return df

def update_remainder_values(df):
    if 'Клиент' not in df.columns:
        return

    if 'Остаток суммы МКЛ' not in df.columns:
        return 

    duplicate_clients = df.duplicated(subset=['Клиент'], keep='first')

    df.loc[duplicate_clients, 'Остаток суммы МКЛ'] = 0

    return df

def remove_column(df: pd.DataFrame, column_name: str) -> pd.DataFrame:
    if column_name not in df.columns:
        print("Error")
        return
    
    df = df.drop(columns=[column_name])
    
    return df

def create_without_r(df):
    contract_column_index = df.columns.get_loc('Контракт')
    add_new_column(df, contract_column_index, 'Без Р', df['Контракт'].values)
    df['Без Р'] = df['Без Р'].str.rstrip('R')

    return df

# Удаляем букву R вначале колонки контракт
def delete_r_in_contract_column(df):
    df['Контракт'] = df['Контракт'].fillna('Неизвестно')
    mask = df['Контракт'].str.startswith('R')
    df.loc[mask, 'Контракт'] = df.loc[mask, 'Контракт'].str.lstrip('R')

    return df

def get_list_without_GESF(df):
    without_GESF_list = []
    mask = df['Ставка ГЭСВ'].isna()
    for i in df[mask]['Контракт']:
        if not i == "Неизвестно":
            without_GESF_list.append(i)

    return without_GESF_list

def insert_GESF_value(df, without_GESF_list, replacement_values):
    if len(without_GESF_list) != len(replacement_values):
        print("Размеры не совпадают")
        return df

    indices_to_replace = df[df['Контракт'].isin(without_GESF_list)].index
    df.loc[indices_to_replace, 'Ставка ГЭСВ'] = replacement_values
    
    return df

# Если есть слово отсрочка в колоне кредитный продукт, то вконце контракта убираем Р 
def delete_r_with_deferment(df):
    def remove_r_from_contract(row):
        contract_value = row['Контракт']
        credit_product_value = row['Кредитный продукт']
        
        if pd.notnull(credit_product_value):
            credit_product_lower = str(credit_product_value).lower()
            if ('отсрочка' in credit_product_lower) or ('сусн' in credit_product_lower):
                return str(contract_value).rstrip('R')  # Удаляем все "R" в конце строки

        return contract_value
    
    df['Контракт'] = df.apply(remove_r_from_contract, axis=1)
    return df

def create_balance_column(df):
    df['Баланс'] = df['Баланс по ОД'] + df['Баланс по %'] + df['Баланс по штрафам']

    return df

def create_segment_column(df, segment_df):
    segment_mapping = dict(zip(segment_df['Кредитный продукт'], segment_df['Сегмент']))
    df['Сегмент'] = df['Кредитный продукт'].map(segment_mapping)

    return df

def fill_na_with_zero(df, column_name):
    if column_name not in df.columns:
        df[column_name] = 0
    else:
        df[column_name] = df[column_name].fillna(0)

    return df

def create_level_of_delinquency_column(value):
    if value == 0:
        return 0
    elif value < 31:
        return 1
    elif value < 61:
        return 2
    elif value < 90:
        return 3
    else:
        return 4

def calculate_restructuring(value):
    print(value)
    if "RRRR" in str(value):
        return 4
    elif "RRR" in str(value):
        return 3
    elif "RR" in str(value):
        return 2
    elif "R" in str(value):
        return 1
    else:
        return 0


df = pd.read_excel(file_path, skiprows=6)

df = remove_column_written_off(df)

df = delete_r_with_deferment(df)

df = create_without_r(df)

df = delete_r_in_contract_column(df)

df = insert_column_after_and_remove(df, "Остаток суммы МКЛ", "Без Р")
df = insert_column_after_and_remove(df, "Дата закрытия МКЛ", "Без Р")
df = insert_column_after_and_remove(df, "Дата открытия МКЛ", "Без Р")
df = insert_column_after_and_remove(df, "Сумма МКЛ", "Без Р")

df = insert_column_after_and_remove(df, "Ставка ГЭСВ", "Баланс по штрафам")
df = insert_column_after_and_remove(df, "Цель микрокредита", "Ставка ГЭСВ")
df = update_remainder_values(df)

# Удаляем ненужные колонки
for column in columns_to_delete:
    df = remove_column(df, column)

without_GESF_list = get_list_without_GESF(df)
# print(len(without_GESF_list))
# insert_GESF_value(df, without_GESF_list, [111, 222])

df = create_balance_column(df)
df = insert_column_after_and_remove(df, "Баланс", "Кредитный продукт")

segment_df = pd.read_excel(segment_file_path)
df = create_segment_column(df, segment_df)
df = insert_column_after_and_remove(df, "Сегмент", "Баланс")

df = fill_na_with_zero(df, "Сумма МКЛ")
df = fill_na_with_zero(df, "Дата открытия МКЛ")
df = fill_na_with_zero(df, "Дата закрытия МКЛ")
df = fill_na_with_zero(df, "Остаток суммы МКЛ")
df = fill_na_with_zero(df, "Количество дней просрочки фактическое")

df['Уровень просрочки'] = df['Количество дней просрочки фактическое'].apply(create_level_of_delinquency_column)

df['Реструктуризация'] = df['Контракт'].apply(calculate_restructuring)

df = fill_na_with_zero(df, "Списания")


df.to_excel("test.xlsx", index=False)