import pandas as pd


columns_to_delete = ["Unnamed: 0", "№ по гр.", "Филиал", "Источник финансирования", "Баланс по сборам", "Баланс по госпошлине", "Баланс предоплаты по ОС", "Баланс предоплаты по %", "Баланс по дисконту", "Баланс по регулярному резерву по ОС", "Баланс по регулярному резерву по процентам", "Баланс по резерву МСФО", "Резерв МСФО по процентам", "Исполнительная надпись", "Тип займа", "Статус контракта", "Группа", "Групповое соглашение", "Остановка начисления процентов", "Дата остановки начисления процентов", "Остановка начисления штрафовв", "Дата остановки начисления  штрафов", "Вид деятельности", "Подвид деятельности", "Подцель микрокредита", "Тип залога", "Рыночная стоимость", "Залоговая стоимость", "Ступень займа", "ИИН", "Рекомендации по скорингу (КДН)", "Дата последней оплаты процентов", "Статус Судебник", "Ставка резерва по контракту на % и штрафы", "ID контракта", "Специалист по займам", "Дата рождения", "Просрочка ОС", "Просрочка %", "ИП", "Группа кредитных продуктов", "Пол клиента", "Факт.дни просрочки ОС, %, штрафы, отсроч %", "Дата создания Линии кредитов", "Номер кошелька", "Подразделение Линии кредитов", "Дата доступности Линии кредитов", "Номер линии кредитов", "ID линии кредитов", "Текущий лимит", "Магазин Линии кредитов", "GUID клиента", "Подпись кредитной линии", "Сумма Лимита Линии кредитов", "ID кошелька", "Баланс ОС, %, штрафы, отсроч %"]
file_path = "./Провизия Арнур/Открытые займы 30.11.2023 НЕ редактированная .xlsx"


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

df = pd.read_excel(file_path, skiprows=6)

df = create_without_r(df)

df = delete_r_in_contract_column(df)

df = insert_column_after_and_remove(df, "Остаток суммы МКЛ", "Без Р")
df = insert_column_after_and_remove(df, "Дата закрытия МКЛ", "Без Р")
df = insert_column_after_and_remove(df, "Дата открытия МКЛ", "Без Р")
df = insert_column_after_and_remove(df, "Сумма МКЛ", "Без Р")

df = update_remainder_values(df)
# Удаляем ненужные колонки
for column in columns_to_delete:
    df = remove_column(df, column)


df.to_excel("test.xlsx", index=False)


