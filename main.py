import pandas as pd

# Параметры файлов
source_file_1 = 'штатный состав 2024-2025 по кафедре радиоэлектроники дополненная на 25 сентября.xlsx'
source_file_2 = 'Проект нагрузки.xlsx'
target_file = 'План по нагрузке.xlsx'

# Чтение данных из первого файла ("Состав кафедры")
try:
    df_orders = pd.read_excel(source_file_1, header=None, usecols=[1, 3])
    df_orders_not_nan = df_orders.dropna()  # Удаляем строки с NaN
    df_orders_not_dupl = df_orders_not_nan.drop_duplicates()  # Удаляем дубликаты

    # Преобразуем данные первого файла в список строк
    rows_orders = df_orders_not_dupl.values.tolist()
except Exception as e:
    print(f"Ошибка при обработке файла '{source_file_1}': {e}")
    rows_orders = []

# Чтение данных из второго файла ("Проект нагрузки")
try:
    df_nagruzka = pd.read_excel(source_file_2, header=None, skiprows=8)  # Пропускаем первые 8 строк

    # Преобразуем данные второго файла в список строк
    rows_nagruzka = df_nagruzka.values.tolist()
except Exception as e:
    print(f"Ошибка при обработке файла '{source_file_2}': {e}")
    rows_nagruzka = []

# Формирование итогового списка строк
combined_rows = []
for nagruzka_row in rows_nagruzka:
    combined_rows.append(nagruzka_row)  # Добавляем строку из "Проект нагрузки"
    combined_rows.extend(rows_orders)  # Добавляем все строки из "Состав кафедры"

# Создаём DataFrame из комбинированных строк
combined_df = pd.DataFrame(combined_rows)

# Сохраняем данные в итоговый файл
with pd.ExcelWriter(target_file, engine='openpyxl', mode='w') as writer:
    combined_df.to_excel(writer, index=False, header=False)

print(f"Данные успешно записаны в файл '{target_file}'.")
