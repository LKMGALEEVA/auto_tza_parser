import pandas as pd
import re
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

# Пути к файлам
file1 = 'C:/Users/User/Documents/КУРСК_2_ТЗА/test_schemes_exit/KUR.0120.20UJA.JNA.SR.EC0001.xlsx'
file2 = 'C:/Users/User/Documents/КУРСК_2_ТЗА/test_tables/KUR.0120.20UJA.JNA.SR.EC0001_Технологическое задание. Запорная арматура.xlsx'
output_file = 'результат_сравнения_3.xlsx'

# Чтение данных из файлов
df1 = pd.read_excel(file1)
df2 = pd.read_excel(file2)

# Получаем значения из первого столбца
col1 = df1.iloc[:, 0].astype(str).str.strip()
col2 = df2.iloc[:, 0].astype(str).str.strip()

# Создаем множество ключей для поиска
keys_set = set(col1)

# Результаты сравнения
results = []

# Словарь для отслеживания использованных ключей (чтобы не дублировать)
used_keys = set()

# Проверяем каждую строку из второго файла
for value in col2:
    matched = False
    for key in keys_set:
        # Ищем точное совпадение в начале строки
        if re.match(r'^' + re.escape(key) + r'\b', value):
            results.append([key, "Совпадает", "green"])
            used_keys.add(key)
            matched = True
            break
    if not matched:
        results.append([value, f"Только в {file2}", "red"])

# Добавляем оставшиеся непарные элементы из первого файла
for key in keys_set:
    if key not in used_keys:
        results.append([key, f"Только в {file1}", "red"])

# Создаем Excel-файл и записываем результаты
wb = Workbook()
ws = wb.active
ws.title = "Результат сравнения"

# Заголовки
headers = ["Значение", "Статус"]
for col_num, data in enumerate(headers, 1):
    ws.cell(row=1, column=col_num, value=data)

# Заполняем данными
for row_num, (value, status, color) in enumerate(results, 2):
    ws.cell(row=row_num, column=1, value=value)
    ws.cell(row=row_num, column=2, value=status)

# Цвета
color_map = {
    "green": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),  # зелёный
    "red": PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")     # красный
}

# Применяем стили
for row_num, (_, _, color) in enumerate(results, 2):
    ws.cell(row=row_num, column=1).fill = color_map[color]
    ws.cell(row=row_num, column=2).fill = color_map[color]

# Автоподгон ширины столбцов
for col in ws.columns:
    max_length = 0
    column = col[0].column_letter
    for cell in col:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        except:
            pass
    adjusted_width = min(max_length + 2, 50)
    ws.column_dimensions[column].width = adjusted_width

# Сохраняем файл
wb.save(output_file)
print(f"✅ Результаты сохранены в '{output_file}'")