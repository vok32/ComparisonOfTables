import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import os

def get_next_filename(output_file):
    base_name, ext = os.path.splitext(output_file)
    version = 1
    while os.path.exists(output_file):
        version += 1
        output_file = f"{base_name}_v{version}{ext}"
    return output_file

def compare_excel_tables(file1_path, file2_path, output_path):
    # Чтение таблиц из файлов Excel
    table1 = pd.read_excel(file1_path)
    table2 = pd.read_excel(file2_path)

    # Проверка наличия одинаковых столбцов
    if list(table1.columns) != list(table2.columns):
        raise ValueError("Таблицы должны иметь одинаковые столбцы для сравнения.")

    # Загрузка второго файла для модификации
    workbook = load_workbook(file2_path)
    sheet = workbook.active

    # Цвет заливки для выделения различий
    fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # Обход всех строк второй таблицы и сравнение с каждой строкой первой таблицы
    rows1, cols1 = table1.shape
    rows2, cols2 = table2.shape

    for row2 in range(rows2):
        matched = False
        for row1 in range(rows1):
            equal = True
            for col in range(cols1):  # Предполагаем, что cols1 == cols2
                if table1.iat[row1, col] != table2.iat[row2, col]:
                    equal = False
                    break
            if equal:
                matched = True
                break
        
        # Если строка из table2 не совпала ни с одной строкой из table1, выделяем её
        if not matched:
            for col in range(cols2):
                sheet.cell(row=row2 + 2, column=col + 1).fill = fill

    # Генерация уникального имени файла для сохранения
    output_path = get_next_filename(output_path)

    # Сохранение файла
    workbook.save(output_path)
    print(f"Результаты сравнения сохранены в файл: {output_path}")

# Пример использования
compare_excel_tables('v1.xlsx', 'v2.xlsx', 'differences.xlsx')
