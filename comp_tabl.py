import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def compare_excel_tables(file1_path, file2_path, output_path):
    # Чтение таблиц из файлов Excel
    table1 = pd.read_excel(file1_path)
    table2 = pd.read_excel(file2_path)

    # Проверка наличия одинаковых столбцов
    if list(table1.columns) != list(table2.columns):
        raise ValueError("Таблицы должны иметь одинаковые столбцы для сравнения.")

    # Сравнение таблиц и создание маски различий
    comparison_values = table1 != table2
    rows, cols = comparison_values.shape

    # Загрузка второго файла для модификации
    workbook = load_workbook(file2_path)
    sheet = workbook.active

    # Цвет заливки для выделения различий
    fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # Обход всех ячеек и выделение различий
    for row in range(2, rows + 2):  # Начинаем с 2, так как первая строка - заголовок
        for col in range(1, cols + 1):
            if comparison_values.iat[row-2, col-1]:
                sheet.cell(row=row, column=col).fill = fill

    # Сохранение файла
    workbook.save(output_path)

# Пример использования
compare_excel_tables('v1.xlsx', 'v2.xlsx', 'differences.xlsx')
