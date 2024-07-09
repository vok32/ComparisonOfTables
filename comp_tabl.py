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
        print("Таблицы имеют разные столбцы.")
        print("Столбцы файла 1:", list(table1.columns))
        print("Столбцы файла 2:", list(table2.columns))
        proceed = input("Продолжить сравнение, исключив несовпадающие столбцы? (да/нет): ").strip().lower()
        if proceed != 'да':
            raise ValueError("Сравнение прервано пользователем.")
        else:
            common_columns = list(set(table1.columns) & set(table2.columns))
            table1 = table1[common_columns]
            table2 = table2[common_columns]
            print("Будут сравниваться только общие столбцы:", common_columns)

    # Спрашиваем у пользователя, по какому столбцу проводить сверку
    print("Доступные столбцы для сверки:", list(table1.columns))
    key_column = input("Введите название столбца для сверки: ").strip()
    if key_column not in table1.columns:
        raise ValueError("Указанный столбец отсутствует в таблицах.")

    # Загрузка второго файла для модификации
    workbook = load_workbook(file2_path)
    sheet = workbook.active

    # Цвет заливки для выделения различий
    fill_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    fill_light_green = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")

    # Создаем словарь для быстрого поиска строк в table1 по значению ключевого столбца
    table1_dict = table1.set_index(key_column).T.to_dict()

    # Обход всех строк второй таблицы и сравнение с соответствующими строками первой таблицы
    for index, row in table2.iterrows():
        key_value = row[key_column]
        if key_value in table1_dict:
            # Найдена строка с таким же ключевым значением, проверяем на изменения
            table1_row = table1_dict[key_value]
            differences = False
            for col_name in table2.columns:
                if col_name in table1_row and row[col_name] != table1_row[col_name]:
                    differences = True
                    sheet.cell(row=index + 2, column=table2.columns.get_loc(col_name) + 1).fill = fill_light_green
            if not differences:
                continue
        else:
            # Строка с таким ключевым значением отсутствует в table1, выделяем всю строку
            for col_index in range(len(row)):
                sheet.cell(row=index + 2, column=col_index + 1).fill = fill_yellow

    # Генерация уникального имени файла для сохранения
    output_path = get_next_filename(output_path)

    # Сохранение файла
    workbook.save(output_path)
    print(f"Результаты сравнения сохранены в файл: {output_path}")

# Пример использования
compare_excel_tables('v1.xlsx', 'v2.xlsx', 'differences.xlsx')
