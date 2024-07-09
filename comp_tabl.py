import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.cell import get_column_letter
import os

def get_next_filename(output_file):
    base_name, ext = os.path.splitext(output_file)
    version = 1
    while os.path.exists(output_file):
        version += 1
        output_file = f"{base_name}_v{version}{ext}"
    return output_file

def compare_excel_tables(file1_path, file2_path, output_path, save_all_rows):
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

    # Создаем новый файл для сохранения результатов
    new_workbook = Workbook()
    new_sheet = new_workbook.active
    new_sheet.append(list(table2.columns))

    # Цвет заливки для выделения различий
    fill_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    fill_light_green = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
    fill_green = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")

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
                    new_sheet.cell(row=index + 2, column=table2.columns.get_loc(col_name) + 1).fill = fill_green
            if differences:
                for col_index in range(len(row)):
                    new_sheet.cell(row=index + 2, column=col_index + 1, value=row[col_index])
                    if new_sheet.cell(row=index + 2, column=col_index + 1).fill != fill_green:
                        new_sheet.cell(row=index + 2, column=col_index + 1).fill = fill_light_green
        else:
            # Строка с таким ключевым значением отсутствует в table1, выделяем всю строку
            for col_index in range(len(row)):
                new_sheet.cell(row=index + 2, column=col_index + 1, value=row[col_index])
                new_sheet.cell(row=index + 2, column=col_index + 1).fill = fill_yellow

    if save_all_rows:
        for index, row in table2.iterrows():
            for col_index in range(len(row)):
                cell = new_sheet.cell(row=index + 2, column=col_index + 1)
                if cell.value is None:
                    cell.value = row[col_index]

    # Генерация уникального имени файла для сохранения
    output_path = get_next_filename(output_path)

    # Сохранение файла
    new_workbook.save(output_path)
    print(f"Результаты сравнения сохранены в файл: {output_path}")

    # Удаление пустых строк
    wb = load_workbook(output_path)
    sheet = wb.active
    max_row = sheet.max_row
    max_col = sheet.max_column

    rows_to_delete = []
    for row in range(2, max_row + 1):
        row_empty = True
        for col in range(1, max_col + 1):
            if sheet.cell(row=row, column=col).value:
                row_empty = False
                break
        if row_empty:
            rows_to_delete.append(row)

    for row in reversed(rows_to_delete):
        sheet.delete_rows(row)

    wb.save(output_path)
    print(f"Удалены пустые строки из файла: {output_path}")

# Пример использования
print("Сохранить все строки или только новые и измененные? (все/только измененные)")
save_all = input().strip().lower() == "все"
compare_excel_tables('v1.xlsx', 'v2.xlsx', 'differences.xlsx', save_all)
