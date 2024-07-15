import os
import pandas as pd
from tkinter import *
from tkinter import filedialog, messagebox, ttk, simpledialog  
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill

def get_next_filename(output_file):
    base_name, ext = os.path.splitext(output_file)
    version = 1
    while os.path.exists(output_file):
        version += 1
        output_file = f"{base_name}_v{version}{ext}"
    return output_file

def remove_empty_rows(sheet):
    rows_to_delete = []
    for row in sheet.iter_rows(min_row=2, max_col=sheet.max_column, values_only=False):
        if all(cell.value is None for cell in row):
            rows_to_delete.append(row[0].row)
    for row_num in reversed(rows_to_delete):
        sheet.delete_rows(row_num)

def custom_messagebox(title, message, root):
    window = Toplevel(root)
    window.title(title)

    # Центрирование окна в левом верхнем углу приложения
    window.geometry(f"400x125+{root.winfo_x()}+{root.winfo_y()}")  # Позиция в левом верхнем углу

    label = Label(window, text=message)
    label.pack(pady=10)

    accept_button = Button(window, text="Принять", command=window.destroy)
    accept_button.pack(pady=5)

    window.transient()  # Делает окно модальным
    window.grab_set()   # Блокирует родительское окно
    window.focus_set()  # Устанавливает фокус на новое окно
    window.wait_window()  # Ожидает закрытия окна

def compare_excel_tables(file1_path, file2_path, output_path, save_option, key_column, root, position):
    # Чтение таблиц из файлов Excel
    table1 = pd.read_excel(file1_path, engine='openpyxl')
    table2 = pd.read_excel(file2_path, engine='openpyxl')

    # Получаем имена столбцов
    columns1 = set(table1.columns)
    columns2 = set(table2.columns)

    # Проверка наличия одинаковых столбцов
    if columns1 != columns2:
        custom_messagebox("Ошибка", "Таблицы имеют разные столбцы.\nБудут использованы только общие столбцы.\n\nОтсылка на кнопку отклонить в 1C.", root)

    # Создаем новый файл для сохранения результатов
    new_workbook = Workbook()
    new_sheet = new_workbook.active
    new_sheet.title = "Результаты сравнения"

    # Записываем заголовки столбцов в том же порядке, как в table2
    for col_index, col_name in enumerate(table2.columns, start=1):
        new_sheet.cell(row=1, column=col_index).value = col_name

    # Цвет заливки для выделения различий
    fill_yellow = PatternFill(start_color="FFEB99", end_color="FFEB99", fill_type="solid")
    fill_light_green = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
    fill_green = PatternFill(start_color="77DD77", end_color="77DD77", fill_type="solid")
    fill_light_orange = PatternFill(start_color="FFDAB9", end_color="FFDAB9", fill_type="solid")

    # Создаем словарь для быстрого поиска строк в table1 по значению ключевого столбца
    table1_dict = table1.set_index(key_column).T.to_dict()

    # Обход всех строк второй таблицы и сравнение с соответствующими строками первой таблицы
    for index, row in table2.iterrows():
        key_value = row[key_column]
        if key_value in table1_dict:
            # Найдена строка с таким же ключевым значением, проверяем на изменения
            table1_row = table1_dict[key_value]
            differences = False
            for col_index, col_name in enumerate(table2.columns, start=1):
                if col_name in table1_row and row[col_name] != table1_row[col_name]:
                    differences = True
                    new_sheet.cell(row=index + 2, column=col_index).fill = fill_green
            if differences and (save_option in ["Все строки", "Только измененные строки", "Новые/измененные строки"]):
                for col_index, col_name in enumerate(table2.columns, start=1):
                    new_sheet.cell(row=index + 2, column=col_index).value = row[col_name]
                    if new_sheet.cell(row=index + 2, column=col_index).fill != fill_green:
                        new_sheet.cell(row=index + 2, column=col_index).fill = fill_light_green
        else:
            # Строка с таким ключевым значением отсутствует в table1, выделяем всю строку
            if save_option in ["Все строки", "Только новые строки", "Новые/измененные строки"]:
                for col_index, col_name in enumerate(table2.columns, start=1):
                    new_sheet.cell(row=index + 2, column=col_index).value = row[col_name]
                    new_sheet.cell(row=index + 2, column=col_index).fill = fill_yellow

    if save_option == "Все строки":
        for index, row in table2.iterrows():
            for col_index, col_name in enumerate(table2.columns, start=1):
                cell = new_sheet.cell(row=index + 2, column=col_index)
                if cell.value is None:
                    cell.value = row[col_name]

    # Удаление пустых строк
    remove_empty_rows(new_sheet)

    # Окрашиваем неучтенные столбцы в светло-оранжевый
    unused_columns1 = columns1 - columns2
    unused_columns2 = columns2 - columns1

    # Определяем фактическое количество строк в новом файле
    max_rows = new_sheet.max_row  # Получаем количество строк в new_sheet

    for col_name in unused_columns1:
        if col_name in table2.columns:
            continue
        col_index = table2.columns.get_loc(col_name) + 1
        for row in range(1, max_rows + 1):  # Включаем заголовок
            new_sheet.cell(row=row, column=col_index).fill = fill_light_orange

    for col_name in unused_columns2:
        col_index = table2.columns.get_loc(col_name) + 1
        for row in range(1, max_rows + 1):  # Включаем заголовок
            new_sheet.cell(row=row, column=col_index).fill = fill_light_orange

    # Генерация уникального имени файла для сохранения
    output_path = get_next_filename(output_path)

    # Сохранение файла
    new_workbook.save(output_path)

    # Отображение окна с сообщением об успешном сохранении
    show_success_window(output_path, root, position)

def show_success_window(output_path, root, position):
    success_window = Toplevel(root)
    success_window.title("Успех")
    
    root.update_idletasks()
    root_position_x = root.winfo_x()
    root_position_y = root.winfo_y()
    success_window.geometry(f"400x150+{root_position_x}+{root_position_y}")
    
    label = Label(success_window, text=f"Результаты сравнения сохранены в файл:\n{output_path}")
    label.pack(pady=10)
    
    open_folder_button = Button(success_window, text="Открыть папку с файлом", command=lambda: open_output_folder(output_path))
    open_folder_button.pack(pady=5)
    
    close_button = Button(success_window, text="Готово", command=success_window.destroy)
    close_button.pack(pady=5)
   
    success_window.transient(root)
    success_window.grab_set()
    success_window.focus_set()
    success_window.wait_window(success_window)

def open_output_folder(output_path):
    os.system(f'explorer /select,"{os.path.abspath(output_path)}"')

def select_files(root):
    position = [root.winfo_x(), root.winfo_y()]  # Получаем позицию главного окна

    def select_file1():
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if filename:
            file1_entry.delete(0, END)
            file1_entry.insert(0, filename)  # Сохраняем полный путь

    def select_file2():
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if filename:
            file2_entry.delete(0, END)
            file2_entry.insert(0, filename)  # Сохраняем полный путь

    def select_output_folder():
        foldername = filedialog.askdirectory()
        if foldername:
            output_entry.delete(0, END)
            # Установим полный путь с расширением .xlsx
            output_entry.insert(0, os.path.join(foldername, "differences.xlsx"))
        else:
            output_entry.delete(0, END)
            desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
            comparison_folder = os.path.join(desktop, "Сравнение таблиц")
            if not os.path.exists(comparison_folder):
                os.makedirs(comparison_folder)
            output_entry.insert(0, os.path.join(comparison_folder, "differences.xlsx"))

    def load_columns(file_path):
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
            return list(df.columns)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить столбцы из файла: {file_path}\nОшибка: {str(e)}")
            return []

    def show_columns_selection():
        file1_path = file1_entry.get()
        file2_path = file2_entry.get()
        output_path = output_entry.get()
        save_option = save_option_var.get()

        if not file1_path or not file2_path or not output_path:
            messagebox.showerror("Ошибка", "Не все поля были заполнены.")
            return

        columns_file1 = load_columns(file1_path)
        columns_file2 = load_columns(file2_path)

        if len(columns_file1) <= len(columns_file2):
            columns = columns_file1
            compare_file = "первого файла"
        else:
            columns = columns_file2
            compare_file = "второго файла"

        position = [root.winfo_x(), root.winfo_y()]  # Получаем позицию главного окна

        window = Toplevel(root)
        window.title("Выберите столбец для сравнения")
        window.geometry(f"400x150+{position[0]}+{position[1]}")

        label = Label(window, text=f"Выберите столбец для сравнения из {compare_file}:")
        label.pack(pady=10)

        key_column_var = StringVar(window)
        key_column_combo = ttk.Combobox(window, width=30, textvariable=key_column_var, values=columns)
        key_column_combo.pack(pady=10)

        def start_comparison():
            key_column = key_column_var.get().strip()
            if not key_column:
                messagebox.showerror("Ошибка", "Выберите столбец для сравнения.")
                return
            window.destroy()
            compare_excel_tables(file1_path, file2_path, output_path, save_option, key_column, root, position)

        button = Button(window, text="Начать сравнение", command=start_comparison)
        button.pack(pady=10)

    frame = Frame(root, padx=10, pady=10)
    frame.pack(padx=10, pady=10)

    file1_label = Label(frame, text="Выберите первый файл:")
    file1_label.grid(row=0, column=0, sticky=W)

    file1_entry = Entry(frame, width=50)
    file1_entry.grid(row=0, column=1, padx=5, pady=5)

    file1_button = Button(frame, text="Выбрать файл", command=select_file1)
    file1_button.grid(row=0, column=2, padx=5, pady=5)

    file2_label = Label(frame, text="Выберите второй файл:")
    file2_label.grid(row=1, column=0, sticky=W)

    file2_entry = Entry(frame, width=50)
    file2_entry.grid(row=1, column=1, padx=5, pady=5)

    file2_button = Button(frame, text="Выбрать файл", command=select_file2)
    file2_button.grid(row=1, column=2, padx=5, pady=5)

    output_label = Label(frame, text="Выберите папку для сохранения:")
    output_label.grid(row=2, column=0, sticky=W)

    output_entry = Entry(frame, width=50)
    output_entry.grid(row=2, column=1, padx=5, pady=5)

    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    comparison_folder = os.path.join(desktop, "Сравнение таблиц")
    if not os.path.exists(comparison_folder):
        os.makedirs(comparison_folder)
    output_entry.insert(0, os.path.join(comparison_folder, "differences.xlsx"))

    output_button = Button(frame, text="Выбрать папку", command=select_output_folder)
    output_button.grid(row=2, column=2, padx=5, pady=5)

    def update_filename():
        filename = simpledialog.askstring("Введите имя файла", "Имя файла:", parent=root)
        if filename:
            folder_path = os.path.dirname(output_entry.get())
            output_entry.delete(0, END)
            output_entry.insert(0, os.path.join(folder_path, f"{filename}.xlsx"))

    filename_button = Button(frame, text="Изменить имя файла", command=update_filename)
    filename_button.grid(row=3, column=2, padx=5, pady=5)

    save_label = Label(frame, text="Что сохранить в файле:")
    save_label.grid(row=4, column=0, sticky=W)

    save_option_var = StringVar(value="Все строки")

    save_radio1 = Radiobutton(frame, text="Все строки", variable=save_option_var, value="Все строки")
    save_radio1.grid(row=4, column=1, columnspan=2, padx=5, pady=5, sticky=W)

    save_radio2 = Radiobutton(frame, text="Только новые строки", variable=save_option_var, value="Только новые строки")
    save_radio2.grid(row=5, column=1, columnspan=2, padx=5, pady=5, sticky=W)

    save_radio3 = Radiobutton(frame, text="Только измененные строки", variable=save_option_var, value="Только измененные строки")
    save_radio3.grid(row=6, column=1, columnspan=2, padx=5, pady=5, sticky=W)

    save_radio4 = Radiobutton(frame, text="Новые+измененные строки", variable=save_option_var, value="Новые/измененные строки")
    save_radio4.grid(row=7, column=1, columnspan=2, padx=5, pady=5, sticky=W)

    start_button = Button(root, text="Далее", command=show_columns_selection, width=20)
    start_button.pack(pady=10, padx=10)

    developer_button = Button(root, text="О разработчике", command=lambda: show_developer_info(root, position), width=20)
    developer_button.pack(pady=10, padx=10)

    save_label = Label(frame, text="© 3МН")
    save_label.grid(row=8, column=2, sticky=E, pady=10)

# О разработчике
def show_developer_info(root, position):
    developer_window = Toplevel(root)
    developer_window.title("О разработчике")
    
    # Центрирование окна "О разработчике" относительно главного окна
    root.update_idletasks()
    root_position_x = root.winfo_x()
    root_position_y = root.winfo_y()
    developer_window.geometry(f"500x150+{root_position_x}+{root_position_y}")
    
    label = Label(developer_window, text="Программный продукт был разработан для облегчения Вашей работы", padx=10, pady=5)
    label.pack()

    label = Label(developer_window, text="Программа создана сотрудником 3 меганаправления, студентом 305 кафедры", padx=10, pady=5)
    label.pack()

    label = Label(developer_window, text="и просто хорошим человеком - Матюшенко Романом", padx=10, pady=5)
    label.pack()

    label = Label(developer_window, text="Ссылка на GitHub - https://github.com/vok32", padx=10, pady=5)
    label.pack()

    back_button = Button(developer_window, text="Назад", command=developer_window.destroy)
    back_button.pack()

def main():
    root = Tk()
    root.title("Сравнение таблиц Excel")
    root.geometry("700x475")

    # Открытие окна по центру экрана
    window_width = 700
    window_height = 475
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_top = int(screen_height / 2 - window_height / 2)
    position_right = int(screen_width / 2 - window_width / 2)
    root.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')
    
    select_files(root)
    root.mainloop()

if __name__ == "__main__":
    main()