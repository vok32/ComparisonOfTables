from __future__ import annotations

import os
import re
import subprocess
from pathlib import Path
from typing import Any, Hashable

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from tkinter import (
    BOTH,
    E,
    END,
    LEFT,
    W,
    Button,
    Entry,
    Frame,
    Label,
    Radiobutton,
    BooleanVar,
    StringVar,
    Tk,
    Toplevel,
    filedialog,
    messagebox,
    simpledialog,
    ttk,
)

APP_TITLE = "Сравнение таблиц Excel"
DEFAULT_OUTPUT_NAME = "differences.xlsx"
OUTPUT_SHEET_TITLE = "Результаты сравнения"

SAVE_ALL = "Все строки"
SAVE_NEW = "Новые строки"
SAVE_CHANGED = "Измененные строки"
SAVE_NEW_CHANGED = "Новые/измененные строки"
SAVE_MISSING = "Потеряшки"
SAVE_SUMMARY = "Для сводки"

SAVE_OPTIONS = (
    SAVE_ALL,
    SAVE_NEW,
    SAVE_CHANGED,
    SAVE_NEW_CHANGED,
    SAVE_MISSING,
    SAVE_SUMMARY,
)

FILL_YELLOW = PatternFill(start_color="FFEB99", end_color="FFEB99", fill_type="solid")
FILL_LIGHT_GREEN = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
FILL_GREEN = PatternFill(start_color="77DD77", end_color="77DD77", fill_type="solid")
FILL_LIGHT_ORANGE = PatternFill(start_color="FFDAB9", end_color="FFDAB9", fill_type="solid")
FILL_RED = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")


class ComparisonError(ValueError):
    """Понятная пользователю ошибка входных данных или сравнения."""


def get_next_filename(output_file: str | Path) -> Path:
    """Не перезаписывает существующий файл: name.xlsx -> name_v2.xlsx."""
    output_path = Path(output_file)
    if not output_path.exists():
        return output_path

    version = 2
    while True:
        candidate = output_path.with_name(
            f"{output_path.stem}_v{version}{output_path.suffix}"
        )
        if not candidate.exists():
            return candidate
        version += 1


def read_excel_table(file_path: str | Path) -> pd.DataFrame:
    """Читает первый лист Excel-файла и выдаёт понятные сообщения об ошибках."""
    path = Path(file_path)
    if not path.is_file():
        raise ComparisonError(f"Файл не найден:\n{path}")

    suffix = path.suffix.lower()
    if suffix not in {".xlsx", ".xlsm", ".xls"}:
        raise ComparisonError(
            "Поддерживаются файлы .xlsx, .xlsm и .xls."
        )

    try:
        if suffix == ".xls":
            # Для старого формата .xls pandas обычно использует пакет xlrd.
            return pd.read_excel(path)
        return pd.read_excel(path, engine="openpyxl")
    except ImportError as exc:
        if suffix == ".xls":
            raise ComparisonError(
                "Для чтения старого формата .xls требуется пакет xlrd. "
                "Проще всего пересохранить файл в формате .xlsx."
            ) from exc
        raise ComparisonError(f"Не установлен необходимый модуль: {exc}") from exc
    except Exception as exc:
        raise ComparisonError(f"Не удалось прочитать файл:\n{path}\n\n{exc}") from exc


def is_missing(value: Any) -> bool:
    """Безопасно определяет пустое значение pandas/Excel."""
    try:
        result = pd.isna(value)
    except (TypeError, ValueError):
        return False
    return bool(result) if isinstance(result, (bool, type(pd.NA))) or not hasattr(result, "__len__") else False


def values_equal(left: Any, right: Any) -> bool:
    """Считает две пустые ячейки равными."""
    left_missing = is_missing(left)
    right_missing = is_missing(right)
    if left_missing or right_missing:
        return left_missing and right_missing

    try:
        result = left == right
    except Exception:
        return False

    try:
        return bool(result)
    except (TypeError, ValueError):
        return False


def excel_value(value: Any) -> Any:
    """Преобразует NaN/NaT/pd.NA в пустую Excel-ячейку."""
    return None if is_missing(value) else value


def validate_key_column(
    table1: pd.DataFrame,
    table2: pd.DataFrame,
    key_column: Hashable,
) -> None:
    if key_column not in table1.columns or key_column not in table2.columns:
        raise ComparisonError(
            f"Ключевой столбец «{key_column}» должен присутствовать в обеих таблицах."
        )

    problems: list[str] = []
    for number, table in ((1, table1), (2, table2)):
        key_series = table[key_column]
        empty_count = int(key_series.isna().sum())
        duplicate_count = int(key_series.duplicated(keep=False).sum())

        if empty_count:
            problems.append(
                f"в таблице {number} пустых ключей: {empty_count}"
            )
        if duplicate_count:
            problems.append(
                f"в таблице {number} строк с повторяющимся ключом: {duplicate_count}"
            )

    if problems:
        raise ComparisonError(
            "Нельзя однозначно сопоставить строки по выбранному столбцу:\n• "
            + "\n• ".join(problems)
            + "\n\nВыберите уникальный столбец без пустых значений."
        )


def style_row(sheet, row_number: int, column_count: int, fill: PatternFill) -> None:
    for column_number in range(1, column_count + 1):
        sheet.cell(row=row_number, column=column_number).fill = fill


def format_sheet(sheet) -> None:
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    for column_cells in sheet.columns:
        max_length = 0
        for cell in column_cells:
            text = "" if cell.value is None else str(cell.value)
            max_length = max(max_length, len(text))
        width = min(max(max_length + 2, 10), 50)
        sheet.column_dimensions[get_column_letter(column_cells[0].column)].width = width


def compare_excel_tables(
    file1_path: str | Path,
    file2_path: str | Path,
    output_path: str | Path,
    save_option: str,
    key_column: Hashable,
    carry_columns: list[Hashable] | None = None,
) -> tuple[Path, list[Hashable], list[Hashable]]:
    """
    Сравнивает первую (старую) и вторую (новую) таблицы.

    Возвращает:
        путь сохранённого файла,
        столбцы только первой таблицы,
        столбцы только второй таблицы.
    """
    if save_option not in SAVE_OPTIONS:
        raise ComparisonError(f"Неизвестный режим сохранения: {save_option}")

    table1 = read_excel_table(file1_path)
    table2 = read_excel_table(file2_path)

    if table1.empty and len(table1.columns) == 0:
        raise ComparisonError("В первой таблице не обнаружены столбцы.")
    if table2.empty and len(table2.columns) == 0:
        raise ComparisonError("Во второй таблице не обнаружены столбцы.")

    validate_key_column(table1, table2, key_column)

    columns1 = list(table1.columns)
    columns2 = list(table2.columns)
    columns1_set = set(columns1)
    columns2_set = set(columns2)

    common_columns = [column for column in columns2 if column in columns1_set]
    only_table1 = [column for column in columns1 if column not in columns2_set]
    only_table2 = [column for column in columns2 if column not in columns1_set]

    selected_carry_columns = list(carry_columns or [])
    if save_option == SAVE_SUMMARY:
        if not selected_carry_columns:
            raise ComparisonError(
                "Для режима «Для сводки» выберите хотя бы один столбец "
                "из первой таблицы."
            )

        invalid_columns = [
            column
            for column in selected_carry_columns
            if column not in only_table1
        ]
        if invalid_columns:
            raise ComparisonError(
                "В сводку можно переносить только столбцы, которые есть "
                "в первой таблице и отсутствуют во второй:\n• "
                + "\n• ".join(map(str, invalid_columns))
            )

        # Не допускаем повторов, сохраняя выбранный пользователем порядок.
        selected_carry_columns = list(dict.fromkeys(selected_carry_columns))

    table1_by_key = table1.set_index(key_column, drop=False)
    table2_by_key = table2.set_index(key_column, drop=False)
    keys1 = set(table1_by_key.index)
    keys2 = set(table2_by_key.index)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = OUTPUT_SHEET_TITLE

    if save_option == SAVE_MISSING:
        output_columns = columns1
        unused_columns = only_table1
    elif save_option == SAVE_SUMMARY:
        output_columns = columns2 + selected_carry_columns
        unused_columns = only_table2
    else:
        output_columns = columns2
        unused_columns = only_table2

    sheet.append([excel_value(column) for column in output_columns])

    def append_row(
        row_values: pd.Series | dict[Hashable, Any],
        status: str,
        changed_columns: set[Hashable] | None = None,
    ) -> None:
        sheet.append([excel_value(row_values.get(column)) for column in output_columns])
        output_row = sheet.max_row

        if status == "new":
            style_row(sheet, output_row, len(output_columns), FILL_YELLOW)
        elif status == "changed":
            style_row(sheet, output_row, len(output_columns), FILL_LIGHT_GREEN)
            for column in changed_columns or set():
                if column in output_columns:
                    column_number = output_columns.index(column) + 1
                    sheet.cell(row=output_row, column=column_number).fill = FILL_GREEN
        elif status == "missing":
            style_row(sheet, output_row, len(output_columns), FILL_RED)

    if save_option == SAVE_MISSING:
        for _, row in table1.iterrows():
            if row[key_column] not in keys2:
                append_row(row, "missing")
    else:
        for _, row in table2.iterrows():
            key_value = row[key_column]

            if key_value not in keys1:
                if save_option in {
                    SAVE_ALL,
                    SAVE_NEW,
                    SAVE_NEW_CHANGED,
                    SAVE_SUMMARY,
                }:
                    if save_option == SAVE_SUMMARY:
                        summary_row = row.to_dict()
                        summary_row.update(
                            {column: None for column in selected_carry_columns}
                        )
                        append_row(summary_row, "new")
                    else:
                        append_row(row, "new")
                continue

            old_row = table1_by_key.loc[key_value]
            changed_columns = {
                column
                for column in common_columns
                if not values_equal(old_row[column], row[column])
            }

            if save_option == SAVE_SUMMARY:
                summary_row = row.to_dict()
                summary_row.update(
                    {
                        column: old_row[column]
                        for column in selected_carry_columns
                    }
                )
                append_row(
                    summary_row,
                    "changed" if changed_columns else "unchanged",
                    changed_columns,
                )
            elif changed_columns:
                if save_option in {SAVE_ALL, SAVE_CHANGED, SAVE_NEW_CHANGED}:
                    append_row(row, "changed", changed_columns)
            elif save_option == SAVE_ALL:
                append_row(row, "unchanged")

    # Столбцы, которых нет в другой таблице, не участвовали в сравнении.
    for column in unused_columns:
        column_number = output_columns.index(column) + 1
        for row_number in range(1, sheet.max_row + 1):
            sheet.cell(row=row_number, column=column_number).fill = FILL_LIGHT_ORANGE

    # В режиме сводки оранжевым отмечаем только заголовки ручных полей.
    # Так сохраняется цветовой статус новых и изменённых строк.
    if save_option == SAVE_SUMMARY:
        for column in selected_carry_columns:
            column_number = output_columns.index(column) + 1
            sheet.cell(row=1, column=column_number).fill = FILL_LIGHT_ORANGE

    format_sheet(sheet)

    final_output_path = get_next_filename(output_path)
    final_output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(final_output_path)

    return final_output_path, only_table1, only_table2


def center_window(
    window: Tk | Toplevel,
    parent: Tk | Toplevel | None = None,
) -> None:
    """Центрирует окно по экрану или относительно родительского окна."""
    window.update_idletasks()
    width = window.winfo_reqwidth()
    height = window.winfo_reqheight()

    if parent is None:
        x = (window.winfo_screenwidth() - width) // 2
        y = (window.winfo_screenheight() - height) // 2
    else:
        parent.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - width) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - height) // 2

    # Размер остаётся автоматическим: задаём только положение окна.
    window.geometry(f"+{max(0, x)}+{max(0, y)}")


def center_after_layout(
    window: Tk | Toplevel,
    parent: Tk | Toplevel | None = None,
) -> None:
    """Центрирует окно после расчёта размеров всех элементов интерфейса."""
    window.after_idle(lambda: center_window(window, parent))


def custom_messagebox(title: str, message: str, root: Tk) -> None:
    window = Toplevel(root)
    window.title(title)

    Label(window, text=message, justify="left", wraplength=420).pack(pady=12, padx=12)
    Button(window, text="Принять", command=window.destroy).pack(pady=5)

    window.transient(root)
    center_after_layout(window, root)
    window.grab_set()
    window.focus_set()
    window.wait_window()


def show_success_window(output_path: Path, root: Tk) -> None:
    success_window = Toplevel(root)
    success_window.title("Успех")

    Label(
        success_window,
        text=f"Результаты сравнения сохранены в файл:\n{output_path}",
        wraplength=470,
    ).pack(pady=10, padx=10)

    Button(
        success_window,
        text="Открыть папку с файлом",
        command=lambda: open_output_folder(output_path),
    ).pack(pady=5)
    Button(success_window, text="Готово", command=success_window.destroy).pack(pady=5)

    success_window.transient(root)
    center_after_layout(success_window, root)
    success_window.grab_set()
    success_window.focus_set()
    success_window.wait_window()


def open_output_folder(output_path: str | Path) -> None:
    absolute_path = str(Path(output_path).resolve())
    if os.name == "nt":
        subprocess.run(["explorer", "/select,", absolute_path], check=False)
    else:
        subprocess.run(["xdg-open", str(Path(absolute_path).parent)], check=False)


def default_output_path() -> Path:
    desktop = Path.home() / "Desktop"
    base_folder = desktop if desktop.exists() else Path.home()
    comparison_folder = base_folder / "Сравнение таблиц"
    comparison_folder.mkdir(parents=True, exist_ok=True)
    return comparison_folder / DEFAULT_OUTPUT_NAME


def sanitize_filename(filename: str) -> str:
    filename = filename.strip()
    if filename.lower().endswith(".xlsx"):
        filename = filename[:-5]
    filename = re.sub(r'[<>:"/\\|?*]', "_", filename).strip(" .")
    if not filename:
        raise ComparisonError("Имя файла не может быть пустым.")
    return f"{filename}.xlsx"


def select_files(root: Tk) -> None:
    def select_file(entry: Entry) -> None:
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")]
        )
        if filename:
            entry.delete(0, END)
            entry.insert(0, filename)

    def select_output_folder() -> None:
        foldername = filedialog.askdirectory()
        if foldername:
            output_entry.delete(0, END)
            output_entry.insert(0, str(Path(foldername) / DEFAULT_OUTPUT_NAME))

    def load_columns(file_path: str) -> list[Hashable] | None:
        try:
            return list(read_excel_table(file_path).columns)
        except ComparisonError as exc:
            messagebox.showerror("Ошибка", str(exc), parent=root)
            return None

    def show_columns_selection() -> None:
        file1_path = file1_entry.get().strip()
        file2_path = file2_entry.get().strip()
        output_path = output_entry.get().strip()
        save_option = save_option_var.get()

        if not file1_path or not file2_path or not output_path:
            messagebox.showerror("Ошибка", "Не все поля были заполнены.", parent=root)
            return

        columns_file1 = load_columns(file1_path)
        columns_file2 = load_columns(file2_path)
        if columns_file1 is None or columns_file2 is None:
            return

        columns2_set = set(columns_file2)
        common_columns = [column for column in columns_file1 if column in columns2_set]
        only_file1_columns = [
            column for column in columns_file1 if column not in columns2_set
        ]
        if not common_columns:
            messagebox.showerror(
                "Ошибка",
                "В таблицах нет ни одного общего столбца для выбора ключа.",
                parent=root,
            )
            return

        if save_option == SAVE_SUMMARY and not only_file1_columns:
            messagebox.showerror(
                "Ошибка",
                "В первой таблице нет дополнительных столбцов, которых нет "
                "во второй. Для такого набора файлов используйте режим "
                "«Все строки».",
                parent=root,
            )
            return

        window = Toplevel(root)
        window.title("Настройка сравнения")
        window.transient(root)

        Label(
            window,
            text="Выберите общий уникальный столбец без пустых значений:",
        ).pack(pady=(12, 6))

        key_column_combo = ttk.Combobox(
            window,
            width=38,
            values=[str(column) for column in common_columns],
            state="readonly",
        )
        key_column_combo.pack(pady=10)
        if common_columns:
            key_column_combo.current(0)

        carry_column_vars: list[tuple[Hashable, BooleanVar]] = []
        if save_option == SAVE_SUMMARY:
            Label(
                window,
                text=(
                    "Выберите поля из первой таблицы, которые нужно перенести "
                    "в сводку.\nДля новых строк эти ячейки останутся пустыми."
                ),
                justify="left",
                wraplength=460,
            ).pack(pady=(14, 6), padx=12)

            carry_columns_frame = ttk.LabelFrame(
                window,
                text="Столбцы для переноса",
                padding=(10, 6),
            )
            carry_columns_frame.pack(fill="x", pady=4, padx=18)

            for column in only_file1_columns:
                selected_var = BooleanVar(value=True)
                ttk.Checkbutton(
                    carry_columns_frame,
                    text=str(column),
                    variable=selected_var,
                ).pack(anchor="w", pady=2)
                carry_column_vars.append((column, selected_var))

        def start_comparison() -> None:
            selected_index = key_column_combo.current()
            if selected_index < 0:
                messagebox.showerror(
                    "Ошибка", "Выберите столбец для сравнения.", parent=window
                )
                return

            key_column = common_columns[selected_index]

            selected_carry_columns: list[Hashable] = []
            if save_option == SAVE_SUMMARY:
                selected_carry_columns = [
                    column
                    for column, selected_var in carry_column_vars
                    if selected_var.get()
                ]
                if not selected_carry_columns:
                    messagebox.showerror(
                        "Ошибка",
                        "Поставьте галочку хотя бы у одного столбца для переноса в сводку.",
                        parent=window,
                    )
                    return

            root.config(cursor="watch")
            window.config(cursor="watch")
            root.update_idletasks()

            try:
                saved_path, only_table1, only_table2 = compare_excel_tables(
                    file1_path,
                    file2_path,
                    output_path,
                    save_option,
                    key_column,
                    selected_carry_columns,
                )
            except ComparisonError as exc:
                messagebox.showerror("Ошибка", str(exc), parent=window)
                return
            except Exception as exc:
                messagebox.showerror(
                    "Непредвиденная ошибка",
                    f"Сравнение не выполнено:\n{exc}",
                    parent=window,
                )
                return
            finally:
                root.config(cursor="")
                window.config(cursor="")

            window.destroy()

            if only_table1 or only_table2:
                details: list[str] = [
                    "Таблицы имеют разный состав столбцов. Сравнивались только общие столбцы."
                ]
                if only_table1:
                    details.append(
                        "Только в первой: " + ", ".join(map(str, only_table1))
                    )
                if only_table2:
                    details.append(
                        "Только во второй: " + ", ".join(map(str, only_table2))
                    )
                custom_messagebox("Обратите внимание", "\n".join(details), root)

            show_success_window(saved_path, root)

        Button(window, text="Создать файл", command=start_comparison).pack(pady=12)
        center_after_layout(window, root)

    frame = Frame(root, padx=10, pady=10)
    frame.pack(padx=10, pady=10, fill=BOTH)

    Label(frame, text="Выберите первый файл (исходная таблица):").grid(
        row=0, column=0, sticky=W
    )
    file1_entry = Entry(frame, width=50)
    file1_entry.grid(row=0, column=1, padx=5, pady=5)
    Button(frame, text="Выбрать файл", command=lambda: select_file(file1_entry)).grid(
        row=0, column=2, padx=5, pady=5
    )

    Label(frame, text="Выберите второй файл (новая таблица):").grid(
        row=1, column=0, sticky=W
    )
    file2_entry = Entry(frame, width=50)
    file2_entry.grid(row=1, column=1, padx=5, pady=5)
    Button(frame, text="Выбрать файл", command=lambda: select_file(file2_entry)).grid(
        row=1, column=2, padx=5, pady=5
    )

    Label(frame, text="Выберите папку для сохранения:").grid(
        row=2, column=0, sticky=W
    )
    output_entry = Entry(frame, width=50)
    output_entry.grid(row=2, column=1, padx=5, pady=5)
    try:
        output_entry.insert(0, str(default_output_path()))
    except OSError:
        output_entry.insert(0, str(Path.home() / DEFAULT_OUTPUT_NAME))

    Button(frame, text="Выбрать папку", command=select_output_folder).grid(
        row=2, column=2, padx=5, pady=5
    )

    def update_filename() -> None:
        filename = simpledialog.askstring(
            "Введите имя файла", "Имя файла:", parent=root
        )
        if filename is None:
            return
        try:
            safe_filename = sanitize_filename(filename)
        except ComparisonError as exc:
            messagebox.showerror("Ошибка", str(exc), parent=root)
            return

        folder_path = Path(output_entry.get()).parent
        output_entry.delete(0, END)
        output_entry.insert(0, str(folder_path / safe_filename))

    Button(frame, text="Изменить имя файла", command=update_filename).grid(
        row=3, column=2, padx=5, pady=5
    )

    Label(frame, text="Что сохранить в файле:").grid(row=4, column=0, sticky=W)
    save_option_var = StringVar(value=SAVE_ALL)

    for row_number, option in enumerate(SAVE_OPTIONS, start=4):
        if option == SAVE_MISSING:
            label = "Потеряшки"
        elif option == SAVE_SUMMARY:
            label = "Для сводки (все + комментарии)"
        else:
            label = option.replace("Новые/измененные", "Новые + изменённые")
        Radiobutton(
            frame,
            text=label,
            variable=save_option_var,
            value=option,
        ).grid(
            row=row_number,
            column=1,
            columnspan=2,
            padx=5,
            pady=5,
            sticky=W,
        )

    Button(root, text="Далее", command=show_columns_selection, width=20).pack(
        pady=10, padx=10
    )

    button_frame = Frame(root)
    button_frame.pack(pady=10)
    Button(
        button_frame,
        text="О разработчике",
        command=lambda: show_developer_info(root),
        width=20,
    ).pack(padx=10, side=LEFT)
    Button(
        button_frame,
        text="О приложении",
        command=lambda: show_app_info(root),
        width=20,
    ).pack(padx=10, side=LEFT)

    Label(frame, text="© 3МН").grid(
        row=4 + len(SAVE_OPTIONS), column=2, sticky=E, pady=10
    )


def show_developer_info(root: Tk) -> None:
    developer_window = Toplevel(root)
    developer_window.title("О разработчике")
    developer_window.transient(root)

    Label(
        developer_window,
        text="Программный продукт был разработан для облегчения Вашей работы",
        padx=10,
        pady=5,
    ).pack()
    Label(
        developer_window,
        text="Программа создана сотрудником 3 меганаправления, студентом 305 кафедры",
        padx=10,
        pady=5,
    ).pack()
    Label(
        developer_window,
        text="в далёком 2024 году.",
        padx=10,
        pady=5,
    ).pack()
    Label(
        developer_window,
        text="GitHub: https://github.com/vok32",
        padx=10,
        pady=5,
    ).pack()
    Button(developer_window, text="Назад", command=developer_window.destroy).pack()
    center_after_layout(developer_window, root)


def show_app_info(root: Tk) -> None:
    info_message = (
        "Цвета выделения:\n"
        "    Новые строки: жёлтый\n"
        "    Изменённая строка: светло-зелёный\n"
        "    Изменённая ячейка: зелёный\n"
        "    Строки из первой таблицы, отсутствующие во второй: красный\n"
        "    Столбцы, не участвовавшие в сравнении: светло-оранжевый\n"
        "    Заголовки ручных полей сводки: светло-оранжевый"
    )

    info_window = Toplevel(root)
    info_window.title("Информация о приложении")
    info_window.transient(root)

    Label(info_window, text=info_message, justify="left").pack(pady=10, padx=10)
    Button(info_window, text="Закрыть", command=info_window.destroy).pack(pady=10)
    center_after_layout(info_window, root)


def main() -> None:
    root = Tk()
    root.title(APP_TITLE)

    select_files(root)
    center_after_layout(root)
    root.mainloop()


if __name__ == "__main__":
    main()
