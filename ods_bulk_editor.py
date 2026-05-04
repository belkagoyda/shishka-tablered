"""
ODS Bulk Editor
===============
Десктоп-приложение на Python/tkinter для массового редактирования ODS-файлов.

Пользователь открывает один файл-образец из группы однотипных таблиц,
вносит правки (значения ячеек, шрифт, цвет текста), после чего может
применить те же правки ко всем файлам группы одним нажатием кнопки.

Группировка файлов производится по количеству строк с данными на первом листе.

Зависимости: python3, python3-tk, python3-odf (odfpy)
Запуск: python3 ods_bulk_editor.py
"""

from __future__ import annotations

# ---------------------------------------------------------------------------
# Стандартная библиотека
# ---------------------------------------------------------------------------
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from tkinter.scrolledtext import ScrolledText
from pathlib import Path
from dataclasses import dataclass, field
from typing import Callable, Iterable

# ---------------------------------------------------------------------------
# Сторонние библиотеки (odfpy / python3-odf)
# ---------------------------------------------------------------------------
from odf.opendocument import load as odf_load, OpenDocumentSpreadsheet
from odf.table import Table, TableRow, TableCell
from odf.text import P
from odf import style, namespaces
from odf.style import Style, TextProperties, TableCellProperties
from odf.namespaces import OFFICENS, STYLENS, FONS


# ===========================================================================
# Модели данных
# ===========================================================================

@dataclass
class FileInfo:
    """Информация об ODS-файле: путь и количество строк с данными."""
    path: Path
    row_count: int          # количество строк с данными на первом листе


@dataclass
class CellData:
    """Данные одной ячейки: значение и параметры форматирования текста."""
    value: str              # строковое представление значения ячейки
    font_name: str
    font_size: int
    bold: bool
    italic: bool
    color: str              # hex "#RRGGBB"; "" если не задан


@dataclass
class SheetData:
    """Данные первого листа ODS-файла."""
    cells: dict[tuple[int, int], CellData]   # (row, col) → CellData
    max_row: int
    max_col: int


@dataclass
class CellDelta:
    """Изменения одной ячейки. None означает «значение не менялось»."""
    value: str | None          # None — значение не менялось
    font_name: str | None
    font_size: int | None
    bold: bool | None
    italic: bool | None
    color: str | None          # hex-строка вида "#RRGGBB" или None


@dataclass
class BulkResult:
    """Результат массового применения дельты к группе файлов."""
    success_count: int
    error_count: int
    errors: list[str]       # описания ошибок для лога


@dataclass
class AppState:
    """Глобальное состояние приложения, передаваемое между компонентами."""
    target_folder: Path
    scan_results: list[FileInfo]          # результат последнего сканирования
    groups: dict[int, list[FileInfo]]     # группировка
    selected_group_key: int | None        # выбранная группа
    current_file: FileInfo | None         # открытый образец
    sheet_data: SheetData | None          # данные образца
    delta: "DeltaStore"
    log: "LogStore"


# ===========================================================================
# LogStore — хранилище лога сессии
# ===========================================================================

class LogStore:
    """Хранит записи лога в памяти (без записи на диск)."""

    def __init__(self) -> None:
        self._entries: list[str] = []

    def add(self, message: str) -> None:
        """Добавить запись в лог."""
        self._entries.append(message)

    def get_all(self) -> list[str]:
        """Вернуть все записи лога."""
        return list(self._entries)

    def clear(self) -> None:
        """Очистить лог."""
        self._entries.clear()


# ===========================================================================
# DeltaStore — хранилище изменений образца в памяти
# ===========================================================================

class DeltaStore:
    """Хранит изменения ячеек образца в памяти до явного сохранения."""

    def __init__(self) -> None:
        self._data: dict[tuple[int, int], CellDelta] = {}

    def set(self, row: int, col: int, delta: CellDelta) -> None:
        """Записать дельту для ячейки (row, col)."""
        self._data[(row, col)] = delta

    def get(self, row: int, col: int) -> CellDelta | None:
        """Получить дельту для ячейки (row, col) или None, если изменений нет."""
        return self._data.get((row, col))

    def items(self) -> Iterable[tuple[tuple[int, int], CellDelta]]:
        """Итерация по всем изменённым ячейкам: ((row, col), CellDelta)."""
        return self._data.items()

    def is_empty(self) -> bool:
        """Вернуть True, если дельта не содержит изменений."""
        return len(self._data) == 0

    def clear(self) -> None:
        """Очистить все изменения."""
        self._data.clear()


# ===========================================================================
# Scanner — поиск ODS-файлов и подсчёт строк с данными
# ===========================================================================

def count_data_rows(sheet) -> int:
    """
    Подсчитывает строки с данными на листе odfpy.
    Игнорирует trailing empty rows (строки в конце, где все ячейки пусты).
    """
    rows = sheet.getElementsByType(TableRow)
    if not rows:
        return 0

    def is_row_empty(row) -> bool:
        cells = row.getElementsByType(TableCell)
        for cell in cells:
            # Проверяем атрибут value (числа, даты)
            if cell.getAttribute("value") is not None:
                return False
            # Проверяем текстовое содержимое
            paragraphs = cell.getElementsByType(P)
            for p in paragraphs:
                text = str(p)
                if text.strip():
                    return False
        return True

    # Строим список флагов пустоты для каждой строки
    emptiness = [is_row_empty(row) for row in rows]

    # Отрезаем trailing empty rows
    last_data_index = -1
    for i in range(len(emptiness) - 1, -1, -1):
        if not emptiness[i]:
            last_data_index = i
            break

    if last_data_index == -1:
        return 0

    # Считаем строки с данными (не пустые) среди первых last_data_index+1 строк
    return sum(1 for i in range(last_data_index + 1) if not emptiness[i])


class Scanner:
    """Ищет ODS-файлы в папке и подсчитывает строки с данными."""

    def scan(self, folder: Path, log: LogStore) -> list[FileInfo]:
        """
        Возвращает список FileInfo для всех .ods-файлов в folder (не рекурсивно).
        Повреждённые файлы пропускаются; ошибки добавляются в лог.
        """
        results: list[FileInfo] = []

        try:
            entries = list(folder.iterdir())
        except OSError as e:
            log.add(f"Ошибка доступа к папке {folder}: {e}")
            return results

        for entry in entries:
            if not entry.is_file():
                continue
            if entry.suffix.lower() != ".ods":
                continue

            try:
                doc = odf_load(str(entry))
                sheets = doc.spreadsheet.getElementsByType(Table)
                if not sheets:
                    row_count = 0
                else:
                    row_count = count_data_rows(sheets[0])
                results.append(FileInfo(path=entry, row_count=row_count))
            except Exception as e:
                log.add(f"Ошибка чтения файла {entry.name}: {e}")

        return results


# ===========================================================================
# GroupManager — группировка файлов по количеству строк
# ===========================================================================

class GroupManager:
    """Формирует и хранит группировку файлов по количеству строк с данными."""

    def __init__(self) -> None:
        self._groups: dict[int, list[FileInfo]] = {}

    def build(self, files: list[FileInfo]) -> None:
        """Строит словарь {row_count: [FileInfo, ...]}."""
        self._groups = {}
        for file_info in files:
            key = file_info.row_count
            if key not in self._groups:
                self._groups[key] = []
            self._groups[key].append(file_info)

    def get_groups(self) -> dict[int, list[FileInfo]]:
        """Возвращает текущую группировку, отсортированную по ключу."""
        return dict(sorted(self._groups.items()))

    def get_group(self, row_count: int) -> list[FileInfo]:
        """Возвращает список файлов для конкретной группы."""
        return self._groups.get(row_count, [])


# ===========================================================================
# OdsReader — чтение содержимого первого листа ODS-файла
# ===========================================================================

class OdsReader:
    """Читает содержимое первого листа ODS-файла."""

    def load(self, path: Path) -> SheetData:
        """
        Загружает первый лист. Возвращает SheetData с матрицей ячеек.
        Поддерживает: строки, числа, даты, формулы (значение формулы).
        """
        doc = odf_load(str(path))
        sheets = doc.spreadsheet.getElementsByType(Table)
        if not sheets:
            return SheetData(cells={}, max_row=0, max_col=0)

        sheet = sheets[0]
        cells: dict[tuple[int, int], CellData] = {}
        max_row = 0
        max_col = 0

        rows = sheet.getElementsByType(TableRow)
        for row_idx, row in enumerate(rows):
            row_cells = row.getElementsByType(TableCell)
            col_idx = 0
            for cell in row_cells:
                # Handle repeated columns
                repeat = cell.getAttribute("numbercolumnsrepeated")
                repeat_count = int(repeat) if repeat else 1

                value = self._get_cell_value(cell)
                fmt = self._get_cell_formatting(cell, doc)

                cell_data = CellData(
                    value=value,
                    font_name=fmt["font_name"],
                    font_size=fmt["font_size"],
                    bold=fmt["bold"],
                    italic=fmt["italic"],
                    color=fmt["color"],
                )

                for _ in range(repeat_count):
                    if value or repeat_count == 1:
                        cells[(row_idx, col_idx)] = cell_data
                        if row_idx > max_row:
                            max_row = row_idx
                        if col_idx > max_col:
                            max_col = col_idx
                    col_idx += 1

        return SheetData(cells=cells, max_row=max_row, max_col=max_col)

    def _get_cell_value(self, cell) -> str:
        """Извлекает строковое значение ячейки."""
        # Check for formula — use cached value
        formula = cell.getAttrNS(namespaces.TABLENS, "formula")
        value_type = cell.getAttrNS(OFFICENS, "value-type")

        if value_type == "float":
            val = cell.getAttrNS(OFFICENS, "value")
            if val is not None:
                # Return as integer if no fractional part
                try:
                    f = float(val)
                    if f == int(f):
                        return str(int(f))
                    return str(f)
                except (ValueError, OverflowError):
                    return str(val)

        if value_type == "date":
            val = cell.getAttrNS(OFFICENS, "date-value")
            if val is not None:
                return str(val)

        if value_type == "boolean":
            val = cell.getAttrNS(OFFICENS, "boolean-value")
            if val is not None:
                return str(val)

        # For string type or formula cached value — read text content
        paragraphs = cell.getElementsByType(P)
        text_parts = []
        for p in paragraphs:
            text_parts.append(str(p))
        return "".join(text_parts)

    def _get_cell_formatting(self, cell, doc) -> dict:
        """Извлекает параметры форматирования ячейки."""
        defaults = {
            "font_name": "",
            "font_size": 11,
            "bold": False,
            "italic": False,
            "color": "",
        }

        # Get style name from cell's paragraph
        paragraphs = cell.getElementsByType(P)
        style_name = None
        for p in paragraphs:
            sn = p.getAttrNS(namespaces.TEXTNS, "style-name")
            if sn:
                style_name = sn
                break

        if not style_name:
            # Try cell's own style
            style_name = cell.getAttrNS(namespaces.TABLENS, "style-name")

        if not style_name:
            return defaults

        return self._lookup_style(style_name, doc, defaults)

    def _lookup_style(self, style_name: str, doc, defaults: dict) -> dict:
        """Ищет стиль по имени и извлекает параметры форматирования."""
        result = dict(defaults)

        # Search in automatic styles
        all_styles = doc.automaticstyles.getElementsByType(Style)
        # Also search in named styles
        named_styles = doc.styles.getElementsByType(Style)
        all_style_list = list(all_styles) + list(named_styles)

        for s in all_style_list:
            if s.getAttrNS(STYLENS, "name") == style_name:
                text_props = s.getElementsByType(TextProperties)
                if text_props:
                    tp = text_props[0]
                    fn = tp.getAttrNS(STYLENS, "font-name")
                    if fn:
                        result["font_name"] = fn
                    fs = tp.getAttrNS(FONS, "font-size")
                    if fs:
                        # font-size may be like "12pt"
                        try:
                            result["font_size"] = int(float(str(fs).replace("pt", "").strip()))
                        except (ValueError, AttributeError):
                            pass
                    fw = tp.getAttrNS(FONS, "font-weight")
                    if fw:
                        result["bold"] = (fw == "bold")
                    fi = tp.getAttrNS(FONS, "font-style")
                    if fi:
                        result["italic"] = (fi == "italic")
                    color = tp.getAttrNS(FONS, "color")
                    if color:
                        result["color"] = color
                break

        return result


# ===========================================================================
# Processor — применение DeltaStore к ODS-файлам
# ===========================================================================

class Processor:
    """Применяет DeltaStore к ODS-файлам и сохраняет результат."""

    def apply_and_save(
        self,
        path: Path,
        delta: DeltaStore,
        log: LogStore,
    ) -> bool:
        """
        Загружает path, применяет delta к первому листу, сохраняет.
        Возвращает True при успехе, False при ошибке (ошибка пишется в log).
        """
        try:
            doc = odf_load(str(path))
            sheets = doc.spreadsheet.getElementsByType(Table)
            if not sheets:
                log.add(f"Файл {path.name}: нет листов")
                return False

            sheet = sheets[0]
            rows = sheet.getElementsByType(TableRow)

            for (row_idx, col_idx), cell_delta in delta.items():
                if row_idx >= len(rows):
                    continue
                row = rows[row_idx]
                row_cells = row.getElementsByType(TableCell)
                if col_idx >= len(row_cells):
                    continue
                cell = row_cells[col_idx]

                self._apply_delta_to_cell(cell, cell_delta, doc)

            doc.save(str(path))
            return True
        except Exception as e:
            log.add(f"Ошибка обработки файла {path.name}: {e}")
            return False

    def _apply_delta_to_cell(self, cell, cell_delta: CellDelta, doc) -> None:
        """Применяет изменения дельты к ячейке."""
        # Apply value change
        if cell_delta.value is not None:
            # Remove formula attribute if present
            try:
                cell.removeAttribute("formula")
            except Exception:
                pass
            # Remove existing paragraphs
            for p in list(cell.getElementsByType(P)):
                cell.removeChild(p)
            # Remove value-type and value attributes (convert to string)
            try:
                cell.removeAttribute("valuetype")
            except Exception:
                pass
            try:
                cell.removeAttribute("value")
            except Exception:
                pass
            try:
                cell.removeAttribute("datevalue")
            except Exception:
                pass
            # Set value-type to string
            cell.setAttrNS(OFFICENS, "value-type", "string")
            # Create new paragraph with value
            p = P(text=cell_delta.value)
            cell.addElement(p)

        # Apply formatting change
        has_formatting = any([
            cell_delta.font_name is not None,
            cell_delta.font_size is not None,
            cell_delta.bold is not None,
            cell_delta.italic is not None,
            cell_delta.color is not None,
        ])

        if has_formatting:
            # Get or create paragraph
            paragraphs = cell.getElementsByType(P)
            if not paragraphs:
                p = P(text="")
                cell.addElement(p)
                paragraphs = [p]
            p = paragraphs[0]

            # Get existing style name or create new one
            existing_style_name = p.getAttrNS(namespaces.TEXTNS, "style-name")
            style_name = self._get_or_create_text_style(
                doc,
                existing_style_name,
                cell_delta,
            )
            p.setAttrNS(namespaces.TEXTNS, "style-name", style_name)

    def _get_or_create_text_style(
        self,
        doc,
        existing_style_name: str | None,
        cell_delta: CellDelta,
    ) -> str:
        """
        Создаёт или обновляет именованный стиль текста с параметрами форматирования.
        Возвращает имя стиля.
        """
        # Build a unique style name based on formatting parameters
        # First, read existing style properties if any
        existing_props: dict = {}
        if existing_style_name:
            all_styles = list(doc.automaticstyles.getElementsByType(Style))
            for s in all_styles:
                if s.getAttrNS(STYLENS, "name") == existing_style_name:
                    text_props = s.getElementsByType(TextProperties)
                    if text_props:
                        tp = text_props[0]
                        fn = tp.getAttrNS(STYLENS, "font-name")
                        if fn:
                            existing_props["font_name"] = fn
                        fs = tp.getAttrNS(FONS, "font-size")
                        if fs:
                            try:
                                existing_props["font_size"] = int(float(str(fs).replace("pt", "").strip()))
                            except (ValueError, AttributeError):
                                pass
                        fw = tp.getAttrNS(FONS, "font-weight")
                        if fw:
                            existing_props["bold"] = (fw == "bold")
                        fi = tp.getAttrNS(FONS, "font-style")
                        if fi:
                            existing_props["italic"] = (fi == "italic")
                        color = tp.getAttrNS(FONS, "color")
                        if color:
                            existing_props["color"] = color
                    break

        # Merge: delta overrides existing
        font_name = cell_delta.font_name if cell_delta.font_name is not None else existing_props.get("font_name", "")
        font_size = cell_delta.font_size if cell_delta.font_size is not None else existing_props.get("font_size", 11)
        bold = cell_delta.bold if cell_delta.bold is not None else existing_props.get("bold", False)
        italic = cell_delta.italic if cell_delta.italic is not None else existing_props.get("italic", False)
        color = cell_delta.color if cell_delta.color is not None else existing_props.get("color", "")

        # Create a unique style name
        style_name = f"T_{font_name}_{font_size}_{int(bold)}_{int(italic)}_{color.replace('#', '')}"

        # Check if style already exists
        all_styles = list(doc.automaticstyles.getElementsByType(Style))
        for s in all_styles:
            if s.getAttrNS(STYLENS, "name") == style_name:
                return style_name

        # Create new style
        new_style = Style(name=style_name, family="text")
        tp_attrs = {}
        if font_name:
            tp_attrs["fontname"] = font_name
        if font_size:
            tp_attrs["fontsize"] = f"{font_size}pt"
        if bold:
            tp_attrs["fontweight"] = "bold"
        else:
            tp_attrs["fontweight"] = "normal"
        if italic:
            tp_attrs["fontstyle"] = "italic"
        else:
            tp_attrs["fontstyle"] = "normal"
        if color:
            tp_attrs["color"] = color

        tp = TextProperties(**tp_attrs)
        new_style.addElement(tp)
        doc.automaticstyles.addElement(new_style)

        return style_name

    def apply_bulk(
        self,
        files: list[FileInfo],
        delta: DeltaStore,
        log: LogStore,
        progress_callback: Callable[[int, int], None],
    ) -> BulkResult:
        """
        Последовательно обрабатывает каждый файл из files.
        Вызывает progress_callback(done, total) после каждого файла.
        """
        total = len(files)
        success_count = 0
        error_count = 0
        errors: list[str] = []

        for done, file_info in enumerate(files, start=1):
            ok = self.apply_and_save(file_info.path, delta, log)
            if ok:
                success_count += 1
            else:
                error_count += 1
                # The error message was already added to log in apply_and_save
                log_entries = log.get_all()
                if log_entries:
                    errors.append(log_entries[-1])
            progress_callback(done, total)

        return BulkResult(
            success_count=success_count,
            error_count=error_count,
            errors=errors,
        )


# ===========================================================================
# GUI-компоненты (tkinter)
# ===========================================================================


class LogPanel(tk.Frame):
    """
    Панель лога в нижней части MainWindow.
    Отображает записи LogStore в ScrolledText (только для чтения).
    """

    def __init__(self, parent: tk.Widget, log: "LogStore", **kwargs) -> None:
        super().__init__(parent, **kwargs)
        self._log = log
        self._build_ui()

    def _build_ui(self) -> None:
        label = tk.Label(self, text="Лог:", anchor="w")
        label.pack(fill=tk.X, padx=4, pady=(4, 0))

        self._text = ScrolledText(self, height=6, state=tk.DISABLED, wrap=tk.WORD)
        self._text.pack(fill=tk.BOTH, expand=True, padx=4, pady=(0, 4))

    def refresh(self) -> None:
        """Перечитать LogStore и обновить отображение."""
        self._text.config(state=tk.NORMAL)
        self._text.delete("1.0", tk.END)
        entries = self._log.get_all()
        if entries:
            self._text.insert(tk.END, "\n".join(entries))
        self._text.config(state=tk.DISABLED)
        # Scroll to end
        self._text.see(tk.END)

    def append(self, message: str) -> None:
        """Добавить одну запись в лог (без полного перечитывания)."""
        self._log.add(message)
        self._text.config(state=tk.NORMAL)
        current = self._text.get("1.0", tk.END).rstrip("\n")
        if current:
            self._text.insert(tk.END, "\n" + message)
        else:
            self._text.insert(tk.END, message)
        self._text.config(state=tk.DISABLED)
        self._text.see(tk.END)


class StatsFrame(tk.Frame):
    """
    Экран статистики: отображает путь к целевой папке, кнопку «Выбрать папку»,
    Listbox групп, список файлов выбранной группы, кнопку «Открыть образец».
    """

    def __init__(
        self,
        parent: tk.Widget,
        state: "AppState",
        on_open_sample: "Callable[[FileInfo], None]",
        on_folder_changed: "Callable[[Path], None]",
        **kwargs,
    ) -> None:
        super().__init__(parent, **kwargs)
        self._state = state
        self._on_open_sample = on_open_sample
        self._on_folder_changed = on_folder_changed
        self._build_ui()

    def _build_ui(self) -> None:
        # ── Верхняя панель: путь к папке + кнопка выбора ──────────────────
        top = tk.Frame(self)
        top.pack(fill=tk.X, padx=8, pady=(8, 4))

        tk.Label(top, text="Целевая папка:").pack(side=tk.LEFT)
        self._folder_var = tk.StringVar(value=str(self._state.target_folder))
        folder_label = tk.Label(
            top,
            textvariable=self._folder_var,
            anchor="w",
            relief=tk.SUNKEN,
            bg="white",
            padx=4,
        )
        folder_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 8))

        btn_choose = tk.Button(top, text="Выбрать папку", command=self._choose_folder)
        btn_choose.pack(side=tk.LEFT)

        # ── Основная область: два Listbox рядом ───────────────────────────
        middle = tk.Frame(self)
        middle.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        # Левый: группы
        left_frame = tk.LabelFrame(middle, text="Группы (строк — файлов)")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 4))

        self._groups_listbox = tk.Listbox(left_frame, selectmode=tk.SINGLE, exportselection=False)
        self._groups_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        groups_scroll = tk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self._groups_listbox.yview)
        groups_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self._groups_listbox.config(yscrollcommand=groups_scroll.set)
        self._groups_listbox.bind("<<ListboxSelect>>", self._on_group_select)

        # Правый: файлы группы
        right_frame = tk.LabelFrame(middle, text="Файлы группы")
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(4, 0))

        self._files_listbox = tk.Listbox(right_frame, selectmode=tk.SINGLE, exportselection=False)
        self._files_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        files_scroll = tk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self._files_listbox.yview)
        files_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self._files_listbox.config(yscrollcommand=files_scroll.set)

        # ── Нижняя панель: кнопка «Открыть образец» ───────────────────────
        bottom = tk.Frame(self)
        bottom.pack(fill=tk.X, padx=8, pady=(4, 8))

        self._btn_open = tk.Button(
            bottom,
            text="Открыть образец",
            command=self._open_sample,
            state=tk.DISABLED,
        )
        self._btn_open.pack(side=tk.RIGHT)

    # ── Внутренние обработчики ─────────────────────────────────────────────

    def _choose_folder(self) -> None:
        """Открыть диалог выбора папки и уведомить MainWindow."""
        folder = filedialog.askdirectory(
            title="Выберите целевую папку",
            initialdir=str(self._state.target_folder),
        )
        if folder:
            new_path = Path(folder)
            self._state.target_folder = new_path
            self._folder_var.set(str(new_path))
            self._on_folder_changed(new_path)

    def _on_group_select(self, event=None) -> None:
        """Обработчик выбора группы в Listbox."""
        selection = self._groups_listbox.curselection()
        if not selection:
            return
        idx = selection[0]
        sorted_keys = sorted(self._state.groups.keys())
        if idx >= len(sorted_keys):
            return
        key = sorted_keys[idx]
        self._state.selected_group_key = key
        self._populate_files(key)
        self._btn_open.config(state=tk.NORMAL)

    def _populate_files(self, key: int) -> None:
        """Заполнить правый Listbox файлами выбранной группы."""
        self._files_listbox.delete(0, tk.END)
        files = self._state.groups.get(key, [])
        for fi in files:
            self._files_listbox.insert(tk.END, fi.path.name)

    def _open_sample(self) -> None:
        """Открыть выбранный файл (или первый файл группы) в EditorFrame."""
        key = self._state.selected_group_key
        if key is None:
            return
        files = self._state.groups.get(key, [])
        if not files:
            return

        # Определяем, какой файл открыть
        file_selection = self._files_listbox.curselection()
        if file_selection:
            file_idx = file_selection[0]
            if file_idx < len(files):
                file_info = files[file_idx]
            else:
                file_info = files[0]
        else:
            file_info = files[0]

        self._on_open_sample(file_info)

    # ── Публичные методы ───────────────────────────────────────────────────

    def refresh(self) -> None:
        """Обновить отображение групп после сканирования."""
        self._folder_var.set(str(self._state.target_folder))
        self._groups_listbox.delete(0, tk.END)
        self._files_listbox.delete(0, tk.END)
        self._btn_open.config(state=tk.DISABLED)
        self._state.selected_group_key = None

        groups = self._state.groups
        if not groups:
            self._groups_listbox.insert(tk.END, "ODS-файлы не найдены")
            return

        for key in sorted(groups.keys()):
            count = len(groups[key])
            label = f"{key} строк — {count} файл(ов)"
            self._groups_listbox.insert(tk.END, label)


class FormatBar(tk.Frame):
    """
    Панель форматирования текста ячейки.
    Содержит: Combobox (гарнитура), Spinbox (размер), Checkbutton (Bold/Italic),
    кнопку выбора цвета текста.
    """

    # Список распространённых шрифтов
    _COMMON_FONTS = [
        "", "Arial", "Calibri", "Courier New", "DejaVu Sans", "DejaVu Serif",
        "FreeMono", "FreeSans", "FreeSerif", "Georgia", "Helvetica",
        "Liberation Mono", "Liberation Sans", "Liberation Serif",
        "Noto Sans", "Noto Serif", "Times New Roman", "Verdana",
    ]

    def __init__(
        self,
        parent: tk.Widget,
        on_format_change: "Callable[[CellDelta], None]",
        **kwargs,
    ) -> None:
        super().__init__(parent, **kwargs)
        self._on_format_change = on_format_change
        self._updating = False  # guard against recursive callbacks
        self._color_value: str = ""
        self._build_ui()

    def _build_ui(self) -> None:
        # ── Гарнитура ──────────────────────────────────────────────────────
        tk.Label(self, text="Шрифт:").pack(side=tk.LEFT, padx=(4, 2))
        self._font_var = tk.StringVar()
        self._font_combo = ttk.Combobox(
            self,
            textvariable=self._font_var,
            values=self._COMMON_FONTS,
            width=16,
            state="normal",
        )
        self._font_combo.pack(side=tk.LEFT, padx=(0, 6))
        self._font_combo.bind("<<ComboboxSelected>>", self._on_change)
        self._font_combo.bind("<Return>", self._on_change)
        self._font_combo.bind("<FocusOut>", self._on_change)

        # ── Размер ─────────────────────────────────────────────────────────
        tk.Label(self, text="Размер:").pack(side=tk.LEFT, padx=(0, 2))
        self._size_var = tk.StringVar(value="11")
        self._size_spin = ttk.Spinbox(
            self,
            textvariable=self._size_var,
            from_=6,
            to=72,
            width=5,
        )
        self._size_spin.pack(side=tk.LEFT, padx=(0, 6))
        self._size_spin.bind("<Return>", self._on_change)
        self._size_spin.bind("<FocusOut>", self._on_change)
        self._size_spin.bind("<<Increment>>", self._on_change)
        self._size_spin.bind("<<Decrement>>", self._on_change)

        # ── Bold ───────────────────────────────────────────────────────────
        self._bold_var = tk.BooleanVar(value=False)
        self._bold_check = ttk.Checkbutton(
            self,
            text="Bold",
            variable=self._bold_var,
            command=self._on_change,
        )
        self._bold_check.pack(side=tk.LEFT, padx=(0, 4))

        # ── Italic ─────────────────────────────────────────────────────────
        self._italic_var = tk.BooleanVar(value=False)
        self._italic_check = ttk.Checkbutton(
            self,
            text="Italic",
            variable=self._italic_var,
            command=self._on_change,
        )
        self._italic_check.pack(side=tk.LEFT, padx=(0, 6))

        # ── Цвет текста ────────────────────────────────────────────────────
        tk.Label(self, text="Цвет:").pack(side=tk.LEFT, padx=(0, 2))
        self._color_btn = tk.Button(
            self,
            text="  ",
            width=3,
            relief=tk.RAISED,
            bg="black",
            command=self._choose_color,
        )
        self._color_btn.pack(side=tk.LEFT, padx=(0, 4))

    # ── Внутренние обработчики ─────────────────────────────────────────────

    def _on_change(self, event=None) -> None:
        """Вызывается при изменении любого параметра форматирования."""
        if self._updating:
            return
        font_name = self._font_var.get().strip()
        try:
            font_size = int(self._size_var.get())
        except ValueError:
            font_size = 11
        bold = self._bold_var.get()
        italic = self._italic_var.get()
        color = self._color_value

        delta = CellDelta(
            value=None,
            font_name=font_name if font_name else None,
            font_size=font_size,
            bold=bold,
            italic=italic,
            color=color if color else None,
        )
        self._on_format_change(delta)

    def _choose_color(self) -> None:
        """Открыть системный диалог выбора цвета."""
        initial = self._color_value if self._color_value else "#000000"
        result = colorchooser.askcolor(color=initial, title="Выберите цвет текста")
        if result and result[1]:
            hex_color = result[1].upper()
            self._color_value = hex_color
            self._color_btn.config(bg=hex_color)
            self._on_change()

    # ── Публичные методы ───────────────────────────────────────────────────

    def load_cell(self, cell_data: "CellData | None") -> None:
        """Загрузить форматирование активной ячейки в панель."""
        self._updating = True
        try:
            if cell_data is None:
                self._font_var.set("")
                self._size_var.set("11")
                self._bold_var.set(False)
                self._italic_var.set(False)
                self._color_value = ""
                self._color_btn.config(bg="SystemButtonFace")
            else:
                self._font_var.set(cell_data.font_name or "")
                self._size_var.set(str(cell_data.font_size) if cell_data.font_size else "11")
                self._bold_var.set(cell_data.bold)
                self._italic_var.set(cell_data.italic)
                color = cell_data.color or ""
                self._color_value = color
                try:
                    self._color_btn.config(bg=color if color else "SystemButtonFace")
                except tk.TclError:
                    self._color_btn.config(bg="SystemButtonFace")
        finally:
            self._updating = False


class ProgressDialog:
    """
    Модальный диалог прогресса для массового применения изменений.
    Отображает ttk.Progressbar и счётчик «Обработано K из N файлов».
    """

    def __init__(self, parent: tk.Widget, total: int) -> None:
        self._total = total
        self._top = tk.Toplevel(parent)
        self._top.title("Применение изменений")
        self._top.resizable(False, False)
        self._top.grab_set()  # modal

        # Центрировать относительно родителя
        self._top.transient(parent)

        self._build_ui(total)
        self._top.update()

    def _build_ui(self, total: int) -> None:
        frame = tk.Frame(self._top, padx=20, pady=16)
        frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(frame, text="Применение изменений к файлам...", anchor="w").pack(
            fill=tk.X, pady=(0, 8)
        )

        self._progress = ttk.Progressbar(
            frame,
            orient=tk.HORIZONTAL,
            length=320,
            mode="determinate",
            maximum=total if total > 0 else 1,
        )
        self._progress.pack(fill=tk.X, pady=(0, 8))

        self._label_var = tk.StringVar(value=f"Обработано 0 из {total} файлов")
        self._label = tk.Label(frame, textvariable=self._label_var, anchor="w")
        self._label.pack(fill=tk.X)

    def update(self, done: int, total: int) -> None:
        """Обновить прогресс-бар и счётчик."""
        self._progress["value"] = done
        self._label_var.set(f"Обработано {done} из {total} файлов")
        try:
            self._top.update()
        except tk.TclError:
            pass

    def close(self) -> None:
        """Закрыть диалог."""
        try:
            self._top.grab_release()
            self._top.destroy()
        except tk.TclError:
            pass


class EditorFrame(tk.Frame):
    """
    Экран редактора: таблица (ttk.Treeview) для отображения данных первого листа,
    панель форматирования (FormatBar), кнопки действий.

    Двойной клик по ячейке переводит её в режим редактирования.
    Enter или потеря фокуса сохраняет значение в DeltaStore.
    Изменённые ячейки визуально выделяются.
    """

    _MODIFIED_TAG = "modified"
    _MODIFIED_BG = "#FFF3CD"   # светло-жёлтый фон для изменённых ячеек

    def __init__(
        self,
        parent: tk.Widget,
        state: "AppState",
        on_close: "Callable[[], None]",
        **kwargs,
    ) -> None:
        super().__init__(parent, **kwargs)
        self._state = state
        self._on_close = on_close
        self._active_cell: tuple[int, int] | None = None  # (row_idx, col_idx)
        self._entry_widget: tk.Entry | None = None
        self._col_ids: list[str] = []   # Treeview column identifiers
        self._row_iids: list[str] = []  # Treeview item iids (one per data row)
        self._build_ui()
        self._load_data()

    # ── Построение UI ──────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        # ── Верхняя панель: имя файла ──────────────────────────────────────
        top = tk.Frame(self)
        top.pack(fill=tk.X, padx=8, pady=(8, 2))

        file_name = (
            self._state.current_file.path.name
            if self._state.current_file
            else "—"
        )
        tk.Label(top, text=f"Редактирование: {file_name}", font=("", 10, "bold")).pack(side=tk.LEFT)

        # ── FormatBar ──────────────────────────────────────────────────────
        self._format_bar = FormatBar(self, on_format_change=self._apply_format)
        self._format_bar.pack(fill=tk.X, padx=8, pady=(2, 4))

        # ── Таблица (Treeview + скроллбары) ───────────────────────────────
        table_frame = tk.Frame(self)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 4))

        self._tree = ttk.Treeview(table_frame, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self._tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Тег для изменённых ячеек
        self._tree.tag_configure(self._MODIFIED_TAG, background=self._MODIFIED_BG)

        # Привязки событий
        self._tree.bind("<Double-1>", self._on_double_click)
        self._tree.bind("<<TreeviewSelect>>", self._on_select)

        # ── Нижняя панель: кнопки действий ────────────────────────────────
        bottom = tk.Frame(self)
        bottom.pack(fill=tk.X, padx=8, pady=(0, 8))

        # Кнопки действий
        self._btn_save = tk.Button(
            bottom,
            text="Сохранить этот файл",
            state=tk.DISABLED,
            command=self._on_save,
        )
        self._btn_save.pack(side=tk.LEFT, padx=(0, 6))

        self._btn_apply_all = tk.Button(
            bottom,
            text="Применить ко всем однотипным ODS",
            state=tk.DISABLED,
            command=self._on_apply_all,
        )
        self._btn_apply_all.pack(side=tk.LEFT, padx=(0, 6))

        self._btn_close = tk.Button(
            bottom,
            text="Закрыть файл / Вернуться к статистике",
            command=self._on_close_editor,
        )
        self._btn_close.pack(side=tk.RIGHT)

    # ── Загрузка данных ────────────────────────────────────────────────────

    def _load_data(self) -> None:
        """Загрузить SheetData из state и заполнить Treeview."""
        sheet = self._state.sheet_data
        if sheet is None:
            return

        max_col = sheet.max_col
        max_row = sheet.max_row

        # Формируем идентификаторы столбцов: A, B, C, ... Z, AA, AB, ...
        self._col_ids = [self._col_letter(c) for c in range(max_col + 1)]

        self._tree["columns"] = self._col_ids
        for col_id in self._col_ids:
            self._tree.heading(col_id, text=col_id)
            self._tree.column(col_id, width=80, minwidth=40, stretch=True)

        # Заполняем строки
        self._row_iids = []
        for r in range(max_row + 1):
            values = []
            for c in range(max_col + 1):
                cell = sheet.cells.get((r, c))
                values.append(cell.value if cell else "")
            iid = self._tree.insert("", tk.END, values=values)
            self._row_iids.append(iid)

        # Применяем теги для уже изменённых ячеек (если дельта не пуста)
        self._refresh_modified_tags()

    @staticmethod
    def _col_letter(col_idx: int) -> str:
        """Преобразует индекс столбца (0-based) в буквенное обозначение (A, B, ..., Z, AA, ...)."""
        result = ""
        n = col_idx
        while True:
            result = chr(ord("A") + n % 26) + result
            n = n // 26 - 1
            if n < 0:
                break
        return result

    # ── Обновление тегов изменённых ячеек ─────────────────────────────────

    def _refresh_modified_tags(self) -> None:
        """Перекрасить строки, содержащие изменённые ячейки."""
        # Собираем строки, в которых есть изменения
        modified_rows: set[int] = set()
        for (row, _col), _delta in self._state.delta.items():
            modified_rows.add(row)

        for r, iid in enumerate(self._row_iids):
            if r in modified_rows:
                self._tree.item(iid, tags=(self._MODIFIED_TAG,))
            else:
                self._tree.item(iid, tags=())

    # ── Редактирование ячейки ──────────────────────────────────────────────

    def _on_double_click(self, event: tk.Event) -> None:
        """Двойной клик — перевести ячейку в режим редактирования."""
        region = self._tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col_id = self._tree.identify_column(event.x)   # "#1", "#2", ...
        iid = self._tree.identify_row(event.y)
        if not iid or not col_id:
            return

        col_idx = int(col_id.lstrip("#")) - 1  # 0-based
        row_idx = self._row_iids.index(iid) if iid in self._row_iids else -1
        if row_idx < 0:
            return

        self._start_edit(iid, col_idx, row_idx)

    def _start_edit(self, iid: str, col_idx: int, row_idx: int) -> None:
        """Разместить Entry-виджет поверх ячейки для редактирования."""
        # Завершить предыдущее редактирование, если есть
        self._finish_edit(save=True)

        self._active_cell = (row_idx, col_idx)

        # Получить текущее значение ячейки
        values = self._tree.item(iid, "values")
        current_value = values[col_idx] if col_idx < len(values) else ""

        # Вычислить координаты ячейки в Treeview
        bbox = self._tree.bbox(iid, column=self._col_ids[col_idx])
        if not bbox:
            return
        x, y, width, height = bbox

        # Создать Entry поверх ячейки
        entry = tk.Entry(self._tree)
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, current_value)
        entry.select_range(0, tk.END)
        entry.focus_set()

        self._entry_widget = entry

        entry.bind("<Return>", lambda e: self._finish_edit(save=True))
        entry.bind("<Escape>", lambda e: self._finish_edit(save=False))
        entry.bind("<FocusOut>", lambda e: self._finish_edit(save=True))
        entry.bind("<Tab>", lambda e: self._finish_edit(save=True))

    def _finish_edit(self, save: bool = True) -> None:
        """Завершить редактирование ячейки."""
        if self._entry_widget is None:
            return

        entry = self._entry_widget
        self._entry_widget = None

        if save and self._active_cell is not None:
            new_value = entry.get()
            row_idx, col_idx = self._active_cell
            self._save_cell_value(row_idx, col_idx, new_value)

        try:
            entry.destroy()
        except tk.TclError:
            pass

    def _save_cell_value(self, row_idx: int, col_idx: int, new_value: str) -> None:
        """Сохранить новое значение ячейки в DeltaStore и обновить Treeview."""
        # Получить существующую дельту или создать новую
        existing = self._state.delta.get(row_idx, col_idx)
        if existing is not None:
            delta = CellDelta(
                value=new_value,
                font_name=existing.font_name,
                font_size=existing.font_size,
                bold=existing.bold,
                italic=existing.italic,
                color=existing.color,
            )
        else:
            delta = CellDelta(
                value=new_value,
                font_name=None,
                font_size=None,
                bold=None,
                italic=None,
                color=None,
            )
        self._state.delta.set(row_idx, col_idx, delta)

        # Обновить значение в Treeview
        if row_idx < len(self._row_iids):
            iid = self._row_iids[row_idx]
            values = list(self._tree.item(iid, "values"))
            if col_idx < len(values):
                values[col_idx] = new_value
                self._tree.item(iid, values=values)

        # Обновить теги
        self._refresh_modified_tags()

        # Включить кнопки сохранения
        self._btn_save.config(state=tk.NORMAL)
        self._btn_apply_all.config(state=tk.NORMAL)

    # ── Выбор ячейки ──────────────────────────────────────────────────────

    def _on_select(self, event=None) -> None:
        """Обработчик выбора строки — обновить FormatBar."""
        selection = self._tree.selection()
        if not selection:
            self._format_bar.load_cell(None)
            return
        iid = selection[0]
        row_idx = self._row_iids.index(iid) if iid in self._row_iids else -1
        if row_idx < 0:
            self._format_bar.load_cell(None)
            return

        # Определяем активный столбец (если есть активная ячейка в этой строке)
        col_idx = 0
        if self._active_cell and self._active_cell[0] == row_idx:
            col_idx = self._active_cell[1]

        self._active_cell = (row_idx, col_idx)
        self._update_format_bar(row_idx, col_idx)

    def _update_format_bar(self, row_idx: int, col_idx: int) -> None:
        """Обновить FormatBar форматированием активной ячейки."""
        # Сначала проверяем дельту
        delta = self._state.delta.get(row_idx, col_idx)
        sheet = self._state.sheet_data

        if delta is not None:
            # Строим CellData из дельты + исходных данных
            orig = sheet.cells.get((row_idx, col_idx)) if sheet else None
            cell_data = CellData(
                value=delta.value if delta.value is not None else (orig.value if orig else ""),
                font_name=delta.font_name if delta.font_name is not None else (orig.font_name if orig else ""),
                font_size=delta.font_size if delta.font_size is not None else (orig.font_size if orig else 11),
                bold=delta.bold if delta.bold is not None else (orig.bold if orig else False),
                italic=delta.italic if delta.italic is not None else (orig.italic if orig else False),
                color=delta.color if delta.color is not None else (orig.color if orig else ""),
            )
        elif sheet:
            cell_data = sheet.cells.get((row_idx, col_idx))
        else:
            cell_data = None

        self._format_bar.load_cell(cell_data)

    # ── Применение форматирования ──────────────────────────────────────────

    def _apply_format(self, format_delta: CellDelta) -> None:
        """Применить форматирование к выделенным ячейкам."""
        selection = self._tree.selection()
        if not selection:
            return

        for iid in selection:
            row_idx = self._row_iids.index(iid) if iid in self._row_iids else -1
            if row_idx < 0:
                continue

            # Применяем ко всем столбцам строки или только к активному
            col_idx = self._active_cell[1] if self._active_cell and self._active_cell[0] == row_idx else 0

            existing = self._state.delta.get(row_idx, col_idx)
            if existing is not None:
                merged = CellDelta(
                    value=existing.value,
                    font_name=format_delta.font_name if format_delta.font_name is not None else existing.font_name,
                    font_size=format_delta.font_size if format_delta.font_size is not None else existing.font_size,
                    bold=format_delta.bold if format_delta.bold is not None else existing.bold,
                    italic=format_delta.italic if format_delta.italic is not None else existing.italic,
                    color=format_delta.color if format_delta.color is not None else existing.color,
                )
            else:
                merged = CellDelta(
                    value=None,
                    font_name=format_delta.font_name,
                    font_size=format_delta.font_size,
                    bold=format_delta.bold,
                    italic=format_delta.italic,
                    color=format_delta.color,
                )
            self._state.delta.set(row_idx, col_idx, merged)

        self._refresh_modified_tags()
        self._btn_save.config(state=tk.NORMAL)
        self._btn_apply_all.config(state=tk.NORMAL)

    # ── Кнопки действий ───────────────────────────────────────────────────

    def _on_save(self) -> None:
        """Сохранить изменения в текущем файле."""
        # Завершить активное редактирование перед сохранением
        self._finish_edit(save=True)

        if self._state.current_file is None:
            return

        processor = Processor()
        ok = processor.apply_and_save(
            self._state.current_file.path,
            self._state.delta,
            self._state.log,
        )

        if ok:
            messagebox.showinfo("Сохранено", "Файл успешно сохранён.")
        else:
            messagebox.showerror("Ошибка", "Не удалось сохранить файл. Подробности в логе.")

        # Обновить панель лога
        self._refresh_log_panel()

    def _on_apply_all(self) -> None:
        """Применить изменения ко всем файлам группы."""
        # Завершить активное редактирование перед применением
        self._finish_edit(save=True)

        key = self._state.selected_group_key
        if key is None:
            return

        files = self._state.groups.get(key, [])
        n = len(files)

        confirmed = messagebox.askyesno(
            "Подтверждение",
            f"Будет обновлено {n} файлов (все с {key} строками). Продолжить?",
        )
        if not confirmed:
            return

        # Показать ProgressDialog
        progress_dialog = ProgressDialog(self, total=n)

        def progress_callback(done: int, total: int) -> None:
            progress_dialog.update(done, total)
            # Обновить UI во время обработки
            try:
                self.update()
            except tk.TclError:
                pass

        processor = Processor()
        result = processor.apply_bulk(
            files,
            self._state.delta,
            self._state.log,
            progress_callback,
        )

        progress_dialog.close()

        # Показать итоговый отчёт
        if result.error_count == 0:
            summary = f"Успешно обновлено файлов: {result.success_count}."
        else:
            summary = (
                f"Успешно обновлено: {result.success_count} файл(ов).\n"
                f"С ошибками: {result.error_count} файл(ов).\n"
                "Подробности в логе."
            )
        messagebox.showinfo("Результат", summary)

        # Обновить панель лога
        self._refresh_log_panel()

    def _on_close_editor(self) -> None:
        """Закрыть редактор и вернуться к статистике."""
        # Завершить активное редактирование
        self._finish_edit(save=True)

        # Проверить наличие несохранённых изменений
        if not self._state.delta.is_empty():
            confirmed = messagebox.askyesno(
                "Несохранённые изменения",
                "Есть несохранённые изменения. Закрыть без сохранения?",
            )
            if not confirmed:
                return

        self._on_close()

    def _refresh_log_panel(self) -> None:
        """Обновить LogPanel, если она доступна через родительское окно."""
        # Поднимаемся по иерархии виджетов до MainWindow
        widget = self.master
        while widget is not None:
            if hasattr(widget, "_log_panel"):
                widget._log_panel.refresh()
                return
            widget = getattr(widget, "master", None)


class MainWindow(tk.Tk):
    """
    Корневое окно приложения.
    Содержит StatsFrame или EditorFrame (переключение через pack/forget).
    Внизу всегда отображается LogPanel.
    """

    def __init__(self) -> None:
        super().__init__()
        self.title("ODS Bulk Editor")
        self.geometry("900x600")
        self.minsize(600, 400)

        # Инициализация состояния
        target_folder = Path(__file__).resolve().parent
        self._state = AppState(
            target_folder=target_folder,
            scan_results=[],
            groups={},
            selected_group_key=None,
            current_file=None,
            sheet_data=None,
            delta=DeltaStore(),
            log=LogStore(),
        )

        self._scanner = Scanner()
        self._group_manager = GroupManager()

        self._build_ui()

        # Автоматическое сканирование при запуске
        self._run_scan()

    def _build_ui(self) -> None:
        # Основная область (StatsFrame / EditorFrame)
        self._content_frame = tk.Frame(self)
        self._content_frame.pack(fill=tk.BOTH, expand=True)

        # StatsFrame
        self._stats_frame = StatsFrame(
            self._content_frame,
            state=self._state,
            on_open_sample=self._show_editor,
            on_folder_changed=self._on_folder_changed,
        )
        self._stats_frame.pack(fill=tk.BOTH, expand=True)

        # Разделитель
        ttk.Separator(self, orient=tk.HORIZONTAL).pack(fill=tk.X)

        # LogPanel внизу
        self._log_panel = LogPanel(self, self._state.log)
        self._log_panel.pack(fill=tk.BOTH, side=tk.BOTTOM)

    # ── Сканирование ───────────────────────────────────────────────────────

    def _run_scan(self) -> None:
        """Запустить сканирование целевой папки и обновить StatsFrame."""
        folder = self._state.target_folder
        if not folder.exists() or not folder.is_dir():
            self._state.log.add(f"Папка не существует: {folder}")
            self._log_panel.refresh()
            self._state.scan_results = []
            self._state.groups = {}
            self._stats_frame.refresh()
            return

        self._state.scan_results = self._scanner.scan(folder, self._state.log)
        self._group_manager.build(self._state.scan_results)
        self._state.groups = self._group_manager.get_groups()
        self._stats_frame.refresh()
        self._log_panel.refresh()

    def _on_folder_changed(self, new_folder: Path) -> None:
        """Вызывается при смене целевой папки — запускает повторное сканирование."""
        self._run_scan()

    # ── Переключение экранов ───────────────────────────────────────────────

    def _show_editor(self, file_info: "FileInfo") -> None:
        """
        Переключиться на EditorFrame для указанного файла.
        Загружает файл через OdsReader, сохраняет SheetData в state,
        скрывает StatsFrame и показывает EditorFrame.
        """
        self._state.current_file = file_info

        # Загрузить данные файла
        reader = OdsReader()
        try:
            sheet_data = reader.load(file_info.path)
        except Exception as e:
            self._state.log.add(f"Ошибка открытия файла {file_info.path.name}: {e}")
            self._log_panel.refresh()
            messagebox.showerror("Ошибка", f"Не удалось открыть файл:\n{e}")
            return

        self._state.sheet_data = sheet_data
        # Сбросить дельту при открытии нового файла
        self._state.delta.clear()

        # Скрыть StatsFrame
        self._stats_frame.pack_forget()

        # Создать и показать EditorFrame
        self._editor_frame = EditorFrame(
            self._content_frame,
            state=self._state,
            on_close=self.show_stats,
        )
        self._editor_frame.pack(fill=tk.BOTH, expand=True)

    def show_stats(self) -> None:
        """
        Вернуться на StatsFrame (вызывается из EditorFrame).
        Скрывает EditorFrame и показывает StatsFrame без повторного сканирования.
        """
        # Скрыть EditorFrame, если он существует
        if hasattr(self, "_editor_frame") and self._editor_frame is not None:
            self._editor_frame.pack_forget()
            self._editor_frame.destroy()
            self._editor_frame = None

        # Сбросить состояние редактора
        self._state.current_file = None
        self._state.sheet_data = None

        # Показать StatsFrame
        self._stats_frame.pack(fill=tk.BOTH, expand=True)


# ===========================================================================
# Точка входа
# ===========================================================================

if __name__ == "__main__":
    MainWindow().mainloop()
