"""
tablered
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
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from tkinter.scrolledtext import ScrolledText
from pathlib import Path
from dataclasses import dataclass, field
from typing import Callable, Iterable
import datetime
import time
import tkinter.font as tkfont


# ---------------------------------------------------------------------------
# Проверка и установка зависимостей
# ---------------------------------------------------------------------------

_MISSING_DEPS: list[str] = []

try:
    from odf.opendocument import load as odf_load, OpenDocumentSpreadsheet
    from odf.table import Table, TableRow, TableCell
    from odf.text import P
    from odf import style, namespaces
    from odf.style import Style, TextProperties, TableCellProperties, FontFace
    from odf.namespaces import OFFICENS, STYLENS, FONS, SVGNS
except ImportError:
    _MISSING_DEPS.append("python3-odf  (pip: odfpy)")


def _check_deps_and_install() -> bool:
    if not _MISSING_DEPS:
        return True
    deps_list = "\n".join(f"  • {d}" for d in _MISSING_DEPS)
    msg = (
        "Не найдены обязательные зависимости:\n"
        f"{deps_list}\n\n"
        "Установить автоматически?\n"
        "• «apt» — системная установка (требует sudo)\n"
        "• «pip» — установка в текущее окружение\n"
        "• «Нет» — выйти для ручной установки"
    )
    root = tk.Tk()
    root.withdraw()
    result = messagebox.askyesnocancel("Зависимости не найдены", msg)
    root.destroy()

    if result is None or result is False and not messagebox.askyesno(
        "Выход", "Установите зависимости вручную и перезапустите.\nВыйти?"
    ):
        sys.exit(1)

    if result is True:
        try:
            subprocess.run(
                ["sudo", "apt", "install", "-y", "python3-odf"],
                check=True,
            )
        except (subprocess.CalledProcessError, FileNotFoundError):
            try:
                subprocess.run(
                    [sys.executable, "-m", "pip", "install", "odfpy"],
                    check=True,
                )
            except (subprocess.CalledProcessError, FileNotFoundError):
                messagebox.showerror(
                    "Ошибка",
                    "Не удалось установить зависимости.\n"
                    "Установите вручную: sudo apt install python3-odf",
                )
                sys.exit(1)
        return True

    sys.exit(1)


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
    value: str
    font_name: str
    font_size: int
    bold: bool
    italic: bool
    underline: bool
    color: str


@dataclass
class SheetData:
    """Данные первого листа ODS-файла."""
    cells: dict[tuple[int, int], CellData]   # (row, col) → CellData
    max_row: int
    max_col: int


@dataclass
class CellDelta:
    """Изменения одной ячейки. None означает «значение не менялось»."""
    value: str | None
    font_name: str | None
    font_size: int | None
    bold: bool | None
    italic: bool | None
    underline: bool | None
    color: str | None


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
    """Хранит записи лога в памяти и дублирует в файл."""

    def __init__(self) -> None:
        self._entries: list[str] = []
        self._log_dir = Path.home() / ".ods_bulk_editor"
        try:
            self._log_dir.mkdir(parents=True, exist_ok=True)
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            self._log_path = self._log_dir / f"session_{ts}.log"
            self._log_file = open(self._log_path, "a", encoding="utf-8")
        except OSError:
            self._log_path = None
            self._log_file = None

    def add(self, message: str) -> None:
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        entry = f"[{ts}] {message}"
        self._entries.append(entry)
        if self._log_file:
            try:
                self._log_file.write(entry + "\n")
                self._log_file.flush()
            except OSError:
                pass

    def get_all(self) -> list[str]:
        return list(self._entries)

    def clear(self) -> None:
        self._entries.clear()

    def close(self) -> None:
        if self._log_file:
            try:
                self._log_file.close()
            except OSError:
                pass


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
                    underline=fmt["underline"],
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
            "underline": False,
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
                    uus = tp.getAttrNS(STYLENS, "text-underline-style")
                    if uus and uus != "none":
                        result["underline"] = True
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

                self._apply_delta_to_cell(cell, cell_delta, doc, log)

            doc.save(str(path))
            return True
        except Exception as e:
            log.add(f"Ошибка обработки файла {path.name}: {e}")
            return False

    def _apply_delta_to_cell(self, cell, cell_delta: CellDelta, doc, log: LogStore) -> None:
        """Применяет изменения дельты к ячейке."""
        if cell_delta.value is not None:
            for attr in ["formula", "valuetype", "value", "datevalue"]:
                try:
                    cell.removeAttribute(attr)
                except Exception:
                    pass
            for p in list(cell.getElementsByType(P)):
                cell.removeChild(p)

            cell.setAttrNS(OFFICENS, "value-type", "string")
            p = P(text=cell_delta.value)
            cell.addElement(p)

        has_formatting = any([
            cell_delta.font_name is not None,
            cell_delta.font_size is not None,
            cell_delta.bold is not None,
            cell_delta.italic is not None,
            cell_delta.underline is not None,
            cell_delta.color is not None,
        ])

        if has_formatting:
            existing_style_name = cell.getAttrNS(namespaces.TABLENS, "style-name")
            new_style_name = self._get_or_create_table_cell_style(
                doc, existing_style_name, cell_delta
            )

            cell.setAttrNS(namespaces.TABLENS, "style-name", new_style_name)

            for p in cell.getElementsByType(P):
                for attr in list(p.attributes.keys()):
                    if attr[1] == "style-name":
                        del p.attributes[attr]

            log.add(f"  Стиль для ячейки: {new_style_name}")

    def _get_or_create_table_cell_style(
        self, doc, existing_style_name: str | None, cell_delta: CellDelta
    ) -> str:
        fn = str(cell_delta.font_name).replace(' ', '')
        fs = str(cell_delta.font_size)
        b = str(cell_delta.bold)
        i = str(cell_delta.italic)
        u = str(cell_delta.underline)
        c = str(cell_delta.color).replace('#', '')

        style_name = f"Ce_{existing_style_name or 'Def'}_{fn}_{fs}_{b}_{i}_{u}_{c}"

        all_styles = list(doc.automaticstyles.getElementsByType(Style))
        for s in all_styles:
            if s.getAttrNS(STYLENS, "name") == style_name:
                return style_name

        if cell_delta.font_name:
            self._ensure_font_face(doc, cell_delta.font_name)

        new_style = Style(name=style_name, family="table-cell")

        if existing_style_name:
            new_style.setAttrNS(STYLENS, "parent-style-name", existing_style_name)

        tp = TextProperties()

        if cell_delta.font_size is not None:
            tp.setAttrNS(FONS, "font-size", f"{cell_delta.font_size}pt")

        if cell_delta.bold is not None:
            tp.setAttrNS(FONS, "font-weight", "bold" if cell_delta.bold else "normal")

        if cell_delta.italic is not None:
            tp.setAttrNS(FONS, "font-style", "italic" if cell_delta.italic else "normal")

        if cell_delta.font_name:
            tp.setAttrNS(STYLENS, "font-name", cell_delta.font_name)
            tp.setAttrNS(FONS, "font-family", f"'{cell_delta.font_name}'")

        if cell_delta.underline is not None:
            if cell_delta.underline:
                tp.setAttrNS(STYLENS, "text-underline-style", "solid")
                tp.setAttrNS(STYLENS, "text-underline-type", "single")
            else:
                tp.setAttrNS(STYLENS, "text-underline-style", "none")

        if cell_delta.color:
            tp.setAttrNS(FONS, "color", cell_delta.color)

        new_style.addElement(tp)
        doc.automaticstyles.addElement(new_style)

        return style_name

    @staticmethod
    def _ensure_font_face(doc, font_name: str) -> None:
        try:
            font_faces = doc.fontfacedecls.getElementsByType(FontFace)
            for ff in font_faces:
                if ff.getAttrNS(STYLENS, "name") == font_name:
                    return
        except AttributeError:
            return

        ff = FontFace(name=font_name)
        ff.setAttrNS(STYLENS, "name", font_name)
        ff.setAttrNS(SVGNS, "font-family", f"'{font_name}'")
        ff.setAttrNS(STYLENS, "font-pitch", "variable")
        try:
            doc.fontfacedecls.addElement(ff)
        except AttributeError:
            pass

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


class BottomPanel(tk.Frame):
    """
    Нижняя панель с вкладками:
      — «Текст ячейки»: полное содержимое выбранной ячейки (редактируемое).
      — «Лог»: записи LogStore.
    """

    def __init__(self, parent: tk.Widget, log: "LogStore", **kwargs) -> None:
        super().__init__(parent, **kwargs)
        self._log = log
        self._on_cell_text_save: Callable[[], None] | None = None
        self._building = False
        self._build_ui()

    def _build_ui(self) -> None:
        self._notebook = ttk.Notebook(self)
        self._notebook.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        # ── Вкладка «Текст ячейки» ────────────────────────────────────────
        cell_frame = tk.Frame(self._notebook)
        self._notebook.add(cell_frame, text=" Текст ячейки ")

        btn_frame = tk.Frame(cell_frame)
        btn_frame.pack(fill=tk.X)
        self._btn_apply_text = tk.Button(
            btn_frame, text="Применить", command=self._save_cell_text,
        )
        self._btn_apply_text.pack(side=tk.RIGHT, padx=2, pady=2)

        self._cell_text = ScrolledText(
            cell_frame, height=6, wrap=tk.WORD, font=("Consolas", 10),
        )
        self._cell_text.pack(fill=tk.BOTH, expand=True)
        self._cell_text.bind("<Control-Return>", lambda e: self._save_cell_text())

        # ── Вкладка «Лог» ──────────────────────────────────────────────────
        log_frame = tk.Frame(self._notebook)
        self._notebook.add(log_frame, text=" Лог ")
        self._log_text = ScrolledText(
            log_frame, height=6, state=tk.DISABLED, wrap=tk.WORD,
        )
        self._log_text.pack(fill=tk.BOTH, expand=True)

    def set_cell_text_callback(self, callback: Callable[[], None]) -> None:
        self._on_cell_text_save = callback

    def _save_cell_text(self) -> None:
        if self._on_cell_text_save:
            self._on_cell_text_save()

    def get_cell_text(self) -> str:
        return self._cell_text.get("1.0", tk.END).rstrip("\n")

    def refresh_log(self) -> None:
        self._log_text.config(state=tk.NORMAL)
        self._log_text.delete("1.0", tk.END)
        entries = self._log.get_all()
        if entries:
            self._log_text.insert(tk.END, "\n".join(entries))
        self._log_text.config(state=tk.DISABLED)
        self._log_text.see(tk.END)

    def append_log(self, message: str) -> None:
        self._log.add(message)
        self._log_text.config(state=tk.NORMAL)
        current = self._log_text.get("1.0", tk.END).rstrip("\n")
        if current:
            self._log_text.insert(tk.END, "\n" + message)
        else:
            self._log_text.insert(tk.END, message)
        self._log_text.config(state=tk.DISABLED)
        self._log_text.see(tk.END)

    def show_cell_content(self, text: str) -> None:
        self._building = True
        self._cell_text.delete("1.0", tk.END)
        self._cell_text.insert(tk.END, text)
        self._building = False


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
        self._italic_check.pack(side=tk.LEFT, padx=(0, 4))

        # ── Underline ───────────────────────────────────────────────────────
        self._underline_var = tk.BooleanVar(value=False)
        self._underline_check = ttk.Checkbutton(
            self,
            text="Underline",
            variable=self._underline_var,
            command=self._on_change,
        )
        self._underline_check.pack(side=tk.LEFT, padx=(0, 6))

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
        underline = self._underline_var.get()
        color = self._color_value

        delta = CellDelta(
            value=None,
            font_name=font_name if font_name else None,
            font_size=font_size,
            bold=bold,
            italic=italic,
            underline=underline,
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
                self._underline_var.set(False)
                self._color_value = ""
                self._color_btn.config(bg="#d9d9d9")
            else:
                self._font_var.set(cell_data.font_name or "")
                self._size_var.set(str(cell_data.font_size) if cell_data.font_size else "11")
                self._bold_var.set(cell_data.bold)
                self._italic_var.set(cell_data.italic)
                self._underline_var.set(cell_data.underline)
                color = cell_data.color or ""
                self._color_value = color
                try:
                    self._color_btn.config(bg=color if color else "#d9d9d9")
                except tk.TclError:
                    self._color_btn.config(bg="#d9d9d9")
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
    Экран редактора: Canvas-таблица с сеткой (границы ячеек) для отображения
    данных первого листа, панель форматирования (FormatBar), кнопки действий.

    Двойной клик по ячейке переводит её в режим редактирования.
    Enter или потеря фокуса сохраняет значение в DeltaStore.
    Изменённые ячейки визуально выделяются.
    """

    _MODIFIED_BG = "#FFF3CD"
    _ROW_SEL_BG = "#D6E4F0"
    _CELL_SEL_BG = "#BDD7EE"
    _NORMAL_BG = "white"
    _HEADER_BG = "#E8E8E8"
    _GRID_COLOR = "#B0B0B0"
    _COL_WIDTH = 100
    _ROW_HEIGHT = 24
    _HEADER_HEIGHT = 24
    _ROW_HEADER_WIDTH = 40

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
        self._active_cell: tuple[int, int] | None = None
        self._entry_widget: tk.Entry | None = None
        self._col_ids: list[str] = []
        self._num_rows: int = 0
        self._num_cols: int = 0
        self._cell_rects: dict[tuple[int, int], int] = {}
        self._cell_texts: dict[tuple[int, int], int] = {}
        self._default_font: tkfont.Font | None = None
        self._edit_guard_until: float = 0.0
        self._build_ui()
        self._load_data()

    def _build_ui(self) -> None:
        top = tk.Frame(self)
        top.pack(fill=tk.X, padx=8, pady=(8, 2))

        file_name = (
            self._state.current_file.path.name
            if self._state.current_file
            else "—"
        )
        tk.Label(top, text=f"Редактирование: {file_name}", font=("", 10, "bold")).pack(side=tk.LEFT)

        self._format_bar = FormatBar(self, on_format_change=self._apply_format)
        self._format_bar.pack(fill=tk.X, padx=8, pady=(2, 4))

        # ── Canvas-таблица с сеткой ───────────────────────────────────────
        grid_container = tk.Frame(self)
        grid_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 4))

        # Угловая ячейка (пересечение заголовков строк и столбцов)
        self._corner = tk.Canvas(
            grid_container, width=self._ROW_HEADER_WIDTH,
            height=self._HEADER_HEIGHT, bg=self._HEADER_BG,
            highlightthickness=1, highlightbackground=self._GRID_COLOR,
        )
        self._corner.grid(row=0, column=0, sticky="nsew")

        # Заголовки столбцов
        self._col_header = tk.Canvas(
            grid_container, height=self._HEADER_HEIGHT,
            bg=self._HEADER_BG, highlightthickness=0,
        )
        self._col_header.grid(row=0, column=1, sticky="ew")

        # Заголовки строк (номера)
        self._row_header = tk.Canvas(
            grid_container, width=self._ROW_HEADER_WIDTH,
            bg=self._HEADER_BG, highlightthickness=0,
        )
        self._row_header.grid(row=1, column=0, sticky="ns")

        # Основной Canvas с ячейками
        self._canvas = tk.Canvas(grid_container, bg="white", highlightthickness=0)
        self._canvas.grid(row=1, column=1, sticky="nsew")

        # Скроллбары
        vsb = ttk.Scrollbar(grid_container, orient=tk.VERTICAL, command=self._on_yscroll)
        vsb.grid(row=0, column=2, rowspan=2, sticky="ns")
        hsb = ttk.Scrollbar(grid_container, orient=tk.HORIZONTAL, command=self._on_xscroll)
        hsb.grid(row=2, column=0, columnspan=2, sticky="ew")

        self._canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        grid_container.grid_rowconfigure(1, weight=1)
        grid_container.grid_columnconfigure(1, weight=1)

        # Обработчики событий
        self._canvas.bind("<Button-1>", self._on_click_release)
        self._canvas.bind("<Double-1>", self._on_double_click)
        self._canvas.bind("<Button-4>", lambda e: self._on_mousewheel(-1))
        self._canvas.bind("<Button-5>", lambda e: self._on_mousewheel(1))
        self._canvas.bind("<MouseWheel>", lambda e: self._on_mousewheel(-1 if e.delta > 0 else 1))

        bp = self._find_bottom_panel()
        if bp:
            bp.set_cell_text_callback(self._save_cell_text_from_bar)

        # ── Кнопки действий ───────────────────────────────────────────────
        bottom = tk.Frame(self)
        bottom.pack(fill=tk.X, padx=8, pady=(0, 8))

        self._btn_save = tk.Button(
            bottom, text="Сохранить этот файл",
            state=tk.DISABLED, command=self._on_save,
        )
        self._btn_save.pack(side=tk.LEFT, padx=(0, 6))

        self._btn_apply_all = tk.Button(
            bottom, text="Применить ко всем однотипным ODS",
            state=tk.DISABLED, command=self._on_apply_all,
        )
        self._btn_apply_all.pack(side=tk.LEFT, padx=(0, 6))

        self._btn_close = tk.Button(
            bottom, text="Закрыть файл / Вернуться к статистике",
            command=self._on_close_editor,
        )
        self._btn_close.pack(side=tk.RIGHT)

    # ── Скроллинг ─────────────────────────────────────────────────────────

    def _on_yscroll(self, *args) -> None:
        self._canvas.yview(*args)
        self._row_header.yview(*args)

    def _on_xscroll(self, *args) -> None:
        self._canvas.xview(*args)
        self._col_header.xview(*args)

    def _on_mousewheel(self, direction: int) -> None:
        self._canvas.yview_scroll(direction * 3, "units")
        self._row_header.yview_scroll(direction * 3, "units")

    # ── Загрузка данных ───────────────────────────────────────────────────

    def _load_data(self) -> None:
        sheet = self._state.sheet_data
        if sheet is None:
            return

        self._default_font = tkfont.Font(family="TkDefaultFont", size=10)

        max_col = sheet.max_col
        max_row = sheet.max_row
        self._num_cols = max_col + 1
        self._num_rows = max_row + 1
        self._col_ids = [self._col_letter(c) for c in range(self._num_cols)]

        total_width = self._num_cols * self._COL_WIDTH
        total_height = self._num_rows * self._ROW_HEIGHT

        # Заголовки столбцов
        for c in range(self._num_cols):
            x = c * self._COL_WIDTH
            self._col_header.create_rectangle(
                x, 0, x + self._COL_WIDTH, self._HEADER_HEIGHT,
                fill=self._HEADER_BG, outline=self._GRID_COLOR,
            )
            self._col_header.create_text(
                x + self._COL_WIDTH // 2, self._HEADER_HEIGHT // 2,
                text=self._col_ids[c], font=("TkDefaultFont", 9, "bold"),
            )
        self._col_header.configure(
            scrollregion=(0, 0, total_width, self._HEADER_HEIGHT),
        )

        # Заголовки строк (номера)
        for r in range(self._num_rows):
            y = r * self._ROW_HEIGHT
            self._row_header.create_rectangle(
                0, y, self._ROW_HEADER_WIDTH, y + self._ROW_HEIGHT,
                fill=self._HEADER_BG, outline=self._GRID_COLOR,
            )
            self._row_header.create_text(
                self._ROW_HEADER_WIDTH // 2, y + self._ROW_HEIGHT // 2,
                text=str(r + 1), font=("TkDefaultFont", 9),
            )
        self._row_header.configure(
            scrollregion=(0, 0, self._ROW_HEADER_WIDTH, total_height),
        )

        # Ячейки
        for r in range(self._num_rows):
            for c in range(self._num_cols):
                x = c * self._COL_WIDTH
                y = r * self._ROW_HEIGHT

                cell = sheet.cells.get((r, c))
                text = cell.value if cell else ""
                display = self._truncate_text(text)

                rect_id = self._canvas.create_rectangle(
                    x, y, x + self._COL_WIDTH, y + self._ROW_HEIGHT,
                    fill=self._NORMAL_BG, outline=self._GRID_COLOR,
                )
                text_id = self._canvas.create_text(
                    x + 4, y + self._ROW_HEIGHT // 2,
                    text=display, anchor="w", font=self._default_font,
                )

                self._cell_rects[(r, c)] = rect_id
                self._cell_texts[(r, c)] = text_id

        self._canvas.configure(scrollregion=(0, 0, total_width, total_height))

    # ── Вспомогательные методы ────────────────────────────────────────────

    def _truncate_text(self, text: str) -> str:
        if not text or self._default_font is None:
            return ""
        single = text.replace("\n", " ")
        max_px = self._COL_WIDTH - 8
        if self._default_font.measure(single) <= max_px:
            return single
        lo, hi = 0, len(single)
        while lo < hi:
            mid = (lo + hi + 1) // 2
            if self._default_font.measure(single[:mid] + "\u2026") <= max_px:
                lo = mid
            else:
                hi = mid - 1
        return single[:lo] + "\u2026" if lo > 0 else ""

    @staticmethod
    def _col_letter(col_idx: int) -> str:
        result = ""
        n = col_idx
        while True:
            result = chr(ord("A") + n % 26) + result
            n = n // 26 - 1
            if n < 0:
                break
        return result

    def _cell_from_canvas_coords(self, cx: float, cy: float) -> tuple[int, int] | None:
        col = int(cx // self._COL_WIDTH)
        row = int(cy // self._ROW_HEIGHT)
        if 0 <= row < self._num_rows and 0 <= col < self._num_cols:
            return (row, col)
        return None

    def _get_cell_data_for_display(self, row_idx: int, col_idx: int) -> CellData | None:
        delta = self._state.delta.get(row_idx, col_idx)
        sheet = self._state.sheet_data
        if delta is not None:
            orig = sheet.cells.get((row_idx, col_idx)) if sheet else None
            return CellData(
                value=delta.value if delta.value is not None else (orig.value if orig else ""),
                font_name=delta.font_name if delta.font_name is not None else (orig.font_name if orig else ""),
                font_size=delta.font_size if delta.font_size is not None else (orig.font_size if orig else 11),
                bold=delta.bold if delta.bold is not None else (orig.bold if orig else False),
                italic=delta.italic if delta.italic is not None else (orig.italic if orig else False),
                underline=delta.underline if delta.underline is not None else (orig.underline if orig else False),
                color=delta.color if delta.color is not None else (orig.color if orig else ""),
            )
        elif sheet:
            return sheet.cells.get((row_idx, col_idx))
        return None

    # ── Обновление цветов ячеек ───────────────────────────────────────────

    def _refresh_cell_colors(self) -> None:
        modified_cells: set[tuple[int, int]] = set()
        for (row, col), _ in self._state.delta.items():
            modified_cells.add((row, col))

        sel_row = self._active_cell[0] if self._active_cell else -1
        sel_col = self._active_cell[1] if self._active_cell else -1

        for r in range(self._num_rows):
            for c in range(self._num_cols):
                rect_id = self._cell_rects.get((r, c))
                if rect_id is None:
                    continue

                if r == sel_row and c == sel_col:
                    bg = self._CELL_SEL_BG
                elif r == sel_row:
                    bg = self._ROW_SEL_BG
                elif (r, c) in modified_cells:
                    bg = self._MODIFIED_BG
                else:
                    bg = self._NORMAL_BG

                self._canvas.itemconfig(rect_id, fill=bg)

    # ── Обработчики кликов ────────────────────────────────────────────────

    def _on_click_release(self, event: tk.Event) -> None:
        if self._entry_widget is not None:
            self._finish_edit(save=True)

        cx = self._canvas.canvasx(event.x)
        cy = self._canvas.canvasy(event.y)
        cell = self._cell_from_canvas_coords(cx, cy)
        if cell is None:
            self._active_cell = None
            self._refresh_cell_colors()
            return

        self._finish_edit(save=True)

        row_idx, col_idx = cell
        self._active_cell = (row_idx, col_idx)
        self._refresh_cell_colors()
        self._update_format_bar(row_idx, col_idx)
        self._update_cell_text_tab(row_idx, col_idx)

    def _on_double_click(self, event: tk.Event) -> str | None:
        cx = self._canvas.canvasx(event.x)
        cy = self._canvas.canvasy(event.y)
        cell = self._cell_from_canvas_coords(cx, cy)
        if cell is None:
            return "break"

        row_idx, col_idx = cell
        self._start_edit(row_idx, col_idx)
        return "break"

    # ── Редактирование ячейки ─────────────────────────────────────────────

    def _start_edit(self, row_idx: int, col_idx: int) -> None:
        self._finish_edit(save=True)
        self._active_cell = (row_idx, col_idx)

        cell_data = self._get_cell_data_for_display(row_idx, col_idx)
        current_value = cell_data.value if cell_data and cell_data.value else ""
        current_value_inline = current_value.replace("\n", " ")

        # Позиция в координатах Canvas
        cx = col_idx * self._COL_WIDTH
        cy = row_idx * self._ROW_HEIGHT
        # Перевод в виджетные координаты
        wx = cx - self._canvas.canvasx(0)
        wy = cy - self._canvas.canvasy(0)

        entry = tk.Entry(self._canvas, font=("TkDefaultFont", 10))
        entry.place(x=wx, y=wy, width=self._COL_WIDTH, height=self._ROW_HEIGHT)
        entry.insert(0, current_value_inline)
        entry.select_range(0, tk.END)

        self._edit_guard_until = time.monotonic() + 0.3

        entry.bind("<Return>", lambda e: self._finish_edit(save=True))
        entry.bind("<Escape>", lambda e: self._finish_edit(save=False))
        entry.bind("<FocusOut>", self._on_entry_focus_out)
        entry.bind("<Tab>", lambda e: self._finish_edit(save=True))

        self._entry_widget = entry

        entry.focus_set()
        self.after(50, self._refocus_entry)

    def _refocus_entry(self) -> None:
        self._edit_guard_until = time.monotonic() + 0.3
        if self._entry_widget is not None:
            try:
                self._entry_widget.focus_set()
            except tk.TclError:
                pass

    def _on_entry_focus_out(self, event=None) -> None:
        if time.monotonic() < self._edit_guard_until:
            return
        self._finish_edit(save=True)

    def _finish_edit(self, save: bool = True) -> None:
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
        existing = self._state.delta.get(row_idx, col_idx)
        if existing is not None:
            delta = CellDelta(
                value=new_value,
                font_name=existing.font_name,
                font_size=existing.font_size,
                bold=existing.bold,
                italic=existing.italic,
                underline=existing.underline,
                color=existing.color,
            )
        else:
            delta = CellDelta(
                value=new_value,
                font_name=None,
                font_size=None,
                bold=None,
                italic=None,
                underline=None,
                color=None,
            )
        self._state.delta.set(row_idx, col_idx, delta)

        # Обновить отображаемый текст на Canvas
        text_id = self._cell_texts.get((row_idx, col_idx))
        if text_id is not None:
            display = self._truncate_text(new_value)
            self._canvas.itemconfig(text_id, text=display)

        self._refresh_cell_colors()
        self._btn_save.config(state=tk.NORMAL)
        self._btn_apply_all.config(state=tk.NORMAL)

    # ── Обновление панелей ────────────────────────────────────────────────

    def _update_format_bar(self, row_idx: int, col_idx: int) -> None:
        cell_data = self._get_cell_data_for_display(row_idx, col_idx)
        self._format_bar.load_cell(cell_data)

    def _update_cell_text_tab(self, row_idx: int, col_idx: int) -> None:
        cell_data = self._get_cell_data_for_display(row_idx, col_idx)
        if cell_data is None:
            text = ""
        else:
            text = cell_data.value
        bp = self._find_bottom_panel()
        if bp:
            bp.show_cell_content(text)

    def _find_bottom_panel(self) -> BottomPanel | None:
        widget = self.master
        while widget is not None:
            if hasattr(widget, "_bottom_panel"):
                return widget._bottom_panel
            widget = widget.master if hasattr(widget, "master") else None
        return None

    def _save_cell_text_from_bar(self) -> None:
        if self._active_cell is None:
            return
        bp = self._find_bottom_panel()
        if not bp:
            return
        new_text = bp.get_cell_text()
        row_idx, col_idx = self._active_cell
        self._finish_edit(save=True)
        self._save_cell_value(row_idx, col_idx, new_text)

    # ── Применение форматирования ─────────────────────────────────────────

    def _apply_format(self, format_delta: CellDelta) -> None:
        if not self._active_cell:
            return

        row_idx, col_idx = self._active_cell

        existing = self._state.delta.get(row_idx, col_idx)
        if existing is not None:
            merged = CellDelta(
                value=existing.value,
                font_name=format_delta.font_name if format_delta.font_name is not None else existing.font_name,
                font_size=format_delta.font_size if format_delta.font_size is not None else existing.font_size,
                bold=format_delta.bold if format_delta.bold is not None else existing.bold,
                italic=format_delta.italic if format_delta.italic is not None else existing.italic,
                underline=format_delta.underline if format_delta.underline is not None else existing.underline,
                color=format_delta.color if format_delta.color is not None else existing.color,
            )
        else:
            merged = CellDelta(
                value=None,
                font_name=format_delta.font_name,
                font_size=format_delta.font_size,
                bold=format_delta.bold,
                italic=format_delta.italic,
                underline=format_delta.underline,
                color=format_delta.color,
            )

        self._state.delta.set(row_idx, col_idx, merged)
        self._state.log.add(
            f"Формат ячейки ({row_idx},{col_idx}): "
            f"шрифт={merged.font_name} размер={merged.font_size} "
            f"ж={merged.bold} к={merged.italic} п={merged.underline} ц={merged.color}"
        )

        self._refresh_cell_colors()
        self._btn_save.config(state=tk.NORMAL)
        self._btn_apply_all.config(state=tk.NORMAL)

        self._update_format_bar(row_idx, col_idx)
        self._update_cell_text_tab(row_idx, col_idx)

    # ── Кнопки действий ───────────────────────────────────────────────────

    def _on_save(self) -> None:
        self._finish_edit(save=True)

        if self._state.current_file is None:
            return

        delta = self._state.delta
        n_changes = sum(1 for _ in delta.items())
        self._state.log.add(f"Сохранение файла: {self._state.current_file.path.name} ({n_changes} ячеек изменено)")
        processor = Processor()
        ok = processor.apply_and_save(
            self._state.current_file.path,
            delta,
            self._state.log,
        )

        if ok:
            self._state.log.add("Файл успешно сохранён")
            messagebox.showinfo("Сохранено", "Файл успешно сохранён.")
        else:
            messagebox.showerror("Ошибка", "Не удалось сохранить файл. Подробности в логе.")

        self._refresh_bottom_panel()

    def _on_apply_all(self) -> None:
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

        self._state.log.add(f"Массовое применение: {n} файлов (группа {key} строк)")

        progress_dialog = ProgressDialog(self, total=n)

        def progress_callback(done: int, total: int) -> None:
            progress_dialog.update(done, total)
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

        if result.error_count == 0:
            self._state.log.add(f"Массовое применение завершено: {result.success_count} файлов")
            summary = f"Успешно обновлено файлов: {result.success_count}."
        else:
            summary = (
                f"Успешно обновлено: {result.success_count} файл(ов).\n"
                f"С ошибками: {result.error_count} файл(ов).\n"
                "Подробности в логе."
            )
        messagebox.showinfo("Результат", summary)

        self._refresh_bottom_panel()

    def _on_close_editor(self) -> None:
        self._finish_edit(save=True)

        if not self._state.delta.is_empty():
            confirmed = messagebox.askyesno(
                "Несохранённые изменения",
                "Есть несохранённые изменения. Закрыть без сохранения?",
            )
            if not confirmed:
                return

        self._on_close()

    def _refresh_bottom_panel(self) -> None:
        bp = self._find_bottom_panel()
        if bp:
            bp.refresh_log()


class MainWindow(tk.Tk):
    """
    Корневое окно приложения.
    Содержит StatsFrame или EditorFrame (переключение через pack/forget).
    Внизу всегда отображается BottomPanel с вкладками.
    """

    def __init__(self) -> None:
        super().__init__()
        self.title("tablered")
        self.geometry("900x600")
        self.minsize(600, 400)

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

        self._state.log.add("Приложение запущено")
        self._run_scan()

        self.protocol("WM_DELETE_WINDOW", self._on_close_app)

    def _build_ui(self) -> None:
        self._content_frame = tk.Frame(self)
        self._content_frame.pack(fill=tk.BOTH, expand=True)

        self._stats_frame = StatsFrame(
            self._content_frame,
            state=self._state,
            on_open_sample=self._show_editor,
            on_folder_changed=self._on_folder_changed,
        )
        self._stats_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Separator(self, orient=tk.HORIZONTAL).pack(fill=tk.X)

        self._bottom_panel = BottomPanel(self, self._state.log)
        self._bottom_panel.pack(fill=tk.BOTH, side=tk.BOTTOM)

    def _on_close_app(self) -> None:
        self._state.log.add("Приложение закрывается")
        self._state.log.close()
        self.destroy()

    def _run_scan(self) -> None:
        folder = self._state.target_folder
        if not folder.exists() or not folder.is_dir():
            self._state.log.add(f"Папка не существует: {folder}")
            self._bottom_panel.refresh_log()
            self._state.scan_results = []
            self._state.groups = {}
            self._stats_frame.refresh()
            return

        self._state.log.add(f"Сканирование папки: {folder}")
        self._state.scan_results = self._scanner.scan(folder, self._state.log)
        self._group_manager.build(self._state.scan_results)
        self._state.groups = self._group_manager.get_groups()
        n_files = len(self._state.scan_results)
        n_groups = len(self._state.groups)
        self._state.log.add(f"Найдено файлов: {n_files}, групп: {n_groups}")
        self._stats_frame.refresh()
        self._bottom_panel.refresh_log()

    def _on_folder_changed(self, new_folder: Path) -> None:
        self._state.log.add(f"Смена папки: {new_folder}")
        self._run_scan()

    def _show_editor(self, file_info: "FileInfo") -> None:
        self._state.current_file = file_info

        self._state.log.add(f"Открытие файла: {file_info.path.name}")
        reader = OdsReader()
        try:
            sheet_data = reader.load(file_info.path)
        except Exception as e:
            self._state.log.add(f"Ошибка открытия файла {file_info.path.name}: {e}")
            self._bottom_panel.refresh_log()
            messagebox.showerror("Ошибка", f"Не удалось открыть файл:\n{e}")
            return

        self._state.sheet_data = sheet_data
        self._state.delta.clear()
        self._state.log.add(
            f"Файл загружен: {sheet_data.max_row + 1} строк, {sheet_data.max_col + 1} столбцов"
        )

        self._stats_frame.pack_forget()

        self._editor_frame = EditorFrame(
            self._content_frame,
            state=self._state,
            on_close=self.show_stats,
        )
        self._editor_frame.pack(fill=tk.BOTH, expand=True)

    def show_stats(self) -> None:
        if hasattr(self, "_editor_frame") and self._editor_frame is not None:
            self._editor_frame.pack_forget()
            self._editor_frame.destroy()
            self._editor_frame = None

        self._state.current_file = None
        self._state.sheet_data = None

        self._stats_frame.pack(fill=tk.BOTH, expand=True)


# ===========================================================================
# Точка входа
# ===========================================================================

if __name__ == "__main__":
    _check_deps_and_install()
    app = MainWindow()
    app.mainloop()
    app._state.log.close()
