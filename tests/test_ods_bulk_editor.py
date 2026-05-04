"""
Тесты для ODS Bulk Editor.

Структура:
- Фикстуры pytest для базовых компонентов
- Заглушки для будущих unit-тестов (задачи 2–5, 12)
- Property-based тесты с использованием Hypothesis
"""

from __future__ import annotations

# ---------------------------------------------------------------------------
# Стандартная библиотека
# ---------------------------------------------------------------------------
import tempfile
from pathlib import Path

# ---------------------------------------------------------------------------
# Тестовый фреймворк
# ---------------------------------------------------------------------------
import pytest

# ---------------------------------------------------------------------------
# Hypothesis (property-based testing)
# ---------------------------------------------------------------------------
from hypothesis import given, settings
import hypothesis.strategies as st

# ---------------------------------------------------------------------------
# Тестируемые классы
# ---------------------------------------------------------------------------
from ods_bulk_editor import (
    FileInfo,
    CellData,
    SheetData,
    CellDelta,
    BulkResult,
    AppState,
    LogStore,
    DeltaStore,
    Scanner,
    GroupManager,
    OdsReader,
    Processor,
    count_data_rows,
)

# ---------------------------------------------------------------------------
# odfpy helpers for building test sheets
# ---------------------------------------------------------------------------
from odf.opendocument import OpenDocumentSpreadsheet
from odf.table import Table, TableRow, TableCell
from odf.text import P


# ===========================================================================
# Фикстуры
# ===========================================================================

@pytest.fixture
def empty_delta_store() -> DeltaStore:
    """Пустой DeltaStore без изменений."""
    return DeltaStore()


@pytest.fixture
def empty_log_store() -> LogStore:
    """Пустой LogStore без записей."""
    return LogStore()


@pytest.fixture
def sample_cell_delta() -> CellDelta:
    """Пример CellDelta с заполненными полями."""
    return CellDelta(
        value="Hello",
        font_name="Arial",
        font_size=12,
        bold=True,
        italic=False,
        color="#FF0000",
    )


@pytest.fixture
def sample_file_info(tmp_path: Path) -> FileInfo:
    """Пример FileInfo с временным путём."""
    fake_path = tmp_path / "sample.ods"
    fake_path.touch()
    return FileInfo(path=fake_path, row_count=10)


# ===========================================================================
# Unit-тесты: LogStore
# ===========================================================================

class TestLogStore:
    def test_initial_empty(self, empty_log_store: LogStore) -> None:
        assert empty_log_store.get_all() == []

    def test_add_single_message(self, empty_log_store: LogStore) -> None:
        empty_log_store.add("error: file not found")
        assert empty_log_store.get_all() == ["error: file not found"]

    def test_add_multiple_messages(self, empty_log_store: LogStore) -> None:
        empty_log_store.add("msg1")
        empty_log_store.add("msg2")
        assert empty_log_store.get_all() == ["msg1", "msg2"]

    def test_clear(self, empty_log_store: LogStore) -> None:
        empty_log_store.add("msg1")
        empty_log_store.clear()
        assert empty_log_store.get_all() == []

    def test_get_all_returns_copy(self, empty_log_store: LogStore) -> None:
        """get_all() должен возвращать независимую копию списка."""
        empty_log_store.add("msg1")
        result = empty_log_store.get_all()
        result.append("injected")
        assert empty_log_store.get_all() == ["msg1"]


# ===========================================================================
# Unit-тесты: DeltaStore
# ===========================================================================

class TestDeltaStore:
    def test_initial_empty(self, empty_delta_store: DeltaStore) -> None:
        assert empty_delta_store.is_empty()

    def test_get_nonexistent_returns_none(self, empty_delta_store: DeltaStore) -> None:
        assert empty_delta_store.get(0, 0) is None

    def test_set_and_get(
        self, empty_delta_store: DeltaStore, sample_cell_delta: CellDelta
    ) -> None:
        empty_delta_store.set(1, 2, sample_cell_delta)
        result = empty_delta_store.get(1, 2)
        assert result is sample_cell_delta

    def test_set_makes_non_empty(
        self, empty_delta_store: DeltaStore, sample_cell_delta: CellDelta
    ) -> None:
        empty_delta_store.set(0, 0, sample_cell_delta)
        assert not empty_delta_store.is_empty()

    def test_overwrite(self, empty_delta_store: DeltaStore) -> None:
        """Повторная запись в ту же ячейку перезаписывает значение."""
        delta1 = CellDelta(value="first", font_name=None, font_size=None,
                           bold=None, italic=None, color=None)
        delta2 = CellDelta(value="second", font_name=None, font_size=None,
                           bold=None, italic=None, color=None)
        empty_delta_store.set(0, 0, delta1)
        empty_delta_store.set(0, 0, delta2)
        assert empty_delta_store.get(0, 0).value == "second"

    def test_items_iteration(
        self, empty_delta_store: DeltaStore, sample_cell_delta: CellDelta
    ) -> None:
        empty_delta_store.set(3, 5, sample_cell_delta)
        items = list(empty_delta_store.items())
        assert len(items) == 1
        assert items[0] == ((3, 5), sample_cell_delta)

    def test_clear(
        self, empty_delta_store: DeltaStore, sample_cell_delta: CellDelta
    ) -> None:
        empty_delta_store.set(0, 0, sample_cell_delta)
        empty_delta_store.clear()
        assert empty_delta_store.is_empty()
        assert empty_delta_store.get(0, 0) is None


# ===========================================================================
# Property-based тест: DeltaStore round-trip (Property 5)
# ===========================================================================

# Feature: ods-bulk-editor, Property 5: Дельта round-trip — значение ячейки
@given(
    row=st.integers(min_value=0, max_value=10_000),
    col=st.integers(min_value=0, max_value=10_000),
    value=st.text(),
)
@settings(max_examples=100)
def test_delta_store_roundtrip(row: int, col: int, value: str) -> None:
    """
    **Validates: Requirements 5.3**

    Для любой ячейки (row, col) и любого строкового значения v,
    если записать DeltaStore.set(row, col, CellDelta(value=v, ...)),
    то DeltaStore.get(row, col).value должно вернуть v.
    """
    store = DeltaStore()
    delta = CellDelta(
        value=value,
        font_name=None,
        font_size=None,
        bold=None,
        italic=None,
        color=None,
    )
    store.set(row, col, delta)
    result = store.get(row, col)
    assert result is not None
    assert result.value == value


# ===========================================================================
# Заглушки для будущих тестов (задачи 2–5, 12)
# ===========================================================================

# ===========================================================================
# Helpers for building odfpy sheets in tests
# ===========================================================================

def _make_sheet_with_rows(rows_data: list[list[str]]) -> Table:
    """
    Создаёт odfpy Table (лист) с заданными строками и ячейками.
    rows_data — список строк, каждая строка — список строковых значений ячеек.
    Пустая строка "" означает пустую ячейку.
    """
    sheet = Table(name="Sheet1")
    for row_data in rows_data:
        row = TableRow()
        for cell_value in row_data:
            cell = TableCell()
            if cell_value:
                p = P(text=cell_value)
                cell.addElement(p)
            row.addElement(cell)
        sheet.addElement(row)
    return sheet


# ===========================================================================
# Unit-тесты: count_data_rows
# ===========================================================================

class TestCountDataRows:
    def test_count_data_rows_empty_sheet(self) -> None:
        """Пустой лист возвращает 0."""
        sheet = _make_sheet_with_rows([])
        assert count_data_rows(sheet) == 0

    def test_count_data_rows_trailing_empty(self) -> None:
        """Trailing empty rows не считаются."""
        # 2 строки с данными, 3 пустые строки в конце
        sheet = _make_sheet_with_rows([
            ["value1", "value2"],
            ["value3", ""],
            ["", ""],
            ["", ""],
            ["", ""],
        ])
        assert count_data_rows(sheet) == 2

    def test_count_data_rows_all_empty(self) -> None:
        """Лист только из пустых строк возвращает 0."""
        sheet = _make_sheet_with_rows([
            ["", ""],
            ["", ""],
        ])
        assert count_data_rows(sheet) == 0

    def test_count_data_rows_no_trailing(self) -> None:
        """Все строки с данными — считаются все."""
        sheet = _make_sheet_with_rows([
            ["a"],
            ["b"],
            ["c"],
        ])
        assert count_data_rows(sheet) == 3

    def test_count_data_rows_empty_in_middle(self) -> None:
        """Пустые строки в середине считаются (не trailing)."""
        sheet = _make_sheet_with_rows([
            ["a"],
            ["", ""],
            ["c"],
        ])
        # 2 строки с данными, пустая в середине не считается как data row
        assert count_data_rows(sheet) == 2


# ===========================================================================
# Unit-тесты: Scanner
# ===========================================================================

class TestScanner:
    def test_scanner_finds_ods_files(self, tmp_path: Path) -> None:
        """Scanner находит .ods файлы в папке."""
        # Создаём реальный .ods файл
        doc = OpenDocumentSpreadsheet()
        sheet = Table(name="Sheet1")
        row = TableRow()
        cell = TableCell()
        p = P(text="data")
        cell.addElement(p)
        row.addElement(cell)
        sheet.addElement(row)
        doc.spreadsheet.addElement(sheet)
        ods_path = tmp_path / "test.ods"
        doc.save(str(ods_path))

        log = LogStore()
        scanner = Scanner()
        results = scanner.scan(tmp_path, log)
        assert len(results) == 1
        assert results[0].path == ods_path
        assert results[0].row_count == 1

    def test_scanner_ignores_non_ods(self, tmp_path: Path) -> None:
        """Scanner игнорирует файлы с другими расширениями."""
        (tmp_path / "file.xlsx").touch()
        (tmp_path / "file.txt").touch()
        (tmp_path / "file.csv").touch()

        log = LogStore()
        scanner = Scanner()
        results = scanner.scan(tmp_path, log)
        assert results == []

    def test_scanner_case_insensitive(self, tmp_path: Path) -> None:
        """Scanner находит .ODS файлы (без учёта регистра)."""
        doc = OpenDocumentSpreadsheet()
        sheet = Table(name="Sheet1")
        doc.spreadsheet.addElement(sheet)
        ods_path = tmp_path / "test.ODS"
        doc.save(str(ods_path))

        log = LogStore()
        scanner = Scanner()
        results = scanner.scan(tmp_path, log)
        assert len(results) == 1

    def test_scanner_no_recursion(self, tmp_path: Path) -> None:
        """Scanner не заходит в подпапки."""
        subdir = tmp_path / "subdir"
        subdir.mkdir()
        doc = OpenDocumentSpreadsheet()
        sheet = Table(name="Sheet1")
        doc.spreadsheet.addElement(sheet)
        doc.save(str(subdir / "nested.ods"))

        log = LogStore()
        scanner = Scanner()
        results = scanner.scan(tmp_path, log)
        assert results == []

    def test_scanner_corrupted_file_logged(self, tmp_path: Path) -> None:
        """Повреждённый файл пропускается и записывается в лог."""
        bad_file = tmp_path / "bad.ods"
        bad_file.write_bytes(b"this is not a valid ods file")

        log = LogStore()
        scanner = Scanner()
        results = scanner.scan(tmp_path, log)
        assert results == []
        assert len(log.get_all()) == 1


# ===========================================================================
# Property-based тест: Scanner (фильтрация файлов) — Property 1
# ===========================================================================

def _create_ods_file(path: Path) -> None:
    """Создаёт минимальный валидный .ods файл."""
    doc = OpenDocumentSpreadsheet()
    sheet = Table(name="Sheet1")
    doc.spreadsheet.addElement(sheet)
    doc.save(str(path))


# Feature: ods-bulk-editor, Property 1: Сканер находит все .ods-файлы и только их
@given(
    extensions=st.lists(
        st.sampled_from(['.ods', '.ODS', '.xlsx', '.txt', '.csv', '.Ods']),
        min_size=0,
        max_size=10,
    )
)
@settings(max_examples=100)
def test_scanner_ods_filter(extensions: list[str]) -> None:
    """
    **Validates: Requirements 2.1**

    Для любого набора файлов с различными расширениями, Scanner.scan возвращает
    ровно те файлы, расширение которых совпадает с .ods без учёта регистра.
    """
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        # Создаём файлы с заданными расширениями
        created_paths: list[Path] = []
        for i, ext in enumerate(extensions):
            file_path = tmp_path / f"file_{i}{ext}"
            if ext.lower() == '.ods':
                _create_ods_file(file_path)
            else:
                file_path.touch()
            created_paths.append(file_path)

        expected_ods = {p for p in created_paths if p.suffix.lower() == '.ods'}

        log = LogStore()
        scanner = Scanner()
        results = scanner.scan(tmp_path, log)
        result_paths = {r.path for r in results}

        assert result_paths == expected_ods


# ===========================================================================
# Property-based тест: count_data_rows trailing invariant — Property 2
# ===========================================================================

# Feature: ods-bulk-editor, Property 2: Подсчёт строк с данными игнорирует trailing empty rows
@given(
    rows_data=st.lists(
        st.lists(st.text(min_size=0, max_size=20), min_size=1, max_size=5),
        min_size=0,
        max_size=10,
    ),
    trailing_count=st.integers(min_value=0, max_value=5),
)
@settings(max_examples=100)
def test_count_data_rows_trailing_invariant(
    rows_data: list[list[str]], trailing_count: int
) -> None:
    """
    **Validates: Requirements 2.3**

    Для любого листа ODS, добавление произвольного количества полностью пустых
    строк в конец не изменяет результат count_data_rows.
    """
    sheet_without_trailing = _make_sheet_with_rows(rows_data)
    count_without = count_data_rows(sheet_without_trailing)

    # Добавляем trailing empty rows
    trailing_rows = [[""] * max(len(r) for r in rows_data) if rows_data else [""]] * trailing_count
    sheet_with_trailing = _make_sheet_with_rows(rows_data + trailing_rows)
    count_with = count_data_rows(sheet_with_trailing)

    assert count_without == count_with


# TODO (задача 2.1): test_count_data_rows_empty_sheet  ← реализовано выше в TestCountDataRows
# TODO (задача 2.1): test_count_data_rows_trailing_empty  ← реализовано выше в TestCountDataRows
# TODO (задача 2.2): test_scanner_ods_filter  (Property 1)  ← реализовано выше
# TODO (задача 2.3): test_count_data_rows_trailing_invariant  (Property 2)  ← реализовано выше

# ===========================================================================
# Unit-тесты: GroupManager (задача 3.1)
# ===========================================================================

class TestGroupManager:
    def test_group_manager_single_group(self, tmp_path: Path) -> None:
        """Все файлы с одинаковым row_count попадают в одну группу."""
        files = [
            FileInfo(path=tmp_path / f"file{i}.ods", row_count=5)
            for i in range(4)
        ]
        gm = GroupManager()
        gm.build(files)
        groups = gm.get_groups()
        assert list(groups.keys()) == [5]
        assert len(groups[5]) == 4

    def test_group_manager_multiple_groups(self, tmp_path: Path) -> None:
        """Файлы с разным row_count разбиваются по отдельным группам."""
        files = [
            FileInfo(path=tmp_path / "a.ods", row_count=3),
            FileInfo(path=tmp_path / "b.ods", row_count=7),
            FileInfo(path=tmp_path / "c.ods", row_count=3),
            FileInfo(path=tmp_path / "d.ods", row_count=10),
        ]
        gm = GroupManager()
        gm.build(files)
        groups = gm.get_groups()
        assert set(groups.keys()) == {3, 7, 10}
        assert len(groups[3]) == 2
        assert len(groups[7]) == 1
        assert len(groups[10]) == 1

    def test_group_manager_empty(self) -> None:
        """Пустой список файлов даёт пустую группировку."""
        gm = GroupManager()
        gm.build([])
        assert gm.get_groups() == {}

    def test_group_manager_get_group(self, tmp_path: Path) -> None:
        """get_group возвращает правильный список файлов для ключа."""
        f1 = FileInfo(path=tmp_path / "a.ods", row_count=5)
        f2 = FileInfo(path=tmp_path / "b.ods", row_count=5)
        gm = GroupManager()
        gm.build([f1, f2])
        result = gm.get_group(5)
        assert result == [f1, f2]

    def test_group_manager_get_group_missing(self) -> None:
        """get_group для несуществующего ключа возвращает пустой список."""
        gm = GroupManager()
        gm.build([])
        assert gm.get_group(99) == []

    def test_group_manager_sorted_keys(self, tmp_path: Path) -> None:
        """get_groups возвращает словарь, отсортированный по ключу."""
        files = [
            FileInfo(path=tmp_path / "c.ods", row_count=30),
            FileInfo(path=tmp_path / "a.ods", row_count=10),
            FileInfo(path=tmp_path / "b.ods", row_count=20),
        ]
        gm = GroupManager()
        gm.build(files)
        keys = list(gm.get_groups().keys())
        assert keys == sorted(keys)


# ===========================================================================
# Property-based тест: GroupManager — покрытие без потерь (Property 3)
# ===========================================================================

# Feature: ods-bulk-editor, Property 3: Группировка покрывает все файлы без потерь
@given(
    row_counts=st.lists(st.integers(min_value=1, max_value=100), min_size=0, max_size=50),
)
@settings(max_examples=100)
def test_group_manager_coverage(row_counts: list[int]) -> None:
    """
    **Validates: Requirements 3.1**

    Для любого списка FileInfo, сумма длин всех групп в GroupManager
    должна быть равна длине исходного списка, а каждый файл должен
    присутствовать ровно в одной группе.
    """
    files = [FileInfo(path=Path(f"file_{i}.ods"), row_count=rc) for i, rc in enumerate(row_counts)]
    gm = GroupManager()
    gm.build(files)
    groups = gm.get_groups()

    # Сумма длин всех групп равна числу файлов
    total = sum(len(v) for v in groups.values())
    assert total == len(files)

    # Каждый файл присутствует ровно в одной группе
    all_files_in_groups = [f for group in groups.values() for f in group]
    assert sorted(all_files_in_groups, key=lambda f: str(f.path)) == sorted(files, key=lambda f: str(f.path))


# ===========================================================================
# Property-based тест: сортировка статистики по возрастанию (Property 4)
# ===========================================================================

# Feature: ods-bulk-editor, Property 4: Сортировка статистики по возрастанию
@given(
    raw_groups=st.dictionaries(
        st.integers(min_value=1, max_value=1000),
        st.lists(st.just(None)),
    )
)
@settings(max_examples=100)
def test_stats_sorted_order(raw_groups: dict) -> None:
    """
    **Validates: Requirements 3.3**

    Для любой группировки файлов, список ключей (количество строк),
    возвращаемый get_groups(), должен быть отсортирован по возрастанию.
    """
    files = [
        FileInfo(path=Path(f"file_{key}_{i}.ods"), row_count=key)
        for key, items in raw_groups.items()
        for i in range(len(items))
    ]
    gm = GroupManager()
    gm.build(files)
    keys = list(gm.get_groups().keys())
    assert keys == sorted(keys)

# ===========================================================================
# Helpers for building real ODS files in tests
# ===========================================================================

def _make_ods_file(tmp_path: Path, filename: str, rows_data: list[list[str]]) -> Path:
    """
    Создаёт реальный .ods файл с заданными строками и ячейками.
    Возвращает путь к созданному файлу.
    """
    from odf.namespaces import OFFICENS
    doc = OpenDocumentSpreadsheet()
    sheet = Table(name="Sheet1")
    for row_data in rows_data:
        row = TableRow()
        for cell_value in row_data:
            cell = TableCell()
            if cell_value:
                cell.setAttrNS(OFFICENS, "value-type", "string")
                p = P(text=cell_value)
                cell.addElement(p)
            row.addElement(cell)
        sheet.addElement(row)
    doc.spreadsheet.addElement(sheet)
    file_path = tmp_path / filename
    doc.save(str(file_path))
    return file_path


# ===========================================================================
# Unit-тесты: Processor (задача 5.1)
# ===========================================================================

class TestProcessor:
    def test_processor_apply_value(self, tmp_path: Path) -> None:
        """Значение ячейки меняется после применения дельты."""
        ods_path = _make_ods_file(tmp_path, "test.ods", [
            ["original_value", "other"],
            ["row2col1", "row2col2"],
        ])

        delta = DeltaStore()
        delta.set(0, 0, CellDelta(
            value="new_value",
            font_name=None, font_size=None, bold=None, italic=None, color=None,
        ))

        log = LogStore()
        processor = Processor()
        result = processor.apply_and_save(ods_path, delta, log)
        assert result is True

        # Re-read and verify
        reader = OdsReader()
        sheet_data = reader.load(ods_path)
        assert sheet_data.cells[(0, 0)].value == "new_value"
        # Other cell unchanged
        assert sheet_data.cells[(0, 1)].value == "other"

    def test_processor_apply_formatting(self, tmp_path: Path) -> None:
        """Форматирование применяется корректно."""
        ods_path = _make_ods_file(tmp_path, "fmt_test.ods", [
            ["hello"],
        ])

        delta = DeltaStore()
        delta.set(0, 0, CellDelta(
            value=None,
            font_name="Arial",
            font_size=14,
            bold=True,
            italic=False,
            color="#FF0000",
        ))

        log = LogStore()
        processor = Processor()
        result = processor.apply_and_save(ods_path, delta, log)
        assert result is True

        # Re-read and verify formatting
        reader = OdsReader()
        sheet_data = reader.load(ods_path)
        cell = sheet_data.cells[(0, 0)]
        assert cell.font_name == "Arial"
        assert cell.font_size == 14
        assert cell.bold is True
        assert cell.color == "#FF0000"

    def test_processor_skip_on_error(self, tmp_path: Path) -> None:
        """Повреждённый файл пропускается, остальные обрабатываются."""
        # Create one valid file
        valid_path = _make_ods_file(tmp_path, "valid.ods", [["data"]])
        # Create one corrupted file
        bad_path = tmp_path / "bad.ods"
        bad_path.write_bytes(b"not a valid ods file")

        delta = DeltaStore()
        delta.set(0, 0, CellDelta(
            value="updated",
            font_name=None, font_size=None, bold=None, italic=None, color=None,
        ))

        log = LogStore()
        processor = Processor()

        files = [
            FileInfo(path=bad_path, row_count=0),
            FileInfo(path=valid_path, row_count=1),
        ]

        def noop_callback(done: int, total: int) -> None:
            pass

        result = processor.apply_bulk(files, delta, log, noop_callback)

        # Bad file should be counted as error, valid file as success
        assert result.error_count == 1
        assert result.success_count == 1
        # Error was logged
        assert len(log.get_all()) >= 1

        # Valid file was updated
        reader = OdsReader()
        sheet_data = reader.load(valid_path)
        assert sheet_data.cells[(0, 0)].value == "updated"

    def test_bulk_result_counts(self, tmp_path: Path) -> None:
        """Счётчики success_count и error_count корректны."""
        # Create 3 valid files and 2 corrupted files
        valid_files = []
        for i in range(3):
            p = _make_ods_file(tmp_path, f"valid_{i}.ods", [["data"]])
            valid_files.append(FileInfo(path=p, row_count=1))

        bad_files = []
        for i in range(2):
            p = tmp_path / f"bad_{i}.ods"
            p.write_bytes(b"corrupted")
            bad_files.append(FileInfo(path=p, row_count=0))

        all_files = valid_files + bad_files

        delta = DeltaStore()
        log = LogStore()
        processor = Processor()

        progress_calls: list[tuple[int, int]] = []

        def track_progress(done: int, total: int) -> None:
            progress_calls.append((done, total))

        result = processor.apply_bulk(all_files, delta, log, track_progress)

        assert result.success_count == 3
        assert result.error_count == 2
        assert result.success_count + result.error_count == len(all_files)
        # Progress callback called once per file
        assert len(progress_calls) == len(all_files)
        # Last call should be (total, total)
        assert progress_calls[-1] == (len(all_files), len(all_files))


# ===========================================================================
# Property-based тест: незатронутые ячейки (Property 6) — задача 5.2
# ===========================================================================

# Feature: ods-bulk-editor, Property 6: Применение дельты не затрагивает незадействованные ячейки
@given(
    grid=st.lists(
        st.lists(
            st.text(min_size=1, max_size=10, alphabet=st.characters(whitelist_categories=("Lu", "Ll", "Nd"))),
            min_size=1,
            max_size=4,
        ),
        min_size=1,
        max_size=4,
    ),
    delta_value=st.text(min_size=1, max_size=10, alphabet=st.characters(whitelist_categories=("Lu", "Ll", "Nd"))),
)
@settings(max_examples=100)
def test_processor_untouched_cells(
    grid: list[list[str]],
    delta_value: str,
) -> None:
    """
    **Validates: Requirements 8.6**

    Применение дельты не затрагивает незадействованные ячейки.
    Для любого ODS-файла и любой DeltaStore, содержащей изменения только для
    подмножества ячеек S, после применения дельты все ячейки вне S сохраняют
    исходные значения.
    """
    import tempfile

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        ods_path = _make_ods_file(tmp_path, "test.ods", grid)

        # Read original values
        reader = OdsReader()
        original = reader.load(ods_path)

        # Pick only the first cell (0, 0) as the delta target
        delta = DeltaStore()
        delta.set(0, 0, CellDelta(
            value=delta_value,
            font_name=None, font_size=None, bold=None, italic=None, color=None,
        ))

        log = LogStore()
        processor = Processor()
        processor.apply_and_save(ods_path, delta, log)

        # Re-read
        updated = reader.load(ods_path)

        # All cells except (0, 0) must be unchanged
        for (row, col), orig_cell in original.cells.items():
            if (row, col) == (0, 0):
                continue
            assert (row, col) in updated.cells, f"Cell ({row},{col}) disappeared after apply"
            updated_cell = updated.cells[(row, col)]
            assert updated_cell.value == orig_cell.value, (
                f"Cell ({row},{col}) value changed: {orig_cell.value!r} -> {updated_cell.value!r}"
            )


# ===========================================================================
# Property-based тест: форматирование round-trip (Property 7) — задача 5.3
# ===========================================================================

# Feature: ods-bulk-editor, Property 7: Форматирование round-trip
@given(
    font_name=st.text(min_size=1, max_size=20, alphabet=st.characters(whitelist_categories=("Lu", "Ll"))),
    font_size=st.integers(min_value=6, max_value=72),
    bold=st.booleans(),
    italic=st.booleans(),
    color=st.from_regex(r'#[0-9A-F]{6}', fullmatch=True),
)
@settings(max_examples=100)
def test_formatting_roundtrip(
    font_name: str,
    font_size: int,
    bold: bool,
    italic: bool,
    color: str,
) -> None:
    """
    **Validates: Requirements 6.2, 6.4**

    Для любой ячейки и любого набора параметров форматирования, если записать их
    в DeltaStore, применить к файлу через Processor и перечитать через OdsReader,
    полученные параметры форматирования должны совпадать с записанными.
    """
    import tempfile

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        ods_path = _make_ods_file(tmp_path, "fmt.ods", [["hello"]])

        delta = DeltaStore()
        delta.set(0, 0, CellDelta(
            value=None,
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            italic=italic,
            color=color,
        ))

        log = LogStore()
        processor = Processor()
        processor.apply_and_save(ods_path, delta, log)

        reader = OdsReader()
        sheet_data = reader.load(ods_path)

        assert (0, 0) in sheet_data.cells
        cell = sheet_data.cells[(0, 0)]
        assert cell.font_name == font_name, f"font_name: {cell.font_name!r} != {font_name!r}"
        assert cell.font_size == font_size, f"font_size: {cell.font_size} != {font_size}"
        assert cell.bold == bold, f"bold: {cell.bold} != {bold}"
        assert cell.italic == italic, f"italic: {cell.italic} != {italic}"
        assert cell.color == color, f"color: {cell.color!r} != {color!r}"


# ===========================================================================
# Property-based тест: счётчик результатов массового применения (Property 8) — задача 5.4
# ===========================================================================

# Feature: ods-bulk-editor, Property 8: Массовое применение — счётчик результатов
@given(
    file_count=st.integers(min_value=0, max_value=10),
)
@settings(max_examples=100)
def test_bulk_result_sum(file_count: int) -> None:
    """
    **Validates: Requirements 8.4, 8.5**

    Для любого списка файлов группы, BulkResult.success_count + BulkResult.error_count
    должно быть равно количеству файлов в списке.
    """
    import tempfile

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)

        files: list[FileInfo] = []
        for i in range(file_count):
            p = _make_ods_file(tmp_path, f"file_{i}.ods", [["data"]])
            files.append(FileInfo(path=p, row_count=1))

        delta = DeltaStore()
        log = LogStore()
        processor = Processor()

        def noop(done: int, total: int) -> None:
            pass

        result = processor.apply_bulk(files, delta, log, noop)

        assert result.success_count + result.error_count == file_count


# ===========================================================================
# Интеграционный тест: сквозной сценарий scan → open → edit → save → reload (задача 12.1)
# ===========================================================================

class TestIntegrationScanEditSaveReload:
    """
    Сквозной интеграционный тест: создание реального .ods-файла,
    сканирование, открытие, редактирование, сохранение, перечитывание.

    Validates: Requirements 2.1, 2.2, 5.3, 7.1, 8.2
    """

    def test_scan_open_edit_save_reload(self, tmp_path: Path) -> None:
        """
        Полный сквозной сценарий:
        1. Создать реальный .ods-файл с известным содержимым через odfpy.
        2. Запустить Scanner.scan() на временной директории.
        3. Убедиться, что файл найден с правильным row_count.
        4. Загрузить файл через OdsReader.
        5. Создать DeltaStore с изменением значения ячейки.
        6. Применить через Processor.apply_and_save().
        7. Перечитать через OdsReader.
        8. Убедиться, что изменённая ячейка содержит новое значение.
        9. Убедиться, что остальные ячейки не изменились.
        """
        # ── Шаг 1: создать реальный .ods-файл с известным содержимым ──────
        rows_data = [
            ["Имя",    "Возраст", "Город"],
            ["Алиса",  "30",      "Москва"],
            ["Борис",  "25",      "Питер"],
        ]
        ods_path = _make_ods_file(tmp_path, "sample.ods", rows_data)

        # ── Шаг 2: запустить Scanner.scan() ─────────────────────────────────
        log = LogStore()
        scanner = Scanner()
        results = scanner.scan(tmp_path, log)

        # ── Шаг 3: убедиться, что файл найден с правильным row_count ────────
        assert len(results) == 1, f"Ожидался 1 файл, найдено: {len(results)}"
        file_info = results[0]
        assert file_info.path == ods_path
        assert file_info.row_count == 3, (
            f"Ожидалось 3 строки с данными, получено: {file_info.row_count}"
        )

        # ── Шаг 4: загрузить файл через OdsReader ───────────────────────────
        reader = OdsReader()
        original_data = reader.load(ods_path)

        # Проверяем, что исходные данные загружены корректно
        assert original_data.cells[(0, 0)].value == "Имя"
        assert original_data.cells[(1, 0)].value == "Алиса"
        assert original_data.cells[(2, 1)].value == "25"

        # ── Шаг 5: создать DeltaStore с изменением значения ячейки ──────────
        delta = DeltaStore()
        # Изменяем ячейку (1, 0): "Алиса" → "Виктор"
        delta.set(1, 0, CellDelta(
            value="Виктор",
            font_name=None,
            font_size=None,
            bold=None,
            italic=None,
            color=None,
        ))

        # ── Шаг 6: применить через Processor.apply_and_save() ───────────────
        processor = Processor()
        success = processor.apply_and_save(ods_path, delta, log)
        assert success is True, f"apply_and_save вернул False; лог: {log.get_all()}"

        # ── Шаг 7: перечитать через OdsReader ───────────────────────────────
        updated_data = reader.load(ods_path)

        # ── Шаг 8: убедиться, что изменённая ячейка содержит новое значение ─
        assert updated_data.cells[(1, 0)].value == "Виктор", (
            f"Ожидалось 'Виктор', получено: {updated_data.cells[(1, 0)].value!r}"
        )

        # ── Шаг 9: убедиться, что остальные ячейки не изменились ────────────
        unchanged_cells = [
            (0, 0, "Имя"),
            (0, 1, "Возраст"),
            (0, 2, "Город"),
            (1, 1, "30"),
            (1, 2, "Москва"),
            (2, 0, "Борис"),
            (2, 1, "25"),
            (2, 2, "Питер"),
        ]
        for row, col, expected_value in unchanged_cells:
            actual = updated_data.cells.get((row, col))
            assert actual is not None, f"Ячейка ({row},{col}) исчезла после сохранения"
            assert actual.value == expected_value, (
                f"Ячейка ({row},{col}): ожидалось {expected_value!r}, получено {actual.value!r}"
            )


# ===========================================================================
# Интеграционный тест: файлы вне дельты не изменяются после apply_bulk (задача 12.2)
# ===========================================================================

class TestIntegrationApplyBulkUnchanged:
    """
    Интеграционный тест: проверка, что файлы вне дельты не изменяются
    после apply_bulk, а файлы в дельте — изменяются корректно.

    Validates: Requirements 8.6
    """

    def test_files_outside_delta_unchanged_after_apply_bulk(self, tmp_path: Path) -> None:
        """
        1. Создать несколько .ods-файлов с известным содержимым.
        2. Создать DeltaStore с изменениями для конкретных ячеек.
        3. Запустить Processor.apply_bulk() на всех файлах.
        4. Перечитать все файлы.
        5. Убедиться, что ячейки НЕ в дельте не изменились во всех файлах.
        6. Убедиться, что ячейки В дельте изменились во всех файлах.
        """
        # ── Шаг 1: создать несколько .ods-файлов с известным содержимым ─────
        rows_data = [
            ["Продукт",  "Цена",  "Количество"],
            ["Яблоко",   "50",    "100"],
            ["Банан",    "30",    "200"],
            ["Вишня",    "80",    "50"],
        ]

        file_count = 4
        file_paths: list[Path] = []
        file_infos: list[FileInfo] = []
        for i in range(file_count):
            p = _make_ods_file(tmp_path, f"store_{i}.ods", rows_data)
            file_paths.append(p)
            file_infos.append(FileInfo(path=p, row_count=4))

        # ── Шаг 2: создать DeltaStore с изменениями для конкретных ячеек ────
        # Изменяем только ячейки (1, 1) и (2, 2) — цена яблока и количество банана
        delta = DeltaStore()
        delta.set(1, 1, CellDelta(
            value="55",
            font_name=None, font_size=None, bold=None, italic=None, color=None,
        ))
        delta.set(2, 2, CellDelta(
            value="250",
            font_name=None, font_size=None, bold=None, italic=None, color=None,
        ))

        # Запомним исходные значения всех ячеек для сравнения
        reader = OdsReader()
        originals: list[dict] = []
        for p in file_paths:
            sheet = reader.load(p)
            originals.append({k: v.value for k, v in sheet.cells.items()})

        # ── Шаг 3: запустить Processor.apply_bulk() на всех файлах ──────────
        log = LogStore()
        processor = Processor()
        progress_calls: list[tuple[int, int]] = []

        def track_progress(done: int, total: int) -> None:
            progress_calls.append((done, total))

        result = processor.apply_bulk(file_infos, delta, log, track_progress)

        # Убедиться, что все файлы обработаны успешно
        assert result.success_count == file_count, (
            f"Ожидалось {file_count} успешных, получено {result.success_count}; "
            f"ошибки: {result.errors}"
        )
        assert result.error_count == 0

        # ── Шаг 4–6: перечитать все файлы и проверить ────────────────────────
        delta_cells = {(1, 1): "55", (2, 2): "250"}

        for i, p in enumerate(file_paths):
            updated = reader.load(p)

            # Шаг 5: ячейки НЕ в дельте не изменились
            for (row, col), orig_value in originals[i].items():
                if (row, col) in delta_cells:
                    continue  # эти ячейки должны были измениться
                actual = updated.cells.get((row, col))
                assert actual is not None, (
                    f"Файл {p.name}: ячейка ({row},{col}) исчезла после apply_bulk"
                )
                assert actual.value == orig_value, (
                    f"Файл {p.name}: ячейка ({row},{col}) изменилась, "
                    f"хотя не входит в дельту: {orig_value!r} → {actual.value!r}"
                )

            # Шаг 6: ячейки В дельте изменились
            for (row, col), expected_new_value in delta_cells.items():
                actual = updated.cells.get((row, col))
                assert actual is not None, (
                    f"Файл {p.name}: ячейка ({row},{col}) исчезла после apply_bulk"
                )
                assert actual.value == expected_new_value, (
                    f"Файл {p.name}: ячейка ({row},{col}) должна была стать "
                    f"{expected_new_value!r}, но содержит {actual.value!r}"
                )
