"""
Pytest configuration: добавляет корень проекта в sys.path,
чтобы тесты могли импортировать ods_bulk_editor.
"""
import sys
from pathlib import Path

# Добавить корень проекта в sys.path
sys.path.insert(0, str(Path(__file__).parent))
