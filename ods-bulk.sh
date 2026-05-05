#!/bin/bash

REPO_URL="https://github.com/belkagoyda/shishka-tablered"
TEMP_DIR=".ods_bulk_editor_temp"

show_menu() {
    echo ""
    echo "ODS Bulk Editor"
    echo "---------------"
    echo "0 - Установить (зависимости + программа)"
    echo "1 - Запустить редактор"
    echo "2 - Посмотреть лог (ods_bulk_editor.log)"
    echo -n "Выберите действие: "
}

install() {
    echo ">>> Установка системных зависимостей (требуется sudo)..."
    sudo apt install -y python3 python3-tk python3-odf git

    echo ">>> Клонирование репозитория..."
    git clone "$REPO_URL" "$TEMP_DIR"

    if [ -f "$TEMP_DIR/ods_bulk_editor.py" ]; then
        echo ">>> Копирование ods_bulk_editor.py в текущую папку..."
        cp "$TEMP_DIR/ods_bulk_editor.py" .
        echo ">>> Очистка временных файлов..."
        rm -rf "$TEMP_DIR"
        echo ">>> Готово! Теперь можно запускать (пункт 1)."
    else
        echo ">>> Ошибка: не удалось найти ods_bulk_editor.py в репозитории."
        rm -rf "$TEMP_DIR"
        exit 1
    fi
}

run() {
    if [ ! -f "ods_bulk_editor.py" ]; then
        echo ">>> Программа не установлена. Сначала выберите пункт 0."
        return
    fi
    echo ">>> Запуск ODS Bulk Editor..."
    python3 ods_bulk_editor.py
}

view_log() {
    if [ -f "ods_bulk_editor.log" ]; then
        less ods_bulk_editor.log
    else
        echo ">>> Лог-файл не найден."
    fi
}

# Бесконечный цикл, выход через Ctrl+C
while true; do
    show_menu
    read choice
    case $choice in
        0) install ;;
        1) run ;;
        2) view_log ;;
        *) echo "Неверный ввод, попробуйте снова." ;;
    esac
done
