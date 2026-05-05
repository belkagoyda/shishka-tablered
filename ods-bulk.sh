#!/bin/bash
# ODS Bulk Editor — launcher script
# Редактирование этого файла КАТЕГОРИЧЕСКИ ЗАПРЕЩЕНО!
cd "$(dirname "$0")"

echo "============================"
echo "   ODS Bulk Editor"
echo "============================"
echo ""
echo "1) Установить зависимости"
echo "2) Запустить программу"
echo ""
read -p "Выберите действие [1/2]: " choice

case "$choice" in
    1)
        echo "Установка зависимостей..."
        sudo apt install -y python3 python3-tk python3-odf
        echo ""
        read -p "Запустить программу? [y/N]: " run_now
        if [ "$run_now" = "y" ] || [ "$run_now" = "Y" ]; then
            exec python3 ods_bulk_editor.py "$@"
        fi
        ;;
    2)
        exec python3 ods_bulk_editor.py "$@"
        ;;
    *)
        echo "Неверный выбор. Запуск программы..."
        exec python3 ods_bulk_editor.py "$@"
        ;;
esac
