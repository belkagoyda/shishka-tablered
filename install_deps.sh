#!/bin/bash
# Скрипт установки зависимостей для ODS Bulk Editor.
# Устанавливает необходимые пакеты: Python 3, tkinter и библиотеку для работы с ODS-файлами (odfpy).
# Запускать с правами суперпользователя: sudo bash install_deps.sh

apt install -y python3 python3-tk python3-odf
