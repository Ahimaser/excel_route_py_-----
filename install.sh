#!/usr/bin/env bash
# Excel Route Manager — Установка библиотек (Linux / macOS)
# Использование: bash install.sh

set -e

echo "============================================"
echo " Excel Route Manager — Установка библиотек"
echo "============================================"
echo ""

# Определяем команду python
PYTHON=""
for cmd in python3 python3.11 python3.10 python; do
    if command -v "$cmd" &>/dev/null; then
        VER=$("$cmd" -c "import sys; print(sys.version_info >= (3,10))" 2>/dev/null)
        if [ "$VER" = "True" ]; then
            PYTHON="$cmd"
            break
        fi
    fi
done

if [ -z "$PYTHON" ]; then
    echo "[ОШИБКА] Python 3.10+ не найден."
    echo "Установите Python: https://www.python.org/downloads/"
    exit 1
fi

echo "[OK] Python найден: $($PYTHON --version)"
echo ""

# Обновляем pip
echo "[1/3] Обновление pip..."
"$PYTHON" -m pip install --upgrade pip --quiet
echo "[OK] pip обновлён"
echo ""

# Устанавливаем зависимости
echo "[2/3] Установка библиотек из requirements.txt..."
"$PYTHON" -m pip install -r requirements.txt
echo ""
echo "[OK] Все библиотеки установлены"
echo ""

# Проверяем установку
echo "[3/3] Проверка установки..."
"$PYTHON" -c "
import xlrd, xlwt, PyQt6
from PyQt6.QtCore import QT_VERSION_STR
print('[OK] xlrd:', xlrd.__version__)
print('[OK] xlwt:', xlwt.__VERSION__)
print('[OK] PyQt6 / Qt:', QT_VERSION_STR)
"

echo ""
echo "============================================"
echo " Установка завершена успешно!"
echo " Запустите приложение: $PYTHON app.py"
echo "============================================"
