@echo off
chcp 65001 >nul
echo ============================================
echo  Маршруты, Сборка — Установка библиотек
echo ============================================
echo.

:: Проверяем наличие Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ОШИБКА] Python не найден.
    echo Скачайте Python 3.10 или новее с https://www.python.org/downloads/
    echo При установке отметьте "Add Python to PATH"
    pause
    exit /b 1
)

echo [OK] Python найден:
python --version
echo.

:: Обновляем pip
echo [1/3] Обновление pip...
python -m pip install --upgrade pip --quiet
echo [OK] pip обновлён
echo.

:: Устанавливаем зависимости
echo [2/3] Установка библиотек из requirements.txt...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [ОШИБКА] Не удалось установить библиотеки.
    echo Проверьте подключение к интернету и повторите попытку.
    pause
    exit /b 1
)
echo.
echo [OK] Все библиотеки установлены
echo.

:: Проверяем установку
echo [3/3] Проверка установки...
python -c "import xlrd, xlwt, PyQt6; print('[OK] xlrd:', xlrd.__version__); print('[OK] xlwt:', xlwt.__VERSION__); from PyQt6.QtCore import QT_VERSION_STR; print('[OK] PyQt6 / Qt:', QT_VERSION_STR)"
if errorlevel 1 (
    echo [ОШИБКА] Проверка не прошла. Попробуйте запустить скрипт ещё раз.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  Установка завершена успешно!
echo  Запустите приложение: python app.py
echo ============================================
pause
