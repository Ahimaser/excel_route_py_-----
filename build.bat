@echo off
chcp 65001 >nul
echo === Build: Маршруты, Сборка ===

python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python not found.
    pause
    exit /b 1
)

echo Installing PyInstaller...
pip install pyinstaller --quiet

echo Installing dependencies...
pip install -r requirements.txt --quiet

echo Running PyInstaller...
pyinstaller --noconfirm Маршруты_Сборка.spec

if errorlevel 1 (
    echo Build failed.
    pause
    exit /b 1
)

echo.
echo Done: dist\Маршруты_Сборка.exe
pause
