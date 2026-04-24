@echo off
chcp 65001 >nul
echo === Генератор рассадки офиса ===
echo.

:: Проверяем Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ОШИБКА] Python не найден.
    echo Установите Python 3.9+ с https://python.org
    echo При установке отметьте "Add Python to PATH"
    pause
    exit /b 1
)

echo Установка зависимостей...
python -m pip install openpyxl pyyaml --quiet

echo Запуск приложения...
python "%~dp0ui.py"

if errorlevel 1 (
    echo.
    echo [ОШИБКА] Приложение завершилось с ошибкой.
    pause
)
