@echo off
chcp 65001 > nul
title OSA Calculator

echo ================================
echo        OSA Calculator
echo ================================
echo.

REM Проверяем наличие venv
if not exist "venv\Scripts\activate.bat" (
    echo [INFO] Создаю виртуальное окружение...
    python -m venv venv

    if errorlevel 1 (
        echo [ERROR] Не удалось создать venv
        pause
        exit /b
    )
)

echo [INFO] Активирую окружение...
call venv\Scripts\activate.bat


REM Устанавливаем зависимости
if exist "requirements.txt" (
    echo [INFO] Проверяю зависимости...
    pip install -r requirements.txt

    if errorlevel 1 (
        echo.
        echo [ERROR] Ошибка установки зависимостей
        pause
        exit /b
    )
) else (
    echo [WARNING] requirements.txt не найден
)


echo.
echo ================================
echo       Запуск расчёта...
echo ================================
echo.


REM Запуск программы
python main.py


echo.
echo ================================
echo       Работа завершена
echo ================================
echo.

pause