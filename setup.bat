@echo off
chcp 65001 > nul
title OSA Calculator Setup

echo ================================
echo      OSA Calculator Setup
echo ================================
echo.

cd /d "%~dp0"


set REPO=https://github.com/nikitakuvsh/osa_calculator/archive/refs/heads/main.zip
set ZIP=osa_calculator.zip


echo [INFO] Скачиваю проект...

powershell -Command "Invoke-WebRequest -Uri '%REPO%' -OutFile '%ZIP%'"


if errorlevel 1 (
    echo.
    echo [ERROR] Не удалось скачать проект
    pause
    exit /b
)


echo [INFO] Распаковываю...


powershell -Command "Expand-Archive -Path '%ZIP%' -DestinationPath '.' -Force"


if errorlevel 1 (
    echo.
    echo [ERROR] Ошибка распаковки
    pause
    exit /b
)


echo [INFO] Перемещаю файлы...


xcopy "osa_calculator-main\*" "." /E /H /Y > nul


if errorlevel 1 (
    echo.
    echo [ERROR] Ошибка копирования файлов
    pause
    exit /b
)


echo [INFO] Удаляю временные файлы...


del "%ZIP%"
rmdir /s /q "osa_calculator-main"


echo.
echo ================================
echo       Установка завершена!
echo ================================
echo.

pause