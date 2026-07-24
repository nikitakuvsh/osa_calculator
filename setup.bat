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


echo [INFO] Downloading...

powershell -Command "Invoke-WebRequest -Uri '%REPO%' -OutFile '%ZIP%'"


if errorlevel 1 (
    echo.
    echo [ERROR] Can't downloading
    pause
    exit /b
)


echo [INFO] Rearchive...

powershell -Command "Expand-Archive -Path '%ZIP%' -DestinationPath '.' -Force"


if errorlevel 1 (
    echo.
    echo [ERROR] Error rearchive...
    pause
    exit /b
)


echo [INFO] Forwards files...


xcopy "osa_calculator-main\*" "." /E /H /Y > nul


if errorlevel 1 (
    echo.
    echo [ERROR] Error copy files
    pause
    exit /b
)


echo [INFO] Delete template files..


del "%ZIP%"
rmdir /s /q "osa_calculator-main"


echo.
echo ================================
echo       Installation succes!
echo ================================
echo.

pause