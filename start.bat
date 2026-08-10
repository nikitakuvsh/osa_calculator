@echo off
title OSA Calculator

echo ================================
echo        OSA Calculator
echo ================================
echo.


if not exist "venv\Scripts\activate.bat" (
    echo [INFO] Creating virtual environment...

    python -m venv venv

    if errorlevel 1 (
        echo [ERROR] Failed to create venv
        pause
        exit /b
    )
)


echo [INFO] Activating environment...

call venv\Scripts\activate.bat


if exist "requirements.txt" (

    echo [INFO] Installing dependencies...

    python -m pip install --upgrade pip
    pip install -r requirements.txt

    if errorlevel 1 (
        echo [ERROR] Failed to install dependencies
        pause
        exit /b
    )

) else (

    echo [WARNING] requirements.txt not found

)


echo.
echo ================================
echo       Starting calculator
echo ================================
echo.

cd src
python main.py


echo.
echo ================================
echo       Finished
echo ================================
echo.

pause