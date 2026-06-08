@echo off
title Rate Sheet Generator

:: Check Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo  Python is not installed.
    echo  Download it from: https://www.python.org/downloads/
    echo  IMPORTANT: tick "Add Python to PATH" during install.
    echo.
    pause
    exit /b
)

:: Create virtual environment on first run
if not exist venv\ (
    echo.
    echo  First-time setup - takes about 60 seconds...
    echo.
    python -m venv venv
)

:: Activate venv and install dependencies
call venv\Scripts\activate.bat
pip install flask pandas openpyxl pywebview pillow --quiet --disable-pip-version-check

:: Launch
echo.
echo  Starting Rate Sheet Generator...
echo.
pythonw launcher.py
