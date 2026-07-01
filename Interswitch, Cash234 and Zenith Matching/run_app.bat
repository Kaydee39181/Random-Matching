@echo off
setlocal

cd /d "%~dp0"
title Transaction Reconciliation App

echo.
echo ==========================================
echo  Transaction Reconciliation App
echo ==========================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo Python was not found on this computer.
    echo.
    echo Please install Python from https://www.python.org/downloads/
    echo During installation, tick "Add Python to PATH".
    echo.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo Setting up the app for first use...
    python -m venv .venv
    if errorlevel 1 (
        echo.
        echo Could not create the app environment.
        pause
        exit /b 1
    )
)

set APP_PYTHON=.venv\Scripts\python.exe

"%APP_PYTHON%" -c "import streamlit, pandas, openpyxl, xlsxwriter, xlrd, plotly" >nul 2>nul
if errorlevel 1 (
    python -c "import streamlit, pandas, openpyxl, xlsxwriter, xlrd, plotly" >nul 2>nul
    if not errorlevel 1 (
        echo Required packages were found in system Python.
        echo Using system Python for this app.
        set APP_PYTHON=python
        echo Requirements available through system Python. > ".venv\requirements-installed.txt"
    )
)

if "%APP_PYTHON%"==".venv\Scripts\python.exe" if not exist ".venv\requirements-installed.txt" (
    echo Installing required packages for first use...
    echo This needs internet the first time only.
    echo Using the bundled package installer already in the app environment.

    ".venv\Scripts\python.exe" -m pip install --default-timeout 120 --retries 10 -r requirements.txt
    if errorlevel 1 (
        echo.
        echo Could not install the required packages.
        echo Please check your internet connection and try again.
        pause
        exit /b 1
    )

    echo Requirements installed successfully. > ".venv\requirements-installed.txt"
    echo.
    echo Setup complete. Future starts will not install packages again.
    echo.
)

echo.
echo Starting the app...
echo If a browser does not open automatically, copy the local address shown below.
echo.
"%APP_PYTHON%" -m streamlit run app.py

pause
