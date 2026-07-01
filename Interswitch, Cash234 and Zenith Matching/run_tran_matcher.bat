@echo off
setlocal

cd /d "%~dp0"
title Tran ID vs Zenith Description Matcher

echo.
echo ==========================================
echo  Tran ID vs Zenith Description Matcher
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

set APP_PYTHON=python

python -c "import streamlit, pandas, openpyxl, xlsxwriter, xlrd, plotly" >nul 2>nul
if errorlevel 1 (
    echo Required packages were not found in system Python.
    echo Run run_app.bat first or install requirements.txt, then try again.
    echo.
    pause
    exit /b 1
)

echo Starting the matcher...
echo If a browser does not open automatically, copy the local address shown below.
echo.
"%APP_PYTHON%" -m streamlit run tran_description_matcher.py

pause
