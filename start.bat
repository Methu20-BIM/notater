@echo off
chcp 65001 > nul
title Notater

set "NOTATER_DIR=%~dp0"
set "PYTHONW_EXE=%NOTATER_DIR%venv\Scripts\pythonw.exe"

if not exist "%PYTHONW_EXE%" (
    echo Notater er ikke installert.
    echo Kjoer install.bat forst.
    pause
    exit /b 1
)

taskkill /f /im pythonw.exe > nul 2>&1
timeout /t 1 /nobreak > nul

start "" "%PYTHONW_EXE%" "%NOTATER_DIR%backend\main.py"
