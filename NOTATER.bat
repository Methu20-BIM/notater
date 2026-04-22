@echo off
chcp 65001 > nul
title Notater

set "DIR=%~dp0"
set "VENV=%DIR%venv\Scripts\pythonw.exe"

if not exist "%VENV%" (
    echo  Foerste gang - installerer alt automatisk...
    call "%DIR%install.bat"
)

echo  Starter Notater og aapner matte.docx...
taskkill /f /im pythonw.exe > nul 2>&1
timeout /t 1 /nobreak > nul
start "" "%VENV%" "%DIR%backend\main.py"
