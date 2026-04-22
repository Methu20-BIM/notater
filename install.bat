@echo off
chcp 65001 > nul
title Notater - Installasjon
color 0A

echo.
echo  ================================================
echo   NOTATER - Matteassistent for Word (R1/R2)
echo  ================================================
echo.

set "NOTATER_DIR=%~dp0"
set "VENV_DIR=%NOTATER_DIR%venv"
set "PYTHON_EXE=%VENV_DIR%\Scripts\python.exe"
set "PYTHONW_EXE=%VENV_DIR%\Scripts\pythonw.exe"
set "PIP_EXE=%VENV_DIR%\Scripts\pip.exe"

echo [1/5] Sjekker Python...
python --version > nul 2>&1
if errorlevel 1 (
    echo.
    echo  FEIL: Python ikke funnet.
    echo  Last ned fra: https://www.python.org/downloads/
    echo  Husk aa huke av "Add Python to PATH" under installasjon.
    pause
    exit /b 1
)
for /f "tokens=*" %%i in ('python --version') do echo  Funnet: %%i
echo.

echo [2/5] Oppretter virtuelt Python-miljoe...
if not exist "%VENV_DIR%" (
    python -m venv "%VENV_DIR%"
    echo  Virtuelt miljoe opprettet.
) else (
    echo  Virtuelt miljoe finnes allerede.
)
echo.

echo [3/5] Installerer Python-pakker...
"%PIP_EXE%" install --quiet --upgrade pip
"%PIP_EXE%" install --quiet -r "%NOTATER_DIR%requirements.txt"
if errorlevel 1 (
    "%PIP_EXE%" install flask flask-cors python-docx requests Pillow pystray pywin32
)
echo  Python-pakker installert.
echo.

echo [4/5] Sjekker Ollama...
ollama --version > nul 2>&1
if errorlevel 1 (
    echo  Ollama ikke funnet. Aapner nedlastingsside...
    start "" "https://ollama.com/download/OllamaSetup.exe"
    echo  Installer Ollama og kjoer install.bat paa nytt.
    pause
    exit /b 1
)
echo  Ollama funnet.
start /b ollama serve > nul 2>&1
timeout /t 3 /nobreak > nul
echo  Laster ned deepseek-r1:7b (ca. 4 GB)...
ollama pull deepseek-r1:7b
echo.

echo [5/5] Setter opp autostart og snarvei...
reg add "HKCU\Software\Microsoft\Office\16.0\Word\Security" /v AccessVBOM /t REG_DWORD /d 1 /f > nul 2>&1
reg add "HKCU\Software\Microsoft\Office\15.0\Word\Security" /v AccessVBOM /t REG_DWORD /d 1 /f > nul 2>&1
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v Notater /t REG_SZ /d "\"%PYTHONW_EXE%\" \"%NOTATER_DIR%backend\main.py\"" /f > nul 2>&1

set "DESKTOP=%USERPROFILE%\OneDrive - Osloskolen\Skrivebord"
if not exist "%DESKTOP%" set "DESKTOP=%USERPROFILE%\Desktop"

powershell -Command "$ws=New-Object -ComObject WScript.Shell;$lnk=$ws.CreateShortcut('%DESKTOP%\matte.lnk');$lnk.TargetPath='%NOTATER_DIR%start.bat';$lnk.WorkingDirectory='%NOTATER_DIR%';$lnk.IconLocation='C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE,0';$lnk.Save()"

echo  Ferdig!
echo.
echo  ================================================
echo   Installasjon fullfort!
echo  ================================================
echo.
echo   Dobbeltklikk "matte" paa skrivebordet for aa starte.
echo   Skriv oppgave + " - los" og trykk "Los oppgaver".
echo.
pause
