@echo off
REM Doppio clic per aprire il programma (Windows).
chcp 65001 >nul
cd /d "%~dp0"
set PY=
where py >nul 2>nul && set PY=py -3
if not defined PY where python >nul 2>nul && set PY=python
if not defined PY (
    echo Python non e' installato.
    echo Scaricalo da https://www.python.org/downloads/  e durante l'installazione
    echo spunta "Add python.exe to PATH". Poi fai di nuovo doppio clic su questo file.
    echo.
    pause
    exit /b 1
)
%PY% avvia.py %*
if errorlevel 1 pause
