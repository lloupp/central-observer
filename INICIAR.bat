@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel% equ 0 (
  py -3 app.py
) else (
  where python >nul 2>nul
  if %errorlevel% equ 0 (
    python app.py
  ) else (
    echo Python 3 nao encontrado. Instale Python 3.10 ou superior.
    echo https://www.python.org/downloads/
    pause
    exit /b 1
  )
)
if errorlevel 1 pause
