@echo off
setlocal EnableExtensions
cd /d "%~dp0"
set "VENV_PY=%~dp0.venv\Scripts\python.exe"

rem The certificate installer creates a local Python environment.
rem Always use it here too, so packages are found in the same interpreter.
if exist "%VENV_PY%" goto :run_venv

where py >nul 2>&1
if not errorlevel 1 (
  py -3 -c "import sys; sys.exit(sys.version_info < (3, 10))" >nul 2>&1
  if not errorlevel 1 goto :run_py
)
where python >nul 2>&1
if not errorlevel 1 (
  python -c "import sys; sys.exit(sys.version_info < (3, 10))" >nul 2>&1
  if not errorlevel 1 goto :run_python
)
echo Python 3.10 ou superior nao encontrado.
echo Instale pelo site https://www.python.org/downloads/windows/
pause
exit /b 1

:run_venv
"%VENV_PY%" app.py
goto :done

:run_py
py -3 app.py
goto :done

:run_python
python app.py

:done
if errorlevel 1 (
  echo.
  echo A Central parou com erro. Confira a mensagem acima.
  pause
  exit /b 1
)
exit /b 0
