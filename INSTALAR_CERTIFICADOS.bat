@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel% equ 0 (
  py -3 -m pip install -r requirements-certificados.txt
) else (
  python -m pip install -r requirements-certificados.txt
)
if errorlevel 1 (
  echo Falha ao instalar as bibliotecas. Verifique sua conexao e o Python.
) else (
  echo Bibliotecas instaladas. Abra INICIAR.bat para usar a Central.
)
pause
