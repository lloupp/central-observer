@echo off
setlocal EnableExtensions
cd /d "%~dp0"
set "VENV_PY=%~dp0.venv\Scripts\python.exe"
set "BOOTSTRAP="

echo ============================================
echo  Central Observer - Instalar certificados
echo ============================================
echo.

rem Use one Python consistently to create the local environment.
where py >nul 2>&1
if not errorlevel 1 (
  py -3 -c "import sys; sys.exit(sys.version_info < (3, 10))" >nul 2>&1
  if not errorlevel 1 set "BOOTSTRAP=py -3"
)
if not defined BOOTSTRAP (
  where python >nul 2>&1
  if not errorlevel 1 (
    python -c "import sys; sys.exit(sys.version_info < (3, 10))" >nul 2>&1
    if not errorlevel 1 set "BOOTSTRAP=python"
  )
)
if not defined BOOTSTRAP goto :no_python

if not exist "%VENV_PY%" (
  echo Criando ambiente isolado .venv com pip...
  %BOOTSTRAP% -m venv "%~dp0.venv"
  if errorlevel 1 goto :venv_error
)

rem Some Python installations omit pip. Try ensurepip in this environment.
"%VENV_PY%" -m pip --version >nul 2>&1
if errorlevel 1 (
  echo pip nao encontrado. Tentando recuperar com ensurepip...
  "%VENV_PY%" -m ensurepip --upgrade --default-pip
  if errorlevel 1 goto :pip_error
)

echo Instalando docxtpl e openpyxl neste ambiente local...
"%VENV_PY%" -m pip install --disable-pip-version-check -r requirements-certificados.txt
if errorlevel 1 goto :install_error

echo.
echo Bibliotecas instaladas com sucesso.
echo Abra INICIAR.bat. Ele usara este mesmo ambiente.
echo.
pause
exit /b 0

:no_python
echo ERRO: Python 3.10 ou superior nao foi encontrado.
echo Instale pelo site https://www.python.org/downloads/windows/
goto :error

:venv_error
echo ERRO: nao foi possivel criar o ambiente .venv.
echo No instalador do Python, selecione Modify e habilite pip e venv.
echo Depois exclua a pasta .venv incompleta e execute este arquivo novamente.
goto :error

:pip_error
echo ERRO: seu Python foi instalado sem o modulo ensurepip.
echo Abra o instalador do Python, selecione Modify e habilite pip.
echo Ou reinstale o Python oficial e execute este arquivo novamente.
goto :error

:install_error
echo ERRO: as bibliotecas nao puderam ser instaladas.
echo Confira a conexao e a mensagem de erro exibida acima.
echo Voce ainda pode usar a Central sem a emissao de certificados.
goto :error

:error
echo.
pause
exit /b 1
