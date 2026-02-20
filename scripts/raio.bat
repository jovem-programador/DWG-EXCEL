@echo off
setlocal

REM Vai para a pasta onde está este .bat (scripts)
cd /d "%~dp0"

REM Ativa venv (ajuste se o venv estiver em outro lugar)
if exist "..\venv\Scripts\activate.bat" (
  call "..\venv\Scripts\activate.bat"
) else if exist "venv\Scripts\activate.bat" (
  call "venv\Scripts\activate.bat"
) else (
  echo [ERRO] Nao encontrei o venv. Crie um venv ou ajuste o caminho.
  pause
  exit /b 1
)
python -m streamlit run app.py

pause