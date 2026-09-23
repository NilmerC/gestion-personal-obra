@echo off
title Gestion de Personal de Obra
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Creando entorno virtual...
  py -m venv .venv
)
call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
pip install -r requirements.txt
echo.
echo Iniciando Plataforma de Gestion de Personal de Obra...
echo.
streamlit run app.py
pause
