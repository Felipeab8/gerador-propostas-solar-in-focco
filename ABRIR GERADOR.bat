@echo off
chcp 65001 >nul
cd /d "%~dp0app"
if not exist python_path.txt (
  echo O programa ainda nao foi instalado. Clique duas vezes em INSTALAR.bat primeiro.
  pause
  exit /b
)
set /p PYW=<python_path.txt
start "" "%PYW%" "%~dp0app\gerador_app.pyw"
