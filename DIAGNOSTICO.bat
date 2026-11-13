@echo off
chcp 65001 >nul
cd /d "%~dp0app"
echo ================= DIAGNOSTICO =================
if not exist python_path.txt (
  echo [ERRO] python_path.txt nao existe: o INSTALAR nao terminou.
  goto fim
)
set /p PYW=<python_path.txt
set "PY=%PYW:pythonw.exe=python.exe%"
echo Python: "%PY%"
if not exist "%PY%" echo [ERRO] Esse arquivo do Python nao existe.
echo.
echo --- Componentes ---
"%PY%" -c "import sys;print(sys.version);import tkinter;print('tkinter ok');import openpyxl,pdfplumber,lxml,pypdfium2,PIL;print('bibliotecas ok')"
echo.
echo --- Programas ---
if exist "%ProgramFiles%\LibreOffice\program\soffice.exe" (echo LibreOffice ok) else (echo [ERRO] LibreOffice nao encontrado)
if exist "%ProgramFiles%\Tesseract-OCR\tesseract.exe" (echo Tesseract ok) else (echo [AVISO] Tesseract nao encontrado em Program Files)
echo.
echo --- Abrindo o programa (se der erro, ele aparece abaixo) ---
"%PY%" "%~dp0app\gerador_app.pyw"
:fim
echo.
echo Tire um print desta janela e envie.
pause
