@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0app\instalar.ps1"
echo.
echo Se apareceu algum erro acima, tire um print e envie.
pause
