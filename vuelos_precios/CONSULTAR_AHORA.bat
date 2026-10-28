@echo off
rem Doble clic: consulta Google Flights ahora, agrega la foto de precios al historial y actualiza vuelos_precios.xlsx
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
py -3 vuelos.py
echo.
pause
