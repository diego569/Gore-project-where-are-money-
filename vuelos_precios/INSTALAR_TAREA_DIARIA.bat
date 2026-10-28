@echo off
rem Crea dos tareas de Windows (todos los dias 08:00 y 20:00) que consultan los precios y guardan el historial.
cd /d "%~dp0"
schtasks /create /tn "VUELOS - precios 0800" /sc daily /st 08:00 /tr "cmd /c \"cd /d %~dp0 && set PYTHONIOENCODING=utf-8 && py -3 vuelos.py\"" /f
schtasks /create /tn "VUELOS - precios 2000" /sc daily /st 20:00 /tr "cmd /c \"cd /d %~dp0 && set PYTHONIOENCODING=utf-8 && py -3 vuelos.py\"" /f
if errorlevel 1 (echo No se pudo crear la tarea.) else (echo Listo. Consulta todos los dias a las 08:00 y 20:00 con la sesion de Windows abierta.)
pause
