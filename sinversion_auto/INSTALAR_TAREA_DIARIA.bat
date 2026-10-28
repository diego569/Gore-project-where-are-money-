@echo off
rem Crea dos tareas de Windows (lunes a viernes 08:30 y 14:30) que actualizan los dos Excel usando las APIs.
cd /d "%~dp0"
schtasks /create /tn "SINVERSION - actualizar 0830" /sc weekly /d MON,TUE,WED,THU,FRI /st 08:30 /tr "\"%~dp0actualizar.exe\" --diario --aviso" /f
schtasks /create /tn "SINVERSION - actualizar 1430" /sc weekly /d MON,TUE,WED,THU,FRI /st 14:30 /tr "\"%~dp0actualizar.exe\" --diario --aviso" /f
if errorlevel 1 (echo No se pudo crear la tarea. Pida ayuda a soporte.) else (echo Listo. Se actualiza de lunes a viernes a las 08:30 y 14:30 con la sesion de Windows abierta.)
pause
