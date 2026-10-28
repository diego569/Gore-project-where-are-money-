@echo off
rem Crea una tarea de Windows que revisa la carpeta "entrada" cada 10 minutos y actualiza el Excel sola.
cd /d "%~dp0"
schtasks /create /tn "SINVERSION - actualizar proyectos" /sc minute /mo 10 /tr "\"%~dp0actualizar.exe\" --vigilar --aviso" /f
if errorlevel 1 (echo No se pudo crear la tarea. Pida ayuda a soporte.) else (echo Listo. La tarea revisa la carpeta "entrada" cada 10 minutos.)
pause
