@echo off
rem Doble clic: genera los dos Excel ahora (consulta las APIs y usa los archivos de "entrada" si los hay) y avisa con un mensaje.
cd /d "%~dp0"
actualizar.exe --diario --aviso
