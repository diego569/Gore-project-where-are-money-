@echo off
rem Doble clic: genera los dos Excel usando SOLO las APIs (no necesita ningun archivo en "entrada").
rem Si alguna API falla, avisa que archivo hay que subir.
cd /d "%~dp0"
set SINV_FUENTE_SIAF=api
actualizar.exe --diario --aviso
