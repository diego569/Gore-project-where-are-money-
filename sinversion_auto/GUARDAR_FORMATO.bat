@echo off
cd /d "%~dp0"
if "%~1"=="" (
  echo Arrastre el Excel que usted retoco (orden de hojas, hojas ocultas, anchos) sobre este archivo.
  pause
  exit /b
)
actualizar.exe --guardar-formato "%~1"
