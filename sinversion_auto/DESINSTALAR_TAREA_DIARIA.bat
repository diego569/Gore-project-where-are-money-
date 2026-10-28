@echo off
schtasks /delete /tn "SINVERSION - actualizar 0830" /f
schtasks /delete /tn "SINVERSION - actualizar 1430" /f
pause
