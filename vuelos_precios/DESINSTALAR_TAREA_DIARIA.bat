@echo off
schtasks /delete /tn "VUELOS - precios 0800" /f
schtasks /delete /tn "VUELOS - precios 2000" /f
pause
