@echo off
rem Pega el token nuevo de Silucia en token_silucia.txt (se hace solo cuando el Excel avise que venció).
cd /d "%~dp0"
echo Copie el token de Silucia (con o sin la palabra Bearer), pegue aqui con clic derecho y presione Enter.
powershell -NoProfile -Command "$t = Read-Host 'Token'; if ($t.Trim().Length -gt 60) { Set-Content -Path 'token_silucia.txt' -Value $t.Trim() -Encoding ASCII; Remove-Item 'ALERTA_TOKEN.txt' -ErrorAction SilentlyContinue; Write-Host 'Listo: token guardado.' } else { Write-Host 'No se guardo: el texto pegado es muy corto.' }"
pause
