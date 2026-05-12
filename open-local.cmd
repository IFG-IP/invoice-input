@echo off
setlocal
set PORT=5173
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$processIds = Get-NetTCPConnection -LocalPort %PORT% -State Listen -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess -Unique; foreach ($processId in $processIds) { Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue }"
start "invoice-input-server" powershell -NoExit -NoProfile -ExecutionPolicy Bypass -File "%~dp0serve.ps1" -Port %PORT%
timeout /t 1 /nobreak > nul
start "" "http://127.0.0.1:%PORT%/"
