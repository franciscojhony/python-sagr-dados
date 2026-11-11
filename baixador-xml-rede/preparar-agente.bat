@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Preparando o Python portatil do Agente (precisa de internet)...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0preparar-agente.ps1"
echo.
pause
