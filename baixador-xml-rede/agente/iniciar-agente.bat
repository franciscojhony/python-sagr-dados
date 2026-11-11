@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"
title Agente Baixador de XML de NF-e

rem 1) Python embutido na pasta (pacote gerado pelo servidor)
if exist "python\python.exe" (
  "python\python.exe" agente.py
  goto fim
)

rem 2) Anaconda/Miniconda, se existir neste computador
set "CONDA="
for %%D in (
  "%LOCALAPPDATA%\anaconda3" "%USERPROFILE%\anaconda3" "%USERPROFILE%\miniconda3"
  "%LOCALAPPDATA%\miniconda3" "C:\ProgramData\anaconda3" "C:\ProgramData\miniconda3"
) do (
  if not defined CONDA if exist "%%~D\Scripts\activate.bat" set "CONDA=%%~D"
)
if defined CONDA (
  call "%CONDA%\Scripts\activate.bat" "%CONDA%"
  python agente.py
  goto fim
)

rem 3) Python do Windows
where py >nul 2>nul
if not errorlevel 1 (
  py -3 agente.py
  goto fim
)
where python >nul 2>nul
if not errorlevel 1 (
  python agente.py
  goto fim
)

echo.
echo Nao encontrei o Python neste computador e este pacote nao trouxe a pasta "python".
echo Peca ao responsavel pelo servidor o pacote completo do Agente.
echo.

:fim
echo.
pause
