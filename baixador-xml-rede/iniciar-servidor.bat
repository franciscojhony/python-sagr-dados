@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Baixador de XML de NF-e

rem Procura o Anaconda/Miniconda nos locais mais comuns
set "CONDA="
for %%D in (
  "%LOCALAPPDATA%\anaconda3"
  "%USERPROFILE%\anaconda3"
  "%USERPROFILE%\miniconda3"
  "%LOCALAPPDATA%\miniconda3"
  "C:\ProgramData\anaconda3"
  "C:\ProgramData\miniconda3"
  "C:\anaconda3"
) do (
  if not defined CONDA if exist "%%~D\Scripts\activate.bat" set "CONDA=%%~D"
)

if defined CONDA (
  echo Usando o Anaconda em: %CONDA%
  call "%CONDA%\Scripts\activate.bat" "%CONDA%"
  python app.py
  goto fim
)

rem Sem Anaconda: tenta o Python do Windows (py launcher) ou o python do PATH
where py >nul 2>nul
if not errorlevel 1 (
  py -3 app.py
  goto fim
)
python app.py

:fim
echo.
pause
