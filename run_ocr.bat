@echo off
setlocal
if "%~1"=="" (
  powershell -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0ocr_drop_gui.ps1"
  exit /b %ERRORLEVEL%
)
set "OUT=%~2"
if "%OUT%"=="" set "OUT=%~dpn1_ocr.xlsx"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_ocr.ps1" -Pdf "%~1" -OutXlsx "%OUT%"
exit /b %ERRORLEVEL%
