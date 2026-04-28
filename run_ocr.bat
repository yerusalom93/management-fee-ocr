@echo off
setlocal
if "%~1"=="" (
  echo Usage: run_ocr.bat "path\to\bill.pdf" [output.xlsx]
  exit /b 2
)
set "OUT=%~2"
if "%OUT%"=="" set "OUT=management_fee_optimized.xlsx"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_ocr.ps1" -Pdf "%~1" -OutXlsx "%OUT%"
exit /b %ERRORLEVEL%
