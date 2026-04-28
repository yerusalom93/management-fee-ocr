param(
  [Parameter(Mandatory = $true)][string]$Pdf,
  [string]$OutXlsx = "management_fee_optimized.xlsx",
  [switch]$ForceRender,
  [switch]$Preprocess
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvPython = Join-Path $ScriptDir ".venv\Scripts\python.exe"
$WorkDir = Join-Path $ScriptDir "work"
$ResolvedOutXlsx = if ([System.IO.Path]::IsPathRooted($OutXlsx)) { $OutXlsx } else { Join-Path (Get-Location) $OutXlsx }

if (-not (Test-Path $VenvPython)) {
  & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $ScriptDir "setup.ps1")
}

New-Item -ItemType Directory -Force -Path $WorkDir | Out-Null

$args = @(
  "-NoProfile",
  "-ExecutionPolicy",
  "BYPASS",
  "-File",
  (Join-Path $ScriptDir "process_management_fee_pdf.ps1"),
  "-Pdf",
  $Pdf,
  "-OutJson",
  (Join-Path $WorkDir "ocr_result.json"),
  "-OutXlsx",
  $ResolvedOutXlsx,
  "-ImageDir",
  (Join-Path $WorkDir "pages"),
  "-DetailImageDir",
  (Join-Path $WorkDir "pages_lower_table"),
  "-DetailJson",
  (Join-Path $WorkDir "ocr_lower_table.json"),
  "-PythonExe",
  $VenvPython
)

if ($ForceRender) { $args += "-ForceRender" }
if ($Preprocess) { $args += "-Preprocess" }

& powershell @args
if ($LASTEXITCODE -ne 0) {
  throw "OCR processing failed."
}
