param(
  [Parameter(Mandatory = $true)][string]$Pdf,
  [string]$OutXlsx = "management_fee_optimized.xlsx",
  [switch]$ForceRender,
  [switch]$Preprocess
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvPython = Join-Path $ScriptDir ".venv\Scripts\python.exe"
$RuntimePython = Join-Path $ScriptDir ".runtime\python\python.exe"
$WorkDir = Join-Path $ScriptDir "work"
$ResolvedOutXlsx = if ([System.IO.Path]::IsPathRooted($OutXlsx)) { $OutXlsx } else { Join-Path (Get-Location) $OutXlsx }

function Test-OcrPython([string]$Path) {
  if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return $false }
  try {
    & $Path -c "import fitz, PIL, openpyxl" 2>$null
    return $LASTEXITCODE -eq 0
  } catch {
    return $false
  }
}

function Get-ProjectPython {
  if (Test-OcrPython $RuntimePython) { return $RuntimePython }
  if (Test-OcrPython $VenvPython) { return $VenvPython }
  return $null
}

$PythonExe = Get-ProjectPython
if (-not $PythonExe) {
  & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $ScriptDir "setup.ps1")
  if ($LASTEXITCODE -ne 0) {
    throw "Setup failed."
  }
  $PythonExe = Get-ProjectPython
}

if (-not $PythonExe) {
  throw "Python runtime is not ready. Run setup.ps1 and check the error output."
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
  $PythonExe
)

if ($ForceRender) { $args += "-ForceRender" }
if ($Preprocess) { $args += "-Preprocess" }

& powershell @args
if ($LASTEXITCODE -ne 0) {
  throw "OCR processing failed."
}
