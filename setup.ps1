param(
  [string]$PythonExe = "python",
  [switch]$Force
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvDir = Join-Path $ScriptDir ".venv"
$VenvPython = Join-Path $VenvDir "Scripts\python.exe"

if ($Force -and (Test-Path $VenvDir)) {
  Remove-Item -LiteralPath $VenvDir -Recurse -Force
}

if (-not (Test-Path $VenvPython)) {
  & $PythonExe -m venv $VenvDir
  if ($LASTEXITCODE -ne 0) {
    throw "Failed to create virtual environment. Install Python 3.11+ or pass -PythonExe."
  }
}

& $VenvPython -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) {
  throw "Failed to upgrade pip."
}

& $VenvPython -m pip install -r (Join-Path $ScriptDir "requirements.txt")
if ($LASTEXITCODE -ne 0) {
  throw "Failed to install Python dependencies."
}

Write-Host "Ready: $VenvPython"
