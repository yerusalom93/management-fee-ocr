param(
  [string]$PythonExe = "",
  [switch]$UseSystemPython,
  [switch]$Force
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RuntimeDir = Join-Path $ScriptDir ".runtime"
$RuntimePythonDir = Join-Path $RuntimeDir "python"
$RuntimePython = Join-Path $RuntimePythonDir "python.exe"
$DownloadDir = Join-Path $RuntimeDir "downloads"
$VenvDir = Join-Path $ScriptDir ".venv"
$VenvPython = Join-Path $VenvDir "Scripts\python.exe"
$Requirements = Join-Path $ScriptDir "requirements.txt"

$EmbeddedPythonVersion = "3.12.10"
$EmbeddedPythonZip = "python-$EmbeddedPythonVersion-embed-amd64.zip"
$EmbeddedPythonUrl = "https://www.python.org/ftp/python/$EmbeddedPythonVersion/$EmbeddedPythonZip"
$GetPipUrl = "https://bootstrap.pypa.io/get-pip.py"

function Test-PythonPath([string]$Path) {
  if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return $false }
  try {
    & $Path -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 11) else 1)" 2>$null
    return $LASTEXITCODE -eq 0
  } catch {
    return $false
  }
}

function Test-OcrPackages([string]$Path) {
  if (-not (Test-PythonPath $Path)) { return $false }
  try {
    & $Path -c "import fitz, PIL, openpyxl" 2>$null
    return $LASTEXITCODE -eq 0
  } catch {
    return $false
  }
}

function Invoke-Download([string]$Url, [string]$OutFile) {
  New-Item -ItemType Directory -Force -Path (Split-Path -Parent $OutFile) | Out-Null
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
  Write-Host "Downloading $Url"
  Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing
}

function Enable-EmbeddedPythonSite([string]$PythonDir) {
  $pth = Get-ChildItem -LiteralPath $PythonDir -Filter "python*._pth" | Select-Object -First 1
  if (-not $pth) { return }

  $lines = Get-Content -LiteralPath $pth.FullName -Encoding UTF8
  $updated = New-Object System.Collections.Generic.List[string]
  $hasSitePackages = $false
  $hasImportSite = $false

  foreach ($line in $lines) {
    if ($line.Trim() -eq "Lib\site-packages") { $hasSitePackages = $true }
    if ($line.Trim() -eq "import site") {
      $hasImportSite = $true
      $updated.Add($line)
    } elseif ($line.Trim() -eq "#import site") {
      $hasImportSite = $true
      $updated.Add("import site")
    } else {
      $updated.Add($line)
    }
  }

  if (-not $hasSitePackages) {
    $updated.Insert([Math]::Max(0, $updated.Count - 1), "Lib\site-packages")
  }
  if (-not $hasImportSite) {
    $updated.Add("import site")
  }

  Set-Content -LiteralPath $pth.FullName -Value $updated -Encoding ASCII
}

function Install-EmbeddedPython {
  if ($Force -and (Test-Path -LiteralPath $RuntimePythonDir)) {
    Remove-Item -LiteralPath $RuntimePythonDir -Recurse -Force
  }

  if (-not (Test-Path -LiteralPath $RuntimePython)) {
    New-Item -ItemType Directory -Force -Path $RuntimeDir | Out-Null
    $zipPath = Join-Path $DownloadDir $EmbeddedPythonZip
    if (-not (Test-Path -LiteralPath $zipPath)) {
      Invoke-Download $EmbeddedPythonUrl $zipPath
    }

    if (Test-Path -LiteralPath $RuntimePythonDir) {
      Remove-Item -LiteralPath $RuntimePythonDir -Recurse -Force
    }
    New-Item -ItemType Directory -Force -Path $RuntimePythonDir | Out-Null
    Expand-Archive -LiteralPath $zipPath -DestinationPath $RuntimePythonDir -Force
    Enable-EmbeddedPythonSite $RuntimePythonDir
  }

  if (-not (Test-PythonPath $RuntimePython)) {
    throw "Portable Python could not be started: $RuntimePython"
  }

  return $RuntimePython
}

function Resolve-SystemPython([string]$Requested) {
  if ($Requested -and (Test-PythonPath $Requested)) {
    return (Resolve-Path -LiteralPath $Requested).Path
  }

  $py = Get-Command py -ErrorAction SilentlyContinue
  if ($py) {
    try {
      $path = & $py.Source -3 -c "import sys; print(sys.executable)" 2>$null
      if ($LASTEXITCODE -eq 0 -and (Test-PythonPath $path.Trim())) { return $path.Trim() }
    } catch {}
  }

  $python = Get-Command python -ErrorAction SilentlyContinue
  if ($python -and $python.Source -notlike "*\WindowsApps\python.exe") {
    try {
      $path = & $python.Source -c "import sys; print(sys.executable)" 2>$null
      if ($LASTEXITCODE -eq 0 -and (Test-PythonPath $path.Trim())) { return $path.Trim() }
    } catch {}
  }

  $known = @(
    (Join-Path $env:LocalAppData "Programs\Python\Python312\python.exe"),
    (Join-Path $env:LocalAppData "Programs\Python\Python311\python.exe"),
    "C:\Program Files\Python312\python.exe",
    "C:\Program Files\Python311\python.exe"
  )
  foreach ($candidate in $known) {
    if (Test-PythonPath $candidate) { return $candidate }
  }

  throw "Python 3.11+ was not found. Run without -UseSystemPython to use the portable runtime."
}

function Install-VenvFromSystemPython([string]$SystemPython) {
  if ($Force -and (Test-Path -LiteralPath $VenvDir)) {
    Remove-Item -LiteralPath $VenvDir -Recurse -Force
  }
  if (-not (Test-Path -LiteralPath $VenvPython)) {
    Write-Host "Using system Python: $SystemPython"
    & $SystemPython -m venv $VenvDir
    if ($LASTEXITCODE -ne 0) {
      throw "Failed to create virtual environment."
    }
  }
  return $VenvPython
}

function Install-PythonPackages([string]$PythonPath) {
  $pipOk = $false
  try {
    & $PythonPath -m pip --version 2>$null | Out-Null
    $pipOk = $LASTEXITCODE -eq 0
  } catch {}

  if (-not $pipOk) {
    $getPipPath = Join-Path $DownloadDir "get-pip.py"
    if (-not (Test-Path -LiteralPath $getPipPath)) {
      Invoke-Download $GetPipUrl $getPipPath
    }
    & $PythonPath $getPipPath --no-warn-script-location
    if ($LASTEXITCODE -ne 0) {
      throw "Failed to bootstrap pip."
    }
  }

  & $PythonPath -m pip install --upgrade pip --no-warn-script-location
  if ($LASTEXITCODE -ne 0) {
    throw "Failed to upgrade pip."
  }

  & $PythonPath -m pip install -r $Requirements --no-warn-script-location
  if ($LASTEXITCODE -ne 0) {
    throw "Failed to install Python dependencies."
  }

  if (-not (Test-OcrPackages $PythonPath)) {
    throw "Python dependencies were installed, but OCR imports still failed."
  }
}

if ($Force -and -not $UseSystemPython -and (Test-Path -LiteralPath $RuntimeDir)) {
  Remove-Item -LiteralPath $RuntimeDir -Recurse -Force
}

if ($UseSystemPython -or $PythonExe) {
  $SystemPython = Resolve-SystemPython $PythonExe
  $SelectedPython = Install-VenvFromSystemPython $SystemPython
} else {
  $SelectedPython = Install-EmbeddedPython
}

Install-PythonPackages $SelectedPython
Write-Host "Ready: $SelectedPython"
