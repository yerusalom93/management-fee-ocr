param(
  [string]$OutZip = "",
  [switch]$Force
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

if (-not $OutZip) {
  $OutZip = Join-Path $ScriptDir "management-fee-ocr-offline.zip"
} elseif (-not [System.IO.Path]::IsPathRooted($OutZip)) {
  $OutZip = Join-Path (Get-Location) $OutZip
}

& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $ScriptDir "setup.ps1")
if ($LASTEXITCODE -ne 0) {
  throw "Setup failed. Offline bundle was not created."
}

$bundleRoot = Join-Path $ScriptDir "work\offline_bundle"
$bundleDir = Join-Path $bundleRoot "management-fee-ocr"

if (Test-Path -LiteralPath $bundleRoot) {
  $resolved = (Resolve-Path -LiteralPath $bundleRoot).Path
  $expectedPrefix = (Resolve-Path -LiteralPath (Join-Path $ScriptDir "work")).Path
  if (-not $resolved.StartsWith($expectedPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Refusing to clean unexpected bundle path: $resolved"
  }
  Remove-Item -LiteralPath $bundleRoot -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $bundleDir | Out-Null

$allowedExtensions = @(".ps1", ".bat", ".py", ".txt")
$allowedNames = @("README.md", ".gitignore")
Get-ChildItem -LiteralPath $ScriptDir -File | ForEach-Object {
  if ($allowedExtensions -contains $_.Extension.ToLowerInvariant() -or $allowedNames -contains $_.Name) {
    Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $bundleDir $_.Name)
  }
}

Copy-Item -LiteralPath (Join-Path $ScriptDir ".runtime") -Destination (Join-Path $bundleDir ".runtime") -Recurse
$runtimeDownloads = Join-Path $bundleDir ".runtime\downloads"
if (Test-Path -LiteralPath $runtimeDownloads) {
  Remove-Item -LiteralPath $runtimeDownloads -Recurse -Force
}

if (Test-Path -LiteralPath $OutZip) {
  if ($Force) {
    Remove-Item -LiteralPath $OutZip -Force
  } else {
    throw "Output ZIP already exists. Use -Force to overwrite: $OutZip"
  }
}

Compress-Archive -LiteralPath $bundleDir -DestinationPath $OutZip -Force
Write-Host "Offline bundle created: $OutZip"
