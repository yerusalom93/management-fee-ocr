param(
  [Parameter(Mandatory = $true)][string]$Pdf,
  [string]$OutJson = ".\ocr_result_optimized.json",
  [string]$OutXlsx = ".\management_fee_optimized.xlsx",
  [string]$ImageDir = ".\pages_optimized",
  [string]$DetailImageDir = ".\pages_lower_table",
  [string]$DetailJson = ".\ocr_lower_table.json",
  [double]$Scale = 3.0,
  [int]$RenderWorkers = 4,
  [int]$OcrWorkers = 2,
  [switch]$ForceRender,
  [switch]$Preprocess,
  [string]$PythonExe = $null
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$renderScript = Join-Path $ScriptDir "render_pdf_pages.py"
$ocrScript = Join-Path $ScriptDir "run_windows_ocr.ps1"
$cropScript = Join-Path $ScriptDir "crop_ocr_regions.py"
$buildScript = Join-Path $ScriptDir "build_management_fee_xlsx.py"

if (-not $PythonExe) {
  $PythonExe = if ($env:OCR_PYTHON_EXE) { $env:OCR_PYTHON_EXE } else { "python" }
}

$pdfPath = (Resolve-Path -LiteralPath $Pdf).Path
$imagePath = if ([System.IO.Path]::IsPathRooted($ImageDir)) { $ImageDir } else { Join-Path (Get-Location) $ImageDir }
$outPath = if ([System.IO.Path]::IsPathRooted($OutJson)) { $OutJson } else { Join-Path (Get-Location) $OutJson }
New-Item -ItemType Directory -Force -Path $imagePath | Out-Null

$renderArgs = @(
  $renderScript,
  $pdfPath,
  $imagePath,
  "--scale",
  ([string]$Scale),
  "--workers",
  ([string]$RenderWorkers)
)
if ($ForceRender) { $renderArgs += "--force" }
if ($Preprocess) { $renderArgs += "--preprocess" }

$renderWatch = [System.Diagnostics.Stopwatch]::StartNew()
& $PythonExe @renderArgs
if ($LASTEXITCODE -ne 0) {
  throw "PDF rendering failed with exit code $LASTEXITCODE"
}
$renderWatch.Stop()

$images = Get-ChildItem -LiteralPath $imagePath -Filter "page*.png" | Sort-Object Name
if ($images.Count -eq 0) {
  throw "No rendered page images found in $imagePath"
}

$workerCount = [Math]::Max(1, [Math]::Min($OcrWorkers, $images.Count))
$chunkSize = [Math]::Ceiling($images.Count / $workerCount)
$chunkFiles = @()
$jobs = @()
$ocrWatch = [System.Diagnostics.Stopwatch]::StartNew()

for ($idx = 0; $idx -lt $workerCount; $idx++) {
  $start = ($idx * $chunkSize) + 1
  $end = [Math]::Min(($idx + 1) * $chunkSize, $images.Count)
  if ($start -gt $end) { continue }
  $chunkJson = Join-Path ([System.IO.Path]::GetDirectoryName($outPath)) ("ocr_chunk_{0:00}.json" -f ($idx + 1))
  $chunkFiles += $chunkJson
  $jobs += Start-Job -ScriptBlock {
    param($workDir, $ocrScript, $imagePath, $chunkJson, $pdfPath, $start, $end)
    Set-Location $workDir
    & powershell -NoProfile -ExecutionPolicy Bypass -File $ocrScript -ImageDir $imagePath -OutJson $chunkJson -Source $pdfPath -PageStart $start -PageEnd $end
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
  } -ArgumentList (Get-Location).Path, $ocrScript, $imagePath, $chunkJson, $pdfPath, $start, $end
}

Wait-Job -Job $jobs | Out-Null
foreach ($job in $jobs) {
  Receive-Job -Job $job
  if ($job.State -ne "Completed") {
    throw "OCR worker failed: $($job.State)"
  }
}
Remove-Job -Job $jobs
$ocrWatch.Stop()

$pages = @()
$language = $null
foreach ($chunkFile in $chunkFiles) {
  $chunk = Get-Content -LiteralPath $chunkFile -Raw -Encoding UTF8 | ConvertFrom-Json
  if (-not $language) { $language = $chunk.recognizerLanguage }
  foreach ($page in $chunk.pages) { $pages += $page }
}
$pages = $pages | Sort-Object page

$payload = [ordered]@{
  source = $pdfPath
  recognizerLanguage = $language
  createdAt = (Get-Date).ToString("s")
  render = [ordered]@{
    scale = $Scale
    preprocessed = [bool]$Preprocess
    seconds = [math]::Round($renderWatch.Elapsed.TotalSeconds, 2)
  }
  ocr = [ordered]@{
    workers = $workerCount
    seconds = [math]::Round($ocrWatch.Elapsed.TotalSeconds, 2)
  }
  pages = $pages
}

$payload | ConvertTo-Json -Depth 9 | Set-Content -LiteralPath $outPath -Encoding UTF8
Remove-Item -LiteralPath $chunkFiles -ErrorAction SilentlyContinue
Write-Host "Saved $outPath"
Write-Host ("Render: {0:n2}s, OCR: {1:n2}s, Pages: {2}" -f $renderWatch.Elapsed.TotalSeconds, $ocrWatch.Elapsed.TotalSeconds, $pages.Count)

if ($OutXlsx) {
  $detailImagePath = if ([System.IO.Path]::IsPathRooted($DetailImageDir)) { $DetailImageDir } else { Join-Path (Get-Location) $DetailImageDir }
  $detailJsonPath = if ([System.IO.Path]::IsPathRooted($DetailJson)) { $DetailJson } else { Join-Path (Get-Location) $DetailJson }
  & $PythonExe $cropScript $imagePath $detailImagePath "--region" "lower_table" "--scale" ([string]$Scale)
  if ($LASTEXITCODE -ne 0) {
    throw "Detail OCR crop failed with exit code $LASTEXITCODE"
  }
  & powershell -NoProfile -ExecutionPolicy Bypass -File $ocrScript -ImageDir $detailImagePath -OutJson $detailJsonPath -Source "$pdfPath#lower_table"
  if ($LASTEXITCODE -ne 0) {
    throw "Detail OCR failed with exit code $LASTEXITCODE"
  }

  $xlsxPath = if ([System.IO.Path]::IsPathRooted($OutXlsx)) { $OutXlsx } else { Join-Path (Get-Location) $OutXlsx }
  & $PythonExe $buildScript $outPath $xlsxPath $detailJsonPath
  if ($LASTEXITCODE -ne 0) {
    throw "XLSX export failed with exit code $LASTEXITCODE"
  }
}
