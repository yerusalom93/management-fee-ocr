param(
  [Parameter(Mandatory = $true)][string]$ImageDir,
  [Parameter(Mandatory = $true)][string]$OutJson,
  [string]$Source = $null,
  [int]$PageStart = 0,
  [int]$PageEnd = 0
)

Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Storage.FileAccessMode, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Globalization.Language, Windows.Globalization, ContentType = WindowsRuntime]

$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() |
  Where-Object {
    $_.Name -eq 'AsTask' -and
    $_.GetParameters().Count -eq 1 -and
    $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
  })[0]

function Await($op, [Type]$type) {
  $asTask = $asTaskGeneric.MakeGenericMethod($type)
  $task = $asTask.Invoke($null, @($op))
  $task.Wait() | Out-Null
  $task.Result
}

function RectObject($rect) {
  [ordered]@{
    x = [math]::Round($rect.X, 2)
    y = [math]::Round($rect.Y, 2)
    width = [math]::Round($rect.Width, 2)
    height = [math]::Round($rect.Height, 2)
  }
}

$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage([Windows.Globalization.Language]::new('ko'))
if ($null -eq $engine) {
  $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
}
if ($null -eq $engine) {
  throw "Windows OCR engine is unavailable."
}

$images = Get-ChildItem -LiteralPath $ImageDir -Filter 'page*.png' | Sort-Object Name
if ($PageStart -gt 0) {
  $images = $images | Where-Object {
    $pageNumber = [int]([regex]::Match($_.BaseName, '\d+').Value)
    $pageNumber -ge $PageStart -and ($PageEnd -le 0 -or $pageNumber -le $PageEnd)
  }
}

$pages = @()
foreach ($img in $images) {
  $pageNumber = [int]([regex]::Match($img.BaseName, '\d+').Value)
  Write-Host "OCR page $pageNumber"
  $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($img.FullName)) ([Windows.Storage.StorageFile])
  $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
  $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
  $bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
  $result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])

  $lines = @()
  foreach ($line in $result.Lines) {
    $words = @()
    foreach ($word in $line.Words) {
      $words += [ordered]@{
        text = $word.Text
        box = RectObject $word.BoundingRect
      }
    }
    $lines += [ordered]@{
      text = $line.Text
      words = $words
    }
  }
  $pages += [ordered]@{
    page = $pageNumber
    image = $img.FullName
    text = $result.Text
    lines = $lines
  }
}

$payload = [ordered]@{
  source = if ($Source) { $Source } else { $null }
  recognizerLanguage = $engine.RecognizerLanguage.LanguageTag
  createdAt = (Get-Date).ToString('s')
  pages = $pages
}

$payload | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $OutJson -Encoding UTF8
