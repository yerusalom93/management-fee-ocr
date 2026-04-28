param(
  [switch]$SelfTest
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RunScript = Join-Path $ScriptDir "run_ocr.ps1"
$WorkDir = Join-Path $ScriptDir "work"
$LogFile = Join-Path $WorkDir "last_gui.log"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if ($SelfTest) {
  Write-Host "GUI dependencies loaded."
  return
}

function Quote-Arg([string]$Value) {
  '"' + ($Value -replace '"', '\"') + '"'
}

function Add-Log([string]$Text) {
  if ([string]::IsNullOrWhiteSpace($Text)) {
    return
  }
  try {
    New-Item -ItemType Directory -Force -Path $WorkDir | Out-Null
    Add-Content -LiteralPath $LogFile -Value $Text -Encoding UTF8
  } catch {}
  $action = [System.Action]{
    $script:LogBox.AppendText($Text + [Environment]::NewLine)
    $script:LogBox.SelectionStart = $script:LogBox.Text.Length
    $script:LogBox.ScrollToCaret()
  }
  if ($script:Form.InvokeRequired) {
    [void]$script:Form.BeginInvoke($action)
  } else {
    $action.Invoke()
  }
}

function Set-Busy([bool]$Busy) {
  $script:SelectButton.Enabled = -not $Busy
  $script:DropPanel.Enabled = -not $Busy
  $script:Progress.Visible = $Busy
  if ($Busy) {
    $script:Progress.Style = "Marquee"
    $script:StatusLabel.Text = "Processing..."
  } else {
    $script:Progress.Style = "Blocks"
    $script:Progress.Value = 0
  }
}

function Start-Ocr([string]$PdfPath) {
  if (-not (Test-Path -LiteralPath $PdfPath)) {
    [System.Windows.Forms.MessageBox]::Show("File not found: $PdfPath", "OCR", "OK", "Error") | Out-Null
    return
  }
  if ([System.IO.Path]::GetExtension($PdfPath).ToLowerInvariant() -ne ".pdf") {
    [System.Windows.Forms.MessageBox]::Show("Please drop a PDF file.", "OCR", "OK", "Warning") | Out-Null
    return
  }
  if ($script:CurrentProcess -and -not $script:CurrentProcess.HasExited) {
    [System.Windows.Forms.MessageBox]::Show("OCR is already running.", "OCR", "OK", "Information") | Out-Null
    return
  }

  $pdfFullPath = (Resolve-Path -LiteralPath $PdfPath).Path
  $pdfDir = Split-Path -Parent $pdfFullPath
  $baseName = [System.IO.Path]::GetFileNameWithoutExtension($pdfFullPath)
  $script:OutputXlsx = Join-Path $pdfDir ($baseName + "_ocr.xlsx")
  New-Item -ItemType Directory -Force -Path $WorkDir | Out-Null
  Set-Content -LiteralPath $LogFile -Value "" -Encoding UTF8

  $script:LogBox.Clear()
  $script:OpenButton.Enabled = $false
  $script:StatusLabel.Text = "Starting..."
  $script:OutputLabel.Text = "Output: $script:OutputXlsx"
  Set-Busy $true
  Add-Log "PDF: $pdfFullPath"
  Add-Log "Output: $script:OutputXlsx"
  Add-Log "Log: $LogFile"

  $powershellExe = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
  $arguments = @(
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", (Quote-Arg $RunScript),
    "-Pdf", (Quote-Arg $pdfFullPath),
    "-OutXlsx", (Quote-Arg $script:OutputXlsx)
  ) -join " "

  $proc = New-Object System.Diagnostics.Process
  $proc.StartInfo.FileName = $powershellExe
  $proc.StartInfo.Arguments = $arguments
  $proc.StartInfo.WorkingDirectory = $pdfDir
  $proc.StartInfo.UseShellExecute = $false
  $proc.StartInfo.RedirectStandardOutput = $true
  $proc.StartInfo.RedirectStandardError = $true
  $proc.StartInfo.CreateNoWindow = $true
  $proc.EnableRaisingEvents = $true

  $proc.add_OutputDataReceived({
    param($sender, $eventArgs)
    if ($eventArgs.Data) {
      Add-Log $eventArgs.Data
    }
  })
  $proc.add_ErrorDataReceived({
    param($sender, $eventArgs)
    if ($eventArgs.Data) {
      Add-Log ("ERROR: " + $eventArgs.Data)
    }
  })
  $proc.add_Exited({
    param($sender, $eventArgs)
    $script:LastExitCode = $sender.ExitCode
    Add-Log ("Exit code: " + $script:LastExitCode)
    $action = [System.Action]{
      Set-Busy $false
      if ($script:LastExitCode -eq 0 -and (Test-Path -LiteralPath $script:OutputXlsx)) {
        $script:StatusLabel.Text = "Done."
        $script:OpenButton.Enabled = $true
        [System.Media.SystemSounds]::Asterisk.Play()
        [System.Windows.Forms.MessageBox]::Show("Excel file created:`n$script:OutputXlsx", "OCR complete", "OK", "Information") | Out-Null
      } else {
        $script:StatusLabel.Text = "Failed. Check the log."
        [System.Media.SystemSounds]::Exclamation.Play()
        [System.Windows.Forms.MessageBox]::Show("OCR failed. Check the log:`n$LogFile", "OCR failed", "OK", "Error") | Out-Null
      }
    }
    [void]$script:Form.BeginInvoke($action)
  })

  $script:CurrentProcess = $proc
  [void]$proc.Start()
  $proc.BeginOutputReadLine()
  $proc.BeginErrorReadLine()
}

[System.Windows.Forms.Application]::EnableVisualStyles()

$script:Form = New-Object System.Windows.Forms.Form
$script:Form.Text = "Management Fee OCR"
$script:Form.Width = 720
$script:Form.Height = 520
$script:Form.StartPosition = "CenterScreen"
$script:Form.MinimumSize = New-Object System.Drawing.Size(620, 420)

$script:DropPanel = New-Object System.Windows.Forms.Panel
$script:DropPanel.Anchor = "Top,Left,Right"
$script:DropPanel.Left = 16
$script:DropPanel.Top = 16
$script:DropPanel.Width = 670
$script:DropPanel.Height = 120
$script:DropPanel.BorderStyle = "FixedSingle"
$script:DropPanel.AllowDrop = $true
$script:DropPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 252)

$dropLabel = New-Object System.Windows.Forms.Label
$dropLabel.Dock = "Fill"
$dropLabel.TextAlign = "MiddleCenter"
$dropLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$dropLabel.Text = "Drop a PDF file here"
$script:DropPanel.Controls.Add($dropLabel)

$script:DropPanel.Add_DragEnter({
  if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
    $files = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    if ($files.Count -gt 0 -and [System.IO.Path]::GetExtension($files[0]).ToLowerInvariant() -eq ".pdf") {
      $_.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    }
  }
})
$script:DropPanel.Add_DragDrop({
  $files = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
  if ($files.Count -gt 0) {
    Start-Ocr $files[0]
  }
})

$script:SelectButton = New-Object System.Windows.Forms.Button
$script:SelectButton.Left = 16
$script:SelectButton.Top = 148
$script:SelectButton.Width = 120
$script:SelectButton.Height = 32
$script:SelectButton.Text = "Select PDF..."
$script:SelectButton.Add_Click({
  $dialog = New-Object System.Windows.Forms.OpenFileDialog
  $dialog.Filter = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*"
  $dialog.Multiselect = $false
  if ($dialog.ShowDialog($script:Form) -eq [System.Windows.Forms.DialogResult]::OK) {
    Start-Ocr $dialog.FileName
  }
})

$script:OpenButton = New-Object System.Windows.Forms.Button
$script:OpenButton.Left = 148
$script:OpenButton.Top = 148
$script:OpenButton.Width = 140
$script:OpenButton.Height = 32
$script:OpenButton.Text = "Open Output"
$script:OpenButton.Enabled = $false
$script:OpenButton.Add_Click({
  if ($script:OutputXlsx -and (Test-Path -LiteralPath $script:OutputXlsx)) {
    Start-Process explorer.exe "/select,`"$script:OutputXlsx`""
  }
})

$script:Progress = New-Object System.Windows.Forms.ProgressBar
$script:Progress.Anchor = "Top,Left,Right"
$script:Progress.Left = 304
$script:Progress.Top = 153
$script:Progress.Width = 382
$script:Progress.Height = 20
$script:Progress.Visible = $false

$script:StatusLabel = New-Object System.Windows.Forms.Label
$script:StatusLabel.Anchor = "Top,Left,Right"
$script:StatusLabel.Left = 16
$script:StatusLabel.Top = 190
$script:StatusLabel.Width = 670
$script:StatusLabel.Height = 22
$script:StatusLabel.Text = "Ready."

$script:OutputLabel = New-Object System.Windows.Forms.Label
$script:OutputLabel.Anchor = "Top,Left,Right"
$script:OutputLabel.Left = 16
$script:OutputLabel.Top = 214
$script:OutputLabel.Width = 670
$script:OutputLabel.Height = 22
$script:OutputLabel.Text = "Output: PDF folder, <filename>_ocr.xlsx"

$script:LogBox = New-Object System.Windows.Forms.TextBox
$script:LogBox.Anchor = "Top,Bottom,Left,Right"
$script:LogBox.Left = 16
$script:LogBox.Top = 244
$script:LogBox.Width = 670
$script:LogBox.Height = 220
$script:LogBox.Multiline = $true
$script:LogBox.ScrollBars = "Vertical"
$script:LogBox.ReadOnly = $true
$script:LogBox.Font = New-Object System.Drawing.Font("Consolas", 9)

$script:Form.Controls.AddRange(@(
  $script:DropPanel,
  $script:SelectButton,
  $script:OpenButton,
  $script:Progress,
  $script:StatusLabel,
  $script:OutputLabel,
  $script:LogBox
))

[void][System.Windows.Forms.Application]::Run($script:Form)
