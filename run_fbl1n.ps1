param(
  [switch]$SkipExcelKill,
  [switch]$DumpDialogRadios,
  [int]$ClipboardMaxRetries = 10,
  [int]$ClipboardDelaySeconds = 2,
  [string[]]$CompanyCodes = @('8000', '8100'),
  [string]$ExportRadioId = ""
)

# Helper runner to export FBL1N open items for AU/NZ (8000/8100) via SAP GUI scripting.
# Uses VBScript for SAP automation (Clipboard export), PowerShell for Excel conversion.

Set-Location 'C:\Users\Azhao.PIVOTAL\Downloads\AP'

$keyDateInput = Read-Host "Enter key date dd/MM/yyyy (default 05/12/2025)"
if ([string]::IsNullOrWhiteSpace($keyDateInput)) { $keyDateInput = '05/12/2025' }

# Normalize date input (replace . with /)
$keyDateInput = $keyDateInput -replace '\.', '/'

# Calculate date file name
$dateParts = $keyDateInput.Split('/')
$day = $dateParts[0]
$month = $dateParts[1]
$year = $dateParts[2]

# Handle 2-digit year
if ($year.Length -eq 2) {
  $shortYear = $year
  $year = "20" + $year
}
else {
  $shortYear = $year.Substring(2)
}

$dateFile = "$day.$month.$shortYear"

# Ensure output directory exists
$outputDir = Resolve-Path "02-inputs/Payment run raw"
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir -Force | Out-Null }

function Stop-ExcelIfNeeded {
  param([string]$Reason)
  if ($SkipExcelKill) {
    Write-Host "SkipExcelKill set - not closing Excel ($Reason)."
    return
  }
  Stop-Process -Name "excel" -ErrorAction SilentlyContinue
  Write-Host "Closed existing Excel instances ($Reason)."
}

Stop-ExcelIfNeeded -Reason "startup"

foreach ($cc in $CompanyCodes) {
  Write-Host "Exporting $cc..."
  
  $subfolder = if ($cc -eq '8000') { "AU" } else { "NZ" }
  $targetDir = Join-Path $outputDir $subfolder
  if (-not (Test-Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }
  $targetDirStr = $targetDir.ToString()
  
  Write-Host "  Target: $targetDirStr"
  
  $vbsPath = Resolve-Path "01-system/tools/ops/sap-fbl1n/sap_fbl1n_export.vbs"
    
  # Run VBScript
  $tokens = @()
  if ($DumpDialogRadios) { $tokens += 'dump' }
  if ($ExportRadioId) { $tokens += "radio=$ExportRadioId" }
  $extraArgs = $tokens | ForEach-Object { ' "' + $_ + '"' } | Out-String
  $argsString = "//Nologo ""$vbsPath"" ""$cc"" ""$keyDateInput"" ""$targetDirStr"" ""/AZ_AP_AG""" + $extraArgs
  $process = Start-Process -FilePath "cscript" -ArgumentList $argsString -Wait -NoNewWindow -PassThru
  
  if ($process.ExitCode -ne 0) {
    Write-Error "VBScript failed for Company Code $cc with exit code $($process.ExitCode)."
    continue
  }
  
  # Get data from Clipboard with retry
  Write-Host "  Getting data from Clipboard..."
  $clipboardText = $null
  $retryCount = 0
  while ([string]::IsNullOrWhiteSpace($clipboardText) -and $retryCount -lt $ClipboardMaxRetries) {
    $clipboardText = Get-Clipboard
    if ([string]::IsNullOrWhiteSpace($clipboardText)) {
      Write-Host "    Clipboard empty, retrying in $ClipboardDelaySeconds s... ($($retryCount + 1)/$ClipboardMaxRetries)"
      Start-Sleep -Seconds $ClipboardDelaySeconds
      $retryCount++
    }
  }
  
  if ([string]::IsNullOrWhiteSpace($clipboardText)) {
    Write-Error "  Clipboard is empty after retries!"
    continue
  }
  
  Write-Host "  Creating Excel instance..."
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  
  try {
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Sheets.Item(1)
      
    # Paste data
    Write-Host "  Pasting data..."
    $ws.Range("A1").Select()
    $ws.Paste()
      
    # Save as XLSX
    $targetFile = Join-Path $targetDirStr "$dateFile.xlsx"
    if (Test-Path $targetFile) { Remove-Item $targetFile -Force }
      
    $wb.SaveAs($targetFile, 51) # xlOpenXMLWorkbook
    Write-Host "  File saved successfully: $targetFile"
      
    $wb.Close($false)
  }
  catch {
    Write-Error "  Error saving Excel file: $_"
  }
  finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  }
  
  Write-Host "Export for $cc completed."
  
  # Clear clipboard safely
  try { Set-Clipboard -Value " " } catch { }
  
  # Ensure Excel is closed before next run
  Stop-ExcelIfNeeded -Reason "post-export"
  Start-Sleep -Seconds 2
}

Write-Host "Done. Check 02-inputs/Payment run raw for FBL1N exports."
