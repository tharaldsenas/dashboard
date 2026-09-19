$ErrorActionPreference = "Stop"
$sourceFolder = "C:\Users\Administrator\OneDrive - Tharaldsen AS\Microsoft Copilot Chat-filer"
$outputPath = Join-Path $PSScriptRoot "order-reserve.json"
$excel = $null
$workbook = $null

try {
  $sourceFile = Get-ChildItem -LiteralPath $sourceFolder -File | Where-Object { $_.Name -match "150926\.xlsx$" } | Select-Object -First 1
  if (-not $sourceFile) { throw "Fant ikke månedsrapporten i $sourceFolder" }

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $workbook = $excel.Workbooks.Open($sourceFile.FullName, $false, $true)
  $sheet = $workbook.Worksheets.Item(1)
  $value = [double]$sheet.Range("K7").Value2

  if ([double]::IsNaN($value) -or [double]::IsInfinity($value)) {
    throw "K7 inneholder ikke et gyldig tall."
  }

  $payload = [ordered]@{
    value = $value
    source = "MonthlyReport!K7"
    updatedAt = (Get-Date).ToUniversalTime().ToString("o")
  }
  $payload | ConvertTo-Json | Set-Content -LiteralPath $outputPath -Encoding UTF8
}
finally {
  if ($workbook) {
    $workbook.Close($false)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
  }
  if ($excel) {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
