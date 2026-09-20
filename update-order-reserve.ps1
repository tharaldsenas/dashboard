$ErrorActionPreference = "Stop"
$sourceFolder = "C:\Users\Administrator\OneDrive - Tharaldsen AS\Microsoft Copilot Chat-filer"
$outputPath = Join-Path $PSScriptRoot "order-reserve.json"
$operatingResultOutputPath = Join-Path $PSScriptRoot "operating-result.json"
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

  $latestMonthColumn = 0
  $parsedMonthValue = 0
  foreach ($column in 3..14) {
    $monthValue = $sheet.Cells.Item(17, $column).Value2
    if ($null -ne $monthValue -and [string]$monthValue -ne "") {
      if (-not [double]::TryParse([string]$monthValue, [ref]$parsedMonthValue)) {
        throw "Rad 17 inneholder en ugyldig månedsverdi i kolonne $column."
      }
      $latestMonthColumn = $column
    }
  }
  if ($latestMonthColumn -eq 0) { throw "Fant ingen siste måned i rad 17." }

  $operatingResult = [double]$sheet.Range("O17").Value2
  $previousOperatingResult = 0
  foreach ($column in 3..$latestMonthColumn) {
    $previousMonthValue = $sheet.Cells.Item(26, $column).Value2
    if ($null -ne $previousMonthValue -and [string]$previousMonthValue -ne "") {
      $previousOperatingResult += [double]$previousMonthValue
    }
  }
  $monthName = [string]$sheet.Cells.Item(19, $latestMonthColumn).Text
  $operatingResultPayload = [ordered]@{
    current = $operatingResult
    previous = $previousOperatingResult
    period = "til og med $monthName"
    source = "MonthlyReport!O17 og C26:$([char](64 + $latestMonthColumn))26"
    updatedAt = (Get-Date).ToUniversalTime().ToString("o")
  }
  $operatingResultPayload | ConvertTo-Json | Set-Content -LiteralPath $operatingResultOutputPath -Encoding UTF8
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
