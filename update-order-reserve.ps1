$ErrorActionPreference = "Stop"
$sourceFolder = "C:\Users\Administrator\OneDrive - Tharaldsen AS\Microsoft Copilot Chat-filer"
$outputPath = Join-Path $PSScriptRoot "order-reserve.json"
$operatingResultOutputPath = Join-Path $PSScriptRoot "operating-result.json"
$revenueOutputPath = Join-Path $PSScriptRoot "revenue-summary.json"
$revenueHistoryOutputPath = Join-Path $PSScriptRoot "revenue-history.json"
$operatingResultHistoryOutputPath = Join-Path $PSScriptRoot "operating-result-history.json"
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

  $latestRevenueMonthColumn = 0
  $parsedRevenueMonthValue = 0
  foreach ($column in 3..14) {
    $monthValue = $sheet.Cells.Item(11, $column).Value2
    $monthText = ([string]$sheet.Cells.Item(11, $column).Text).Trim()
    if ($monthText -ne "-" -and $null -ne $monthValue -and [double]::TryParse([string]$monthValue, [ref]$parsedRevenueMonthValue)) {
      $latestRevenueMonthColumn = $column
    }
  }
  if ($latestRevenueMonthColumn -eq 0) { throw "Fant ingen siste måned i omsetningsraden." }

  $revenue = [double]$sheet.Range("O11").Value2
  $previousRevenue = 0
  foreach ($column in 3..$latestRevenueMonthColumn) {
    $previousMonthValue = $sheet.Cells.Item(20, $column).Value2
    if ($null -ne $previousMonthValue -and [string]$previousMonthValue -ne "") {
      $previousRevenue += [double]$previousMonthValue
    }
  }
  $revenueMonthName = [string]$sheet.Cells.Item(10, $latestRevenueMonthColumn).Text
  $revenuePayload = [ordered]@{
    current = $revenue
    previous = $previousRevenue
    period = "til og med $revenueMonthName"
    source = "MonthlyReport!O11 og C20:$([char](64 + $latestRevenueMonthColumn))20"
    updatedAt = (Get-Date).ToUniversalTime().ToString("o")
  }
  $revenuePayload | ConvertTo-Json | Set-Content -LiteralPath $revenueOutputPath -Encoding UTF8

  $monthLabels = @()
  $revenueValues2026 = @()
  $revenueValues2025 = @()
  $revenueValues2024 = @()
  foreach ($column in 3..14) {
    $monthLabels += ([string]$sheet.Cells.Item(10, $column).Text).Trim()
    foreach ($row in @(11, 20, 29)) {
      $cellText = ([string]$sheet.Cells.Item($row, $column).Text).Trim()
      $cellValue = $sheet.Cells.Item($row, $column).Value2
      $parsedValue = 0
      $value = $null
      if ($cellText -ne "-" -and $null -ne $cellValue -and [double]::TryParse([string]$cellValue, [ref]$parsedValue)) {
        $value = [double]$cellValue
      }
      if ($row -eq 11) { $revenueValues2026 += $value }
      elseif ($row -eq 20) { $revenueValues2025 += $value }
      else { $revenueValues2024 += $value }
    }
  }
  $revenueHistoryPayload = [ordered]@{
    months = $monthLabels
    series = @(
      [ordered]@{ year = 2026; values = $revenueValues2026 }
      [ordered]@{ year = 2025; values = $revenueValues2025 }
      [ordered]@{ year = 2024; values = $revenueValues2024 }
    )
    source = "MonthlyReport!C11:N11, C20:N20 og C29:N29"
    updatedAt = (Get-Date).ToUniversalTime().ToString("o")
  }
  $revenueHistoryPayload | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $revenueHistoryOutputPath -Encoding UTF8

  $resultValues2026 = @()
  $resultValues2025 = @()
  $resultValues2024 = @()
  foreach ($column in 3..14) {
    foreach ($row in @(17, 26, 34)) {
      $cellText = ([string]$sheet.Cells.Item($row, $column).Text).Trim()
      $cellValue = $sheet.Cells.Item($row, $column).Value2
      $parsedValue = 0
      $value = $null
      if ($cellText -ne "-" -and $null -ne $cellValue -and [double]::TryParse([string]$cellValue, [ref]$parsedValue)) {
        $value = [double]$cellValue
      }
      if ($row -eq 17) { $resultValues2026 += $value }
      elseif ($row -eq 26) { $resultValues2025 += $value }
      else { $resultValues2024 += $value }
    }
  }
  $operatingResultHistoryPayload = [ordered]@{
    months = $monthLabels
    series = @(
      [ordered]@{ year = 2026; values = $resultValues2026 }
      [ordered]@{ year = 2025; values = $resultValues2025 }
      [ordered]@{ year = 2024; values = $resultValues2024 }
    )
    source = "MonthlyReport!C17:N17, C26:N26 og C34:N34"
    updatedAt = (Get-Date).ToUniversalTime().ToString("o")
  }
  $operatingResultHistoryPayload | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $operatingResultHistoryOutputPath -Encoding UTF8
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
