# Requires: ImportExcel module
# Compares VALUES (not formulas) between two Excel files for columns A–H and J–AA.
# Files must have the same sheet names; R is included but values-only (should be blank in both for now).

param(
  [string]$Generated = "ERCOT-new-product-term-formulas.xlsx",
  [string]$Original  = "original.xlsx",
  [string]$OutFile   = "ERCOT-value-differences.xlsx"
)

Import-Module ImportExcel -ErrorAction Stop

function Get-Excel-ColumnsOrdered($path, $sheetName) {
  $data = Import-Excel -Path $path -WorksheetName $sheetName -ErrorAction Stop
  return $data
}

function Get-SheetNames($path) {
  ($path | Get-ExcelSheetInfo).Name
}

function Get-ColLetter([int]$index) {
  # 1 -> A, 27 -> AA
  $col = ""
  $i = $index
  while ($i -gt 0) {
    $mod = ($i - 1) % 26
    $col = [char](65 + $mod) + $col
    $i = [math]::Floor(($i - $mod) / 26)
    $i--
  }
  $col
}

function Values-Equal($a, $b) {
  if ($null -eq $a -and $null -eq $b) { return $true }
  if ($null -eq $a -or  $null -eq $b) { return $false }
  # Try numeric compare
  $da = $null; $db = $null
  $isNumA = [double]::TryParse($a.ToString(), [ref]$da)
  $isNumB = [double]::TryParse($b.ToString(), [ref]$db)
  if ($isNumA -and $isNumB) {
    return [math]::Abs($da - $db) -lt 1e-9
  }
  # Try date compare
  $ta = $null; $tb = $null
  $isDateA = [datetime]::TryParse($a.ToString(), [ref]$ta)
  $isDateB = [datetime]::TryParse($b.ToString(), [ref]$tb)
  if ($isDateA -and $isDateB) {
    return ($ta.ToString('yyyy-MM-dd')) -eq ($tb.ToString('yyyy-MM-dd'))
  }
  # Fallback: trimmed string compare
  return ($a.ToString().Trim()) -eq ($b.ToString().Trim())
}

# Determine sheet names to compare
if (-not (Test-Path $Generated)) { throw "Generated file not found: $Generated" }
if (-not (Test-Path $Original))  { throw "Original file not found: $Original" }

$genSheets  = Get-SheetNames $Generated
$origSheets = Get-SheetNames $Original
$commonSheets = $genSheets | Where-Object { $_ -in $origSheets }
if (-not $commonSheets) { throw "No common sheet names between files." }

$diffRows = @()

foreach ($sheet in $commonSheets) {
  $gen = Get-Excel-ColumnsOrdered -path $Generated -sheetName $sheet
  $org = Get-Excel-ColumnsOrdered -path $Original  -sheetName $sheet

  $rowCount = [Math]::Max($gen.Count, $org.Count)
  if ($gen.Count -ne $org.Count) {
    Write-Warning "Row count mismatch in sheet '$sheet': generated=$($gen.Count), original=$($org.Count)"
  }

  # Property order reflects column order from Excel
  $genCols = if ($gen.Count -gt 0) { $gen[0].PSObject.Properties.Name } else { @() }
  $orgCols = if ($org.Count -gt 0) { $org[0].PSObject.Properties.Name } else { @() }

  # Sanity check: ensure both have at least 27 columns (AA)
  if ($genCols.Count -lt 27 -or $orgCols.Count -lt 27) {
    throw "Expecting at least 27 columns (A..AA). Sheet '$sheet' has generated=$($genCols.Count), original=$($orgCols.Count)."
  }

  $indexesAH  = 1..8
  $indexesJAA = 10..27

  foreach ($r in 0..($rowCount-1)) {
    $genRow = if ($r -lt $gen.Count) { $gen[$r] } else { $null }
    $orgRow = if ($r -lt $org.Count) { $org[$r] } else { $null }

    foreach ($idx in ($indexesAH + $indexesJAA)) {
      $colLetter = Get-ColLetter $idx
      $genProp = $genCols[$idx-1]
      $orgProp = $orgCols[$idx-1]

      $genVal = if ($genRow) { $genRow.$genProp } else { $null }
      $orgVal = if ($orgRow) { $orgRow.$orgProp } else { $null }

      if (-not (Values-Equal $genVal $orgVal)) {
        $diffRows += [pscustomobject]@{
          Sheet          = $sheet
          Row            = $r + 2  # account for header row in Excel
          ColumnLetter   = $colLetter
          GeneratedValue = $genVal
          OriginalValue  = $orgVal
        }
      }
    }
  }
}

if ($diffRows.Count -gt 0) {
  $diffRows | Export-Excel -Path $OutFile -AutoSize -TableStyle Medium2
  Write-Host "Differences found: $($diffRows.Count). Exported to $OutFile" -ForegroundColor Yellow
} else {
  Write-Host "SUCCESS: All compared values match (A–H and J–AA)." -ForegroundColor Green
}

