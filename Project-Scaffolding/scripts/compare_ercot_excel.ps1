param(
    [Parameter(Mandatory=$false)] [string]$FileA = "./ERCOT-new-product-term-formulas.xlsx",
    [Parameter(Mandatory=$false)] [string]$FileB = "./original.xlsx",
    [Parameter(Mandatory=$false)] [string]$OutFile = "./differences.xlsx",
    [Parameter(Mandatory=$false)] [string[]]$Columns = @(
        'Start Month','State','Utility','Congestion Zone','Load Factor','Term','Product','0-200,000'
    ),
    [Parameter(Mandatory=$false)] [string]$WorksheetName
)

# Ensure ImportExcel is available
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "The ImportExcel module is required. Install it with:" -ForegroundColor Yellow
    Write-Host "  Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor Yellow
    exit 5
}

Import-Module ImportExcel -ErrorAction Stop

# Helper to intersect columns that exist in both sheets
function Get-CommonColumns {
    param($aRows, $bRows, $desired)
    $aCols = @($aRows | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name)
    $bCols = @($bRows | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name)
    if (-not $aCols -or -not $bCols) { return @() }
    $desired | Where-Object { $aCols -contains $_ -and $bCols -contains $_ }
}

# List sheets
$infoA = Get-ExcelSheetInfo -Path $FileA
$infoB = Get-ExcelSheetInfo -Path $FileB

$sheetNamesA = @($infoA | Select-Object -ExpandProperty Name)
$sheetNamesB = @($infoB | Select-Object -ExpandProperty Name)

$targetSheets = @()
if ($WorksheetName) {
    if ($sheetNamesA -contains $WorksheetName -and $sheetNamesB -contains $WorksheetName) {
        $targetSheets = @($WorksheetName)
    } else {
        Write-Host "Worksheet '$WorksheetName' not found in both files. Falling back to common sheets." -ForegroundColor Yellow
    }
}
if (-not $targetSheets) {
    $targetSheets = @($sheetNamesA | Where-Object { $sheetNamesB -contains $_ })
}
if (-not $targetSheets) {
    Write-Error "No common worksheets found between $FileA and $FileB"
    exit 2
}

# Prepare output file (remove if exists)
if (Test-Path $OutFile) { Remove-Item $OutFile -Force }

foreach ($sheet in $targetSheets) {
    Write-Host "Comparing sheet: $sheet" -ForegroundColor Cyan
    $a = Import-Excel -Path $FileA -WorksheetName $sheet
    $b = Import-Excel -Path $FileB -WorksheetName $sheet

    # Determine comparable columns
    $cols = Get-CommonColumns -aRows $a -bRows $b -desired $Columns
    if (-not $cols -or $cols.Count -eq 0) {
        Write-Host "  No comparable columns (from desired set) found on this sheet. Skipping." -ForegroundColor Yellow
        continue
    }

    # Align to common columns and normalize ordering
    $aSel = $a | Select-Object -Property $cols
    $bSel = $b | Select-Object -Property $cols

    # Compare and collect differences
    $diff = Compare-Object -ReferenceObject $aSel -DifferenceObject $bSel -Property $cols -PassThru

    # Export per-sheet results
    if ($diff) {
        $diff | Export-Excel -Path $OutFile -WorksheetName ("{0}-diff" -f $sheet) -AutoSize -AutoFilter -Append
    } else {
        # Write a tiny table saying no differences
        [PSCustomObject]@{ Status = 'No differences'; Sheet = $sheet } |
            Export-Excel -Path $OutFile -WorksheetName ("{0}-diff" -f $sheet) -AutoSize -AutoFilter -Append
    }

    # Also export the exact columns being compared for traceability
    ($aSel | Select-Object -First 5) | Export-Excel -Path $OutFile -WorksheetName ("{0}-sampleA" -f $sheet) -AutoSize -AutoFilter -Append
    ($bSel | Select-Object -First 5) | Export-Excel -Path $OutFile -WorksheetName ("{0}-sampleB" -f $sheet) -AutoSize -AutoFilter -Append
}

Write-Host "Comparison complete. Output: $OutFile" -ForegroundColor Green

