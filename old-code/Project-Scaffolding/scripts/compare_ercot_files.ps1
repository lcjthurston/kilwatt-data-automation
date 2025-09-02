# Import the ImportExcel module
Import-Module ImportExcel

# Define file paths
$generatedFile = "ERCOT-new-product-term-formulas.xlsx"
$originalFile = "original.xlsx"
$differencesFile = "ERCOT-comparison-differences.xlsx"

Write-Host "Starting comparison between generated and original files..." -ForegroundColor Green

# Check if files exist
if (-not (Test-Path $generatedFile)) {
    Write-Host "ERROR: Generated file '$generatedFile' not found!" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $originalFile)) {
    Write-Host "ERROR: Original file '$originalFile' not found!" -ForegroundColor Red
    exit 1
}

try {
    # Import both Excel files
    Write-Host "Loading generated file: $generatedFile" -ForegroundColor Yellow
    $generatedData = Import-Excel -Path $generatedFile
    
    Write-Host "Loading original file: $originalFile" -ForegroundColor Yellow
    $originalData = Import-Excel -Path $originalFile
    
    Write-Host "Generated file rows: $($generatedData.Count)" -ForegroundColor Cyan
    Write-Host "Original file rows: $($originalData.Count)" -ForegroundColor Cyan
    
    # Get column names from both files
    $generatedColumns = $generatedData[0].PSObject.Properties.Name
    $originalColumns = $originalData[0].PSObject.Properties.Name
    
    Write-Host "`nGenerated file columns: $($generatedColumns -join ', ')" -ForegroundColor Cyan
    Write-Host "Original file columns: $($originalColumns -join ', ')" -ForegroundColor Cyan
    
    # Define columns to compare
    $columnsAH = @('Start Month', 'State', 'Utility', 'Congestion Zone', 'Load Factor', 'Term', 'Product', '0-200,000')
    $columnsJAA = @('J Index', 'K Concat', 'L ConstDate', 'M =B', 'N Region', 'O LF Norm', 'P Supplier', 'Q TermMonths', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA')
    
    # Find actual column names that exist in both files
    $actualColumnsAH = $columnsAH | Where-Object { $_ -in $generatedColumns -and $_ -in $originalColumns }
    $actualColumnsJAA = $columnsJAA | Where-Object { $_ -in $generatedColumns -and $_ -in $originalColumns }
    
    Write-Host "`nColumns A-H found in both files: $($actualColumnsAH -join ', ')" -ForegroundColor Green
    Write-Host "Columns J-AA found in both files: $($actualColumnsJAA -join ', ')" -ForegroundColor Green
    
    # Compare the data
    Write-Host "`nComparing data..." -ForegroundColor Yellow
    $differences = Compare-Object -ReferenceObject $originalData -DifferenceObject $generatedData -PassThru
    
    if ($differences) {
        Write-Host "Found $($differences.Count) differences!" -ForegroundColor Red
        
        # Export differences to Excel file
        $differences | Export-Excel -Path $differencesFile -AutoSize -TableStyle Medium2
        Write-Host "Differences exported to: $differencesFile" -ForegroundColor Yellow
        
        # Display summary of differences
        Write-Host "`n=== DIFFERENCES SUMMARY ===" -ForegroundColor Red
        $differences | ForEach-Object {
            $side = if ($_.SideIndicator -eq "<=") { "ORIGINAL ONLY" } else { "GENERATED ONLY" }
            Write-Host "$side - Row with values: $($_.PSObject.Properties.Value -join ' | ')" -ForegroundColor Red
        }
        
        # Detailed column-by-column comparison for A-H columns
        Write-Host "`n=== DETAILED COMPARISON FOR COLUMNS A-H ===" -ForegroundColor Magenta
        foreach ($col in $actualColumnsAH) {
            $genValues = $generatedData | Select-Object -ExpandProperty $col
            $origValues = $originalData | Select-Object -ExpandProperty $col
            
            $colDiffs = Compare-Object -ReferenceObject $origValues -DifferenceObject $genValues
            if ($colDiffs) {
                Write-Host "Column '$col' has differences:" -ForegroundColor Red
                $colDiffs | ForEach-Object {
                    $side = if ($_.SideIndicator -eq "<=") { "ORIGINAL" } else { "GENERATED" }
                    Write-Host "  $side: $($_.InputObject)" -ForegroundColor Red
                }
            } else {
                Write-Host "Column '$col': MATCH" -ForegroundColor Green
            }
        }
        
        # Detailed column-by-column comparison for J-AA columns
        Write-Host "`n=== DETAILED COMPARISON FOR FORMULA COLUMNS J-AA ===" -ForegroundColor Magenta
        foreach ($col in $actualColumnsJAA) {
            $genValues = $generatedData | Select-Object -ExpandProperty $col
            $origValues = $originalData | Select-Object -ExpandProperty $col
            
            $colDiffs = Compare-Object -ReferenceObject $origValues -DifferenceObject $genValues
            if ($colDiffs) {
                Write-Host "Formula Column '$col' has differences:" -ForegroundColor Red
                $colDiffs | ForEach-Object {
                    $side = if ($_.SideIndicator -eq "<=") { "ORIGINAL" } else { "GENERATED" }
                    Write-Host "  $side: $($_.InputObject)" -ForegroundColor Red
                }
            } else {
                Write-Host "Formula Column '$col': MATCH" -ForegroundColor Green
            }
        }
        
    } else {
        Write-Host "SUCCESS: No differences found! Files match perfectly." -ForegroundColor Green
    }
    
} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}

Write-Host "`nComparison complete." -ForegroundColor Green
