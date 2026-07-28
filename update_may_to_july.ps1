# Script to update May date references in CSV files to July
# This will copy May CSV files and update them for July closing

$sourceDir = "."  # Use current folder

# List of May CSV files to convert to July
$mayFiles = @(
    "IBMBOB_Workforce_Cost_Output_May2026.csv"
)

# Additional files that contain May data that need July versions
# These are the April files that you want to convert to July
$filesToCopy = @(
    "AR_Subledger_May2026.csv",
    "Accrual_Adjustment_Journals_May2026.csv",
    "Accruals_Register_May2026.csv",
    "Bank_Reconciliation_Items_May2026.csv",
    "Bank_Reconciliation_Journals_May2026.csv",
    "Bank_Statements_May2026.csv",
    "Budget_May2026_Detailed.csv",
    "GL_Cash_Balances_May2026.csv",
    "Intercompany_Elimination_Journals_May2026.csv",
    "Intercompany_Elimination_Journals_v2_May2026.csv",
    "Intercompany_Reconciliation_May2026.csv",
    "Intercompany_Transactions_May2026.csv",
    "Prepayments_Register_May2026.csv",
    "Prepayment_Amortization_Journals_May2026.csv",
    "Raw_GL_Export_With_CostCenters_May2026.csv"
)

Write-Host "Starting conversion from May to July 2026 data..." -ForegroundColor Green
Write-Host "Source directory: $sourceDir" -ForegroundColor Cyan

# Copy May files and update to July
foreach ($file in $mayFiles) {
    $sourcePath = Join-Path $sourceDir $file
    $destPath = Join-Path $sourceDir ($file -replace 'May2026', 'Jul2026')
    
    if (Test-Path $sourcePath) {
        Copy-Item -Path $sourcePath -Destination $destPath -Force
        Write-Host "Copied: $file -> $(Split-Path $destPath -Leaf)"
        
        $content = Get-Content $destPath -Raw
        $content = $content -replace '2026-05', '2026-07'
        $content = $content -replace 'May 2026', 'Jul 2026'
        $content = $content -replace 'May_2026', 'Jul_2026'
        $content = $content -replace '\bMay\b', 'Jul'
        Set-Content -Path $destPath -Value $content -NoNewline
        Write-Host "Updated dates in: $(Split-Path $destPath -Leaf)"
    }
    else {
        Write-Host "File not found: $file" -ForegroundColor Yellow
    }
}

# Create July versions from existing May files
foreach ($file in $filesToCopy) {
    $sourcePath = Join-Path $sourceDir $file
    
    if (Test-Path $sourcePath) {
        Write-Host "Processing: $file"
        $content = Get-Content $sourcePath -Raw
        
        # Replace date references
        $content = $content -replace '2026-05', '2026-07'
        $content = $content -replace '2025-05', '2025-07'
        $content = $content -replace 'May 2026', 'Jul 2026'
        $content = $content -replace 'May_2026', 'Jul_2026'
        $content = $content -replace 'May_2025', 'Jul_2025'
        $content = $content -replace '\bMay\b', 'Jul'
        
        # Create new filename for July
        $newFileName = $file -replace 'May2026', 'Jul2026'
        $destPath = Join-Path $sourceDir $newFileName
        
        Set-Content -Path $destPath -Value $content -NoNewline
        Write-Host "Created July version: $newFileName"
    }
    else {
        Write-Host "File not found: $file" -ForegroundColor Yellow
    }
}

Write-Host "`nAll files updated successfully for July 2026 closing!" -ForegroundColor Green
Write-Host "Next: Update app.py to reference Jul2026 files instead of May2026"