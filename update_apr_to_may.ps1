# Script to update April date references in CSV files to May
# This will copy April CSV files and update them for May closing

$sourceDir = "E:\Octane\Finance_workflow\financeapril"

# List of April CSV files to convert to May
$aprilFiles = @(
    "IBMBOB_Workforce_Cost_Output_Apr2026.csv"
)

# Additional files that contain April data that need May versions
$filesToCopy = @(
    "AR_Subledger_Mar2026.csv",
    "Accrual_Adjustment_Journals_Mar2026.csv",
    "Accruals_Register_Mar2026.csv",
    "Bank_Reconciliation_Items_Mar2026.csv",
    "Bank_Reconciliation_Journals_Mar2026.csv",
    "Bank_Statements_Mar2026.csv",
    "Budget_Mar2026_Detailed.csv",
    "GL_Cash_Balances_Mar2026.csv",
    "Intercompany_Elimination_Journals_Mar2026.csv",
    "Intercompany_Elimination_Journals_v2_Mar2026.csv",
    "Intercompany_Reconciliation_Mar2026.csv",
    "Intercompany_Transactions_Mar2026.csv",
    "Prepayments_Register_Mar2026.csv",
    "Prepayment_Amortization_Journals_Mar2026.csv",
    "Raw_GL_Export_With_CostCenters_Mar2026.csv"
)

Write-Host "Starting conversion from April to May 2026 data..." -ForegroundColor Green

# Copy April files and update to May
foreach ($file in $aprilFiles) {
    $sourcePath = Join-Path $sourceDir $file
    $destPath = Join-Path $sourceDir ($file -replace 'Apr2026', 'May2026')
    
    if (Test-Path $sourcePath) {
        Copy-Item -Path $sourcePath -Destination $destPath -Force
        Write-Host "Copied: $file -> $(Split-Path $destPath -Leaf)"
        
        # Update date references within the copied file
        $content = Get-Content $destPath -Raw
        $content = $content -replace '2026-04', '2026-05'
        $content = $content -replace 'Apr 2026', 'May 2026'
        $content = $content -replace 'Apr_2026', 'May_2026'
        $content = $content -replace '\bApr\b', 'May'
        Set-Content -Path $destPath -Value $content -NoNewline
        Write-Host "Updated dates in: $(Split-Path $destPath -Leaf)"
    }
    else {
        Write-Host "File not found: $file" -ForegroundColor Yellow
    }
}

# Create May versions from existing March files (they already contain April data that should become May)
foreach ($file in $filesToCopy) {
    $sourcePath = Join-Path $sourceDir $file
    
    if (Test-Path $sourcePath) {
        # Create May version by replacing Apr with May in content
        $content = Get-Content $sourcePath -Raw
        
        # Replace date references
        $content = $content -replace '2026-04', '2026-05'
        $content = $content -replace '2025-04', '2025-05'
        $content = $content -replace 'Apr 2026', 'May 2026'
        $content = $content -replace 'Apr_2026', 'May_2026'
        $content = $content -replace 'Apr_2025', 'May_2025'
        $content = $content -replace '\bApr\b', 'May'
        $content = $content -replace '\bApril\b', 'May'
        
        # Create new filename for May
        $newFileName = $file -replace 'Mar2026', 'May2026'
        $destPath = Join-Path $sourceDir $newFileName
        
        Set-Content -Path $destPath -Value $content -NoNewline
        Write-Host "Created May version: $newFileName"
    }
}

Write-Host "`nAll files updated successfully for May 2026 closing!" -ForegroundColor Green
Write-Host "Next: Update app.py to reference May2026 files instead of Mar2026"
