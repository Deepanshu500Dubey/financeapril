# Update all CSV files from May to July dates
Get-ChildItem -Filter '*.csv' | ForEach-Object {
    $content = Get-Content $_.FullName -Raw
    $content = $content -replace '2026-05', '2026-07'
    $content = $content -replace '2025-05', '2025-07'
    $content = $content -replace 'May 2026', 'Jul 2026'
    $content = $content -replace 'May_2025', 'Jul_2025'
    $content = $content -replace 'May', 'July'
    $content = $content -replace ',May,', ',Jul,'
    $content = $content -replace ' May ', ' Jul '
    Set-Content -Path $_.FullName -Value $content -NoNewline
    Write-Host "Updated: $($_.Name)"
}
Write-Host "All CSV files updated successfully!"