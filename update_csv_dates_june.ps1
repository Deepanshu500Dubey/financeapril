# Update all CSV files from May to June dates
Get-ChildItem -Filter '*.csv' | ForEach-Object {
    $content = Get-Content $_.FullName -Raw
    $content = $content -replace '2026-05', '2026-06'
    $content = $content -replace '2025-05', '2025-06'
    $content = $content -replace 'May 2026', 'Jun 2026'
    $content = $content -replace 'May_2025', 'Jun_2025'
    $content = $content -replace 'May', 'June'
    $content = $content -replace ',May,', ',Jun,'
    $content = $content -replace ' May ', ' Jun '
    Set-Content -Path $_.FullName -Value $content -NoNewline
    Write-Host "Updated: $($_.Name)"
}
Write-Host "All CSV files updated successfully!"