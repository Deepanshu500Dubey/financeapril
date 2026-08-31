# Update all CSV files from May to August dates
Get-ChildItem -Filter '*.csv' | ForEach-Object {
    $content = Get-Content $_.FullName -Raw
    $content = $content -replace '2026-05', '2026-08'
    $content = $content -replace '2025-05', '2025-08'
    $content = $content -replace 'May 2026', 'Aug 2026'
    $content = $content -replace 'May_2025', 'Aug_2025'
    $content = $content -replace 'May', 'August'
    $content = $content -replace ',May,', ',Aug,'
    $content = $content -replace ' May ', ' Aug '
    Set-Content -Path $_.FullName -Value $content -NoNewline
    Write-Host "Updated: $($_.Name)"
}
Write-Host "All CSV files updated successfully!"