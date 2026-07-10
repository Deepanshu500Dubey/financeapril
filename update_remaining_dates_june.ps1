# Script to update remaining May date references in CSV files

$csvFiles = Get-ChildItem -Path "." -Filter "*_Jun2026.csv"

foreach ($file in $csvFiles) {
    Write-Host "Processing: $($file.Name)"
    
    $content = Get-Content $file.FullName -Raw
    
    $content = $content -replace '(\d{1,2})[-/]05[-/](2026)', '$1-06-$2'
    $content = $content -replace '05/(\d{1,2})/(2026)', '06/$1/$2'
    $content = $content -replace '(2026)-05-(\d{1,2})', '$1-06-$2'
    $content = $content -replace '(\d{1,2})[-/]05[-/](2025)', '$1-06-$2'
    $content = $content -replace '05/(\d{1,2})/(2025)', '06/$1/$2'
    $content = $content -replace '(2025)-05-(\d{1,2})', '$1-06-$2'
    $content = $content -replace '\bMay\b', 'Jun'
    $content = $content -replace '\bMay_2025\b', 'Jun_2025'
    $content = $content -replace '\bMay_2026\b', 'Jun_2026'
    $content = $content -replace 'May_2025_Actual', 'Jun_2025_Actual'
    $content = $content -replace 'May_2026_Actual', 'Jun_2026_Actual'
    
    Set-Content -Path $file.FullName -Value $content -NoNewline
    Write-Host "Updated: $($file.Name)"
}

Write-Host "`nAll remaining May dates updated to June!"