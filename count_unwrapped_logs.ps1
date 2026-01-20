# Count File.AppendAllText/WriteAllText calls that are NOT wrapped with DeploymentConfiguration.DeploymentMode check

$unwrappedCount = 0
$totalCount = 0
$files = Get-ChildItem -Path "Services" -Recurse -Filter "*.cs" | Where-Object { $_.FullName -notlike "*\Backup\*" -and $_.FullName -notlike "*\Archive\*" }

foreach ($file in $files) {
    $content = Get-Content $file.FullName -Raw
    $lines = Get-Content $file.FullName
    
    for ($i = 0; $i -lt $lines.Length; $i++) {
        if ($lines[$i] -match "File\.(AppendAllText|WriteAllText)") {
            $totalCount++
            
            # Check if previous 5 lines contain deployment mode check
            $isWrapped = $false
            $startCheck = [Math]::Max(0, $i - 5)
            
            for ($j = $startCheck; $j -le $i; $j++) {
                if ($lines[$j] -match "if\s*\(!?\s*DeploymentConfiguration\.DeploymentMode") {
                    $isWrapped = $true
                    break
                }
            }
            
            if (-not $isWrapped) {
                $unwrappedCount++
                Write-Host "$($file.Name):$($i+1) - $($lines[$i].Trim())"
            }
        }
    }
}

Write-Host "`n=========================================="
Write-Host "TOTAL File.AppendAllText/WriteAllText calls: $totalCount"
Write-Host "UNWRAPPED (no deployment mode check): $unwrappedCount"
Write-Host "WRAPPED (with deployment mode check): $($totalCount - $unwrappedCount)"
Write-Host "=========================================="


