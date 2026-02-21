$file = Get-ChildItem 'C:\Users\jse2084\.claude\projects\c--Jse-Developments-JSE-MEPOPENING-23\c4f1fa93-a50e-4889-92c1-53d18884eae9\tool-results\toolu_015Q3qVJbH6FzriJu4uz1izM.txt'
$lines = Get-Content $file.FullName

# Critical warnings that can cause runtime issues
$criticalCodes = @('CS0162', 'CS0168', 'CS0219', 'CS0414', 'CS0618', 'CS0105')

Write-Host "=== CRITICAL / NOTABLE WARNINGS ==="
Write-Host ""

# CS0162 - Unreachable code
Write-Host "--- CS0162: Unreachable code detected ---"
foreach ($line in $lines) {
    if ($line -match 'warning CS0162') {
        $short = $line -replace '.*\\([^\\]+\.cs)', '$1'
        Write-Host "  $short"
    }
}

Write-Host ""
Write-Host "--- CS0618: Obsolete member usage ---"
foreach ($line in $lines) {
    if ($line -match 'warning CS0618') {
        $short = $line -replace '.*\\([^\\]+\.cs)', '$1'
        Write-Host "  $short"
    }
}

Write-Host ""
Write-Host "--- CS0414: Field assigned but never used ---"
foreach ($line in $lines) {
    if ($line -match 'warning CS0414') {
        $short = $line -replace '.*\\([^\\]+\.cs)', '$1'
        Write-Host "  $short"
    }
}

Write-Host ""
Write-Host "--- CS0168: Variable declared but never used ---"
$cs168 = @()
foreach ($line in $lines) {
    if ($line -match 'warning CS0168') {
        $short = $line -replace '.*\\([^\\]+\.cs)', '$1'
        $cs168 += $short
    }
}
$cs168 | Sort-Object -Unique | ForEach-Object { Write-Host "  $_" }

Write-Host ""
Write-Host "--- CS0219: Variable assigned but never used ---"
$cs219 = @()
foreach ($line in $lines) {
    if ($line -match 'warning CS0219') {
        $short = $line -replace '.*\\([^\\]+\.cs)', '$1'
        $cs219 += $short
    }
}
$cs219 | Sort-Object -Unique | ForEach-Object { Write-Host "  $_" }

Write-Host ""
Write-Host "--- CS0105: Duplicate using directives ---"
foreach ($line in $lines) {
    if ($line -match 'warning CS0105') {
        $short = $line -replace '.*\\([^\\]+\.cs)', '$1'
        Write-Host "  $short"
    }
}
