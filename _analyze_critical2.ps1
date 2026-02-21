$file = 'C:\Users\jse2084\.claude\projects\c--Jse-Developments-JSE-MEPOPENING-23\c4f1fa93-a50e-4889-92c1-53d18884eae9\tool-results\toolu_015Q3qVJbH6FzriJu4uz1izM.txt'
$lines = Get-Content $file

# Only show from real .csproj (not wpftmp duplicates)
$realLines = $lines | Where-Object { $_ -notmatch 'wpftmp' }

$criticals = @('CS0162','CS0618','CS0414','CS0168','CS0219')

foreach ($code in $criticals) {
    $matches_found = $realLines | Where-Object { $_ -match "warning $code" }
    if ($matches_found.Count -gt 0) {
        switch ($code) {
            'CS0162' { Write-Host "`n=== CS0162: UNREACHABLE CODE ($($matches_found.Count)) ===" }
            'CS0618' { Write-Host "`n=== CS0618: OBSOLETE API USAGE ($($matches_found.Count)) ===" }
            'CS0414' { Write-Host "`n=== CS0414: FIELD ASSIGNED NEVER USED ($($matches_found.Count)) ===" }
            'CS0168' { Write-Host "`n=== CS0168: VARIABLE DECLARED NEVER USED ($($matches_found.Count)) ===" }
            'CS0219' { Write-Host "`n=== CS0219: VARIABLE ASSIGNED NEVER USED ($($matches_found.Count)) ===" }
        }
        foreach ($m in $matches_found) {
            # Extract just filename(line): warning message
            if ($m -match '([^\\]+\.cs\(\d+,\d+\)).*?(warning CS\d+:.+?)\s*\[') {
                Write-Host "  $($Matches[1]) -> $($Matches[2])"
            }
        }
    }
}

# Summary
Write-Host "`n=== NULLABLE WARNINGS SUMMARY (not individually listed) ==="
$nullableCount = ($realLines | Where-Object { $_ -match 'warning CS86' }).Count
Write-Host "  CS86xx nullable warnings: $nullableCount (mostly cosmetic with Nullable=enable)"

$totalReal = ($realLines | Where-Object { $_ -match 'warning CS' }).Count
Write-Host "`nTotal real warnings (excl wpftmp duplicates): $totalReal"
