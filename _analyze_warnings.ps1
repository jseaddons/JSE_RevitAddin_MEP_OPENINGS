$file = Get-ChildItem 'C:\Users\jse2084\.claude\projects\c--Jse-Developments-JSE-MEPOPENING-23\c4f1fa93-a50e-4889-92c1-53d18884eae9\tool-results\toolu_015Q3qVJbH6FzriJu4uz1izM.txt'
$lines = Get-Content $file.FullName
$warnCodes = @()
foreach ($line in $lines) {
    if ($line -match 'warning (CS\d+)') {
        $warnCodes += $Matches[1]
    }
}
Write-Host "Total warnings: $($warnCodes.Count)"
Write-Host ""
$warnCodes | Group-Object | Sort-Object Count -Descending | ForEach-Object {
    Write-Host ("{0,5} x {1}" -f $_.Count, $_.Name)
}
