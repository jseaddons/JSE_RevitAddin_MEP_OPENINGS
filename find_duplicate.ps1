# Find Duplicate Cluster Placement Script
# This script searches for all cluster placement calls in your codebase
# and identifies potential duplicates

# Navigate to project directory
Set-Location "C:\JSE_CSharp_Projects\JSE_MEPOPENING_23"

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host " SEARCHING FOR CLUSTER PLACEMENT CALLS..." -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host ""

# Initialize results array
$results = @()

# Search all C# files EXCEPT in _BACKUPS and BACKUP folders
Get-ChildItem -Recurse -Include *.cs -ErrorAction SilentlyContinue | 
    Where-Object { $_.FullName -notmatch '\\BACKUP\\|\\BACKUPS\\|_BACKUP' } |
    ForEach-Object {
    $file = $_.FullName
    $relativePath = $file.Replace("C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\", "")
    
    try {
        # Search for ClusterSleeves calls (OLD method)
        $matches = Select-String -Path $file -Pattern "\.ClusterSleeves\(" -ErrorAction SilentlyContinue
        if ($matches) {
            foreach ($match in $matches) {
                $results += [PSCustomObject]@{
                    File = $relativePath
                    Line = $match.LineNumber
                    Method = "ClusterSleeves (OLD)"
                    Context = $match.Line.Trim()
                }
            }
        }
        
        # Search for ClusterSleevesV2 calls (NEW method)
        $matches = Select-String -Path $file -Pattern "\.ClusterSleevesV2\(" -ErrorAction SilentlyContinue
        if ($matches) {
            foreach ($match in $matches) {
                $results += [PSCustomObject]@{
                    File = $relativePath
                    Line = $match.LineNumber
                    Method = "ClusterSleevesV2 (NEW)"
                    Context = $match.Line.Trim()
                }
            }
        }
        
        # Search for PlaceClusterSleeve calls (PLACEMENT SERVICE)
        $matches = Select-String -Path $file -Pattern "PlaceClusterSleeve\(" -ErrorAction SilentlyContinue
        if ($matches) {
            foreach ($match in $matches) {
                $results += [PSCustomObject]@{
                    File = $relativePath
                    Line = $match.LineNumber
                    Method = "PlaceClusterSleeve (PLACEMENT)"
                    Context = $match.Line.Trim()
                }
            }
        }
    }
    catch {
        # Ignore errors reading individual files
    }
}

# Display results
if ($results.Count -gt 0) {
    Write-Host "FOUND $($results.Count) CLUSTER PLACEMENT CALLS:" -ForegroundColor Yellow
    Write-Host ""
    $results | Format-Table -Property File, Line, Method -AutoSize
    
    Write-Host "`nDETAILED VIEW WITH CODE CONTEXT:" -ForegroundColor Cyan
    Write-Host "================================================================================"
    foreach ($result in $results) {
        Write-Host "`nFile: $($result.File)" -ForegroundColor Green
        Write-Host "Line: $($result.Line)" -ForegroundColor Green
        Write-Host "Method: $($result.Method)" -ForegroundColor Yellow
        Write-Host "Code: $($result.Context)" -ForegroundColor White
        Write-Host "--------------------------------------------------------------------------------"
    }
}
else {
    Write-Host "NO CLUSTER PLACEMENT CALLS FOUND!" -ForegroundColor Red
    Write-Host "This might indicate a search issue or the methods have different names."
}

Write-Host "`n================================================================================" -ForegroundColor Cyan
Write-Host " ANALYSIS" -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan

$clusterSleevesCount = ($results | Where-Object { $_.Method -eq "ClusterSleeves (OLD)" }).Count
$clusterSleevesV2Count = ($results | Where-Object { $_.Method -eq "ClusterSleevesV2 (NEW)" }).Count
$placementCount = ($results | Where-Object { $_.Method -eq "PlaceClusterSleeve (PLACEMENT)" }).Count

Write-Host "`nClusterSleeves (OLD) calls: $clusterSleevesCount" -ForegroundColor $(if ($clusterSleevesCount -gt 0) { "Yellow" } else { "Green" })
Write-Host "ClusterSleevesV2 (NEW) calls: $clusterSleevesV2Count" -ForegroundColor $(if ($clusterSleevesV2Count -gt 0) { "Yellow" } else { "Green" })
Write-Host "PlaceClusterSleeve (PLACEMENT) calls: $placementCount" -ForegroundColor $(if ($placementCount -gt 2) { "Red" } else { "Green" })

Write-Host ""

# Analyze and provide recommendations
if ($placementCount -gt 2) {
    Write-Host "WARNING: $placementCount calls to PlaceClusterSleeve found!" -ForegroundColor Red
    Write-Host "   Expected: 1-2 calls (one for definition, maybe one for actual placement)" -ForegroundColor Red
    Write-Host "   Found: $placementCount calls - This suggests DUPLICATE PLACEMENT!" -ForegroundColor Red
    Write-Host ""
    Write-Host "   ACTION REQUIRED:" -ForegroundColor Yellow
    Write-Host "   1. Check the files listed above" -ForegroundColor White
    Write-Host "   2. Find calls that are NOT in loops or are duplicates" -ForegroundColor White
    Write-Host "   3. Comment out the redundant call(s)" -ForegroundColor White
}
elseif ($clusterSleevesCount -gt 0 -and $clusterSleevesV2Count -gt 0) {
    Write-Host "WARNING: BOTH OLD and NEW cluster methods are being called!" -ForegroundColor Yellow
    Write-Host "   This might cause duplicate placement." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "   RECOMMENDATION:" -ForegroundColor Yellow
    Write-Host "   - Keep ClusterSleevesV2 (NEW) method" -ForegroundColor White
    Write-Host "   - Comment out ClusterSleeves (OLD) method" -ForegroundColor White
}
else {
    Write-Host "No obvious duplicate method calls detected." -ForegroundColor Green
    Write-Host "  The duplicate might be within a method (loop running twice)" -ForegroundColor White
}

Write-Host "`n================================================================================" -ForegroundColor Cyan
Write-Host " NEXT STEPS" -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host "1. Review the files and line numbers listed above" -ForegroundColor White
Write-Host "2. Open the identified files in Visual Studio" -ForegroundColor White
Write-Host "3. Navigate to the line numbers shown" -ForegroundColor White
Write-Host "4. Comment out duplicate/redundant calls" -ForegroundColor White
Write-Host "5. Rebuild solution (Ctrl+Shift+B)" -ForegroundColor White
Write-Host "6. Test in Revit and verify 2.8x performance improvement" -ForegroundColor White
Write-Host ""

# Pause to keep window open
Write-Host "Press any key to exit..." -ForegroundColor Cyan
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
