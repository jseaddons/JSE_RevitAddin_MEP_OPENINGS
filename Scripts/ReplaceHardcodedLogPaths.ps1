# PowerShell script to find all hardcoded log paths in C# files
# Run this to see all files that need updating

$hardcodedPathPattern = 'C:\\JSE_CSharp_Projects\\JSE_MEPOPENING_23\\Log'
$files = Get-ChildItem -Path "Services","Views","Commands" -Filter "*.cs" -Recurse | 
    Where-Object { 
        (Get-Content $_.FullName -Raw) -match $hardcodedPathPattern 
    }

Write-Host "Files with hardcoded log paths:"
$files | ForEach-Object { 
    $matches = (Select-String -Path $_.FullName -Pattern $hardcodedPathPattern -AllMatches).Matches
    Write-Host "  $($_.FullName) - $($matches.Count) occurrences"
}

Write-Host "`nTotal files to fix: $($files.Count)"

