$dbPath = "C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\bin\Debug R24\ClashZones.db"

# Find correct SQLite DLL path
$possiblePaths = @(
    "C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\packages\Stub.System.Data.SQLite.Core.NetFramework.1.0.118\build\net46\System.Data.SQLite.dll",
    "C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\packages\System.Data.SQLite.Core.1.0.118\lib\net46\System.Data.SQLite.dll"
)

$dllPath = $null
foreach ($path in $possiblePaths) {
    if (Test-Path $path) {
        $dllPath = $path
        break
    }
}

if (-not $dllPath) {
    Write-Host "ERROR: Cannot find System.Data.SQLite.dll"
    exit 1
}

Add-Type -Path $dllPath

$conn = New-Object System.Data.SQLite.SQLiteConnection("Data Source=$dbPath")
$conn.Open()

$cmd = $conn.CreateCommand()
$cmd.CommandText = @"
SELECT 
    SleeveInstanceId,
    SleevePlacementActiveY,
    SleeveHeight,
    SleeveCorner1Y,
    SleeveCorner2Y,
    SleeveCorner3Y,
    SleeveCorner4Y,
    (SleeveCorner1Y + SleeveCorner3Y) / 2.0 as CalculatedCenterY
FROM ClashZones 
WHERE SleeveInstanceId IN (1106385, 1106392)
"@

$reader = $cmd.ExecuteReader()

while ($reader.Read()) {
    Write-Host "===== Sleeve $($reader['SleeveInstanceId']) ====="
    Write-Host "  SleevePlacementActiveY: $($reader['SleevePlacementActiveY']) ft"
    Write-Host "  SleeveHeight: $($reader['SleeveHeight']) ft"
    Write-Host "  SleeveCorner1Y (bottom): $($reader['SleeveCorner1Y']) ft"
    Write-Host "  SleeveCorner2Y (bottom): $($reader['SleeveCorner2Y']) ft"
    Write-Host "  SleeveCorner3Y (top): $($reader['SleeveCorner3Y']) ft"
    Write-Host "  SleeveCorner4Y (top): $($reader['SleeveCorner4Y']) ft"
    Write-Host "  CalculatedCenterY: $($reader['CalculatedCenterY']) ft"
    Write-Host ""
}

$reader.Close()
$conn.Close()
