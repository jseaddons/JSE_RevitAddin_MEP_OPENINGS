# Deploy R24 Build to Revit 2024
# Close Revit 2024 before running this script!

$source = "bin\Debug R24\JSE_RevitAddin_MEP_OPENINGS.dll"
$destination = "$env:APPDATA\Autodesk\Revit\Addins\2024\"

Write-Host "🔥 Deploying R24 Build with Critical Unit Conversion Fix..." -ForegroundColor Cyan
Write-Host ""

# Check if Revit process is running
$revitProcess = Get-Process -Name "Revit" -ErrorAction SilentlyContinue
if ($revitProcess) {
    Write-Host "❌ ERROR: Revit is still running!" -ForegroundColor Red
    Write-Host "   Please close Revit 2024 completely and run this script again." -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

# Check source file exists
if (-not (Test-Path $source)) {
    Write-Host "❌ ERROR: Source DLL not found: $source" -ForegroundColor Red
    Write-Host "   Please build the project first." -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

# Show what we're copying
$sourceInfo = Get-Item $source
Write-Host "Source DLL:" -ForegroundColor Green
Write-Host "  Path: $source"
Write-Host "  Size: $($sourceInfo.Length) bytes"
Write-Host "  Date: $($sourceInfo.LastWriteTime)"
Write-Host ""

# Check current deployed version
$destPath = Join-Path $destination "JSE_RevitAddin_MEP_OPENINGS.dll"
if (Test-Path $destPath) {
    $destInfo = Get-Item $destPath
    Write-Host "Current Deployed DLL:" -ForegroundColor Yellow
    Write-Host "  Path: $destPath"
    Write-Host "  Size: $($destInfo.Length) bytes"
    Write-Host "  Date: $($destInfo.LastWriteTime)"
    Write-Host ""
}

# Copy the file
try {
    Copy-Item $source -Destination $destination -Force
    Write-Host "✅ Successfully deployed R24 build!" -ForegroundColor Green
    Write-Host ""
    
    # Verify the copy
    $newInfo = Get-Item $destPath
    Write-Host "New Deployed DLL:" -ForegroundColor Cyan
    Write-Host "  Path: $destPath"
    Write-Host "  Size: $($newInfo.Length) bytes"
    Write-Host "  Date: $($newInfo.LastWriteTime)"
    Write-Host ""
    Write-Host "🚀 You can now start Revit 2024 and test intersection detection!" -ForegroundColor Green
}
catch {
    Write-Host "❌ ERROR: Failed to copy DLL" -ForegroundColor Red
    Write-Host "   $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host ""
    exit 1
}
