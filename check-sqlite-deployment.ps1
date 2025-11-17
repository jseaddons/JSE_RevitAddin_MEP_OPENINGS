# SQLite Native DLL Deployment Checker
# Checks if e_sqlite3.dll is in the correct location for Revit add-in

param(
    [string]$RevitVersion = "2023"
)

Write-Host "🔍 Checking SQLite Native DLL Deployment..." -ForegroundColor Cyan
Write-Host ""

# Check common Revit Addins locations
$addinLocations = @(
    "$env:APPDATA\Autodesk\Revit\Addins\$RevitVersion",
    "C:\ProgramData\Autodesk\Revit\Addins\$RevitVersion"
)

$foundDll = $false
$dllLocation = ""

foreach ($location in $addinLocations) {
    $dllPath = Join-Path $location "e_sqlite3.dll"
    $mainDllPath = Join-Path $location "JSE_RevitAddin_MEP_OPENINGS.dll"
    
    Write-Host "Checking: $location" -ForegroundColor Gray
    
    if (Test-Path $mainDllPath) {
        Write-Host "  ✅ Main DLL found: JSE_RevitAddin_MEP_OPENINGS.dll" -ForegroundColor Green
        
        if (Test-Path $dllPath) {
            Write-Host "  ✅ Native DLL found: e_sqlite3.dll" -ForegroundColor Green
            $foundDll = $true
            $dllLocation = $dllPath
            
            # Check file size
            $fileInfo = Get-Item $dllPath
            Write-Host "  📦 File size: $($fileInfo.Length) bytes" -ForegroundColor Gray
        } else {
            Write-Host "  ❌ Native DLL MISSING: e_sqlite3.dll" -ForegroundColor Red
            Write-Host "     Expected at: $dllPath" -ForegroundColor Yellow
        }
        
        # List all files in directory
        Write-Host ""
        Write-Host "  Files in directory:" -ForegroundColor Gray
        Get-ChildItem $location | ForEach-Object {
            $icon = if ($_.Name -eq "e_sqlite3.dll") { "🔑" } elseif ($_.Name -like "*.dll") { "📦" } else { "📄" }
            Write-Host "    $icon $($_.Name)" -ForegroundColor Gray
        }
    } else {
        Write-Host "  ⚠️  Main DLL not found (add-in not deployed here)" -ForegroundColor Yellow
    }
    Write-Host ""
}

# Check build output
$buildOutputs = @(
    ".\bin\Release R$RevitVersion",
    ".\bin\Debug R$RevitVersion"
)

Write-Host "Checking build output directories..." -ForegroundColor Cyan
foreach ($output in $buildOutputs) {
    if (Test-Path $output) {
        $dllPath = Join-Path $output "e_sqlite3.dll"
        Write-Host "  $output" -ForegroundColor Gray
        if (Test-Path $dllPath) {
            Write-Host "    ✅ Native DLL found in build output: e_sqlite3.dll" -ForegroundColor Green
        } else {
            Write-Host "    ❌ Native DLL NOT in build output" -ForegroundColor Red
        }
    }
}

Write-Host ""

if ($foundDll) {
    Write-Host "✅ SQLite native DLL is properly deployed!" -ForegroundColor Green
    Write-Host "   Location: $dllLocation" -ForegroundColor Gray
} else {
    Write-Host "❌ SQLite native DLL is MISSING from Revit Addins folder!" -ForegroundColor Red
    Write-Host ""
    Write-Host "🔧 SOLUTION:" -ForegroundColor Yellow
    Write-Host "   1. Build the project" -ForegroundColor White
    Write-Host "   2. Copy e_sqlite3.dll from build output to Revit Addins folder" -ForegroundColor White
    Write-Host "   3. Or run: .\deploy-addin.ps1 -RevitVersion $RevitVersion" -ForegroundColor White
    Write-Host ""
    Write-Host "   Build output locations to check:" -ForegroundColor Gray
    foreach ($output in $buildOutputs) {
        if (Test-Path $output) {
            Write-Host "     - $output" -ForegroundColor Gray
        }
    }
}

