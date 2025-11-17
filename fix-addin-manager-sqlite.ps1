# Quick Fix Script for Add-in Manager SQLite Native DLL Issue
# This script copies e_sqlite3.dll to your build output directory
# Run this after building if you're using Add-in Manager to test

param(
    [string]$Configuration = "Debug",
    [string]$RevitVersion = "24"
)

Write-Host "🔧 Fixing SQLite Native DLL for Add-in Manager..." -ForegroundColor Cyan
Write-Host ""

# Determine build output path
$buildOutput = ".\bin\$Configuration R$RevitVersion"

if (-not (Test-Path $buildOutput)) {
    Write-Host "❌ Build output directory not found: $buildOutput" -ForegroundColor Red
    Write-Host "Please build the project first or adjust Configuration/RevitVersion parameters" -ForegroundColor Yellow
    exit 1
}

Write-Host "📁 Build output directory: $buildOutput" -ForegroundColor Gray

# Check if main DLL exists
$mainDll = Join-Path $buildOutput "JSE_RevitAddin_MEP_OPENINGS.dll"
if (-not (Test-Path $mainDll)) {
    Write-Host "❌ Main DLL not found: $mainDll" -ForegroundColor Red
    Write-Host "Please build the project first" -ForegroundColor Yellow
    exit 1
}

Write-Host "✅ Main DLL found: JSE_RevitAddin_MEP_OPENINGS.dll" -ForegroundColor Green

# Check if native DLL already exists
$nativeDll = Join-Path $buildOutput "e_sqlite3.dll"
if (Test-Path $nativeDll) {
    Write-Host "✅ Native DLL already exists: e_sqlite3.dll" -ForegroundColor Green
    Write-Host "   Location: $nativeDll" -ForegroundColor Gray
    exit 0
}

Write-Host "⚠️  Native DLL not found. Searching for source..." -ForegroundColor Yellow

# Try to find DLL in other build outputs
$otherBuildOutputs = @(
    ".\bin\Release R$RevitVersion",
    ".\bin\Debug R$RevitVersion"
)

$sourceDll = $null
foreach ($output in $otherBuildOutputs) {
    $testPath = Join-Path $output "e_sqlite3.dll"
    if (Test-Path $testPath) {
        $sourceDll = $testPath
        Write-Host "✅ Found DLL in: $output" -ForegroundColor Green
        break
    }
}

# If not found, try NuGet cache or nested build output paths
if (-not $sourceDll) {
    # Check for nested NuGet package structure in build output (common issue)
    $nestedPaths = @(
        ".\bin\$Configuration R$RevitVersion\Release R$RevitVersion\2.1.6\runtimes\win-x64\native\e_sqlite3.dll",
        ".\bin\$Configuration R$RevitVersion\Release R$RevitVersion\2.1.7\runtimes\win-x64\native\e_sqlite3.dll",
        ".\bin\$Configuration R$RevitVersion\Debug R$RevitVersion\2.1.6\runtimes\win-x64\native\e_sqlite3.dll",
        ".\bin\$Configuration R$RevitVersion\Debug R$RevitVersion\2.1.7\runtimes\win-x64\native\e_sqlite3.dll"
    )
    
    foreach ($nestedPath in $nestedPaths) {
        if (Test-Path $nestedPath) {
            $sourceDll = $nestedPath
            Write-Host "✅ Found DLL in nested build output: $nestedPath" -ForegroundColor Green
            break
        }
    }
    
    # Try NuGet cache
    if (-not $sourceDll) {
        $nugetPaths = @(
            "$env:USERPROFILE\.nuget\packages\sqlitepclraw.bundle_e_sqlite3\2.1.7\runtimes\win-x64\native\e_sqlite3.dll",
            "$env:USERPROFILE\.nuget\packages\sqlitepclraw.bundle_e_sqlite3\2.1.6\runtimes\win-x64\native\e_sqlite3.dll"
        )
        
        foreach ($nugetPath in $nugetPaths) {
            if (Test-Path $nugetPath) {
                $sourceDll = $nugetPath
                Write-Host "✅ Found DLL in NuGet cache: $nugetPath" -ForegroundColor Green
                break
            }
        }
    }
}

if (-not $sourceDll) {
    Write-Host ""
    Write-Host "❌ Could not find e_sqlite3.dll!" -ForegroundColor Red
    Write-Host ""
    Write-Host "🔧 Solutions:" -ForegroundColor Yellow
    Write-Host "   1. Restore NuGet packages: dotnet restore" -ForegroundColor White
    Write-Host "   2. Rebuild the project" -ForegroundColor White
    Write-Host "   3. Manually download from: https://www.sqlite.org/download.html" -ForegroundColor White
    Write-Host ""
    Write-Host "   Expected NuGet location:" -ForegroundColor Gray
    Write-Host "   $env:USERPROFILE\.nuget\packages\sqlitepclraw.bundle_e_sqlite3\2.1.7\runtimes\win-x64\native\e_sqlite3.dll" -ForegroundColor Gray
    exit 1
}

# Copy the DLL
try {
    Copy-Item $sourceDll $nativeDll -Force
    Write-Host ""
    Write-Host "✅ Successfully copied e_sqlite3.dll!" -ForegroundColor Green
    Write-Host "   From: $sourceDll" -ForegroundColor Gray
    Write-Host "   To:   $nativeDll" -ForegroundColor Gray
    Write-Host ""
    Write-Host "✅ Both DLLs are now in the same directory:" -ForegroundColor Green
    Write-Host "   - JSE_RevitAddin_MEP_OPENINGS.dll" -ForegroundColor Gray
    Write-Host "   - e_sqlite3.dll" -ForegroundColor Gray
    Write-Host ""
    Write-Host "📋 Next steps:" -ForegroundColor Cyan
    Write-Host "   1. Restart Revit Add-in Manager" -ForegroundColor White
    Write-Host "   2. Reload the add-in DLL" -ForegroundColor White
    Write-Host "   3. Test SQLite functionality" -ForegroundColor White
} catch {
    Write-Host ""
    Write-Host "❌ Failed to copy DLL: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

