# JSE MEP Openings Add-in Deployment Script
# Deploys the add-in and all SQLite dependencies to Revit Addins folder
# Usage: .\deploy-addin.ps1 [-RevitVersion "2023"]

param(
    [string]$RevitVersion = "2023",
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

# Determine source path based on configuration
$sourcePath = ".\bin\$Configuration R$RevitVersion"
if (-not (Test-Path $sourcePath)) {
    Write-Host "❌ Build output not found at: $sourcePath" -ForegroundColor Red
    Write-Host "Please build the project first in $Configuration R$RevitVersion configuration" -ForegroundColor Yellow
    exit 1
}

# Target path in Revit Addins folder
$targetPath = "$env:APPDATA\Autodesk\Revit\Addins\$RevitVersion\JSE_MEP_Openings"

Write-Host "🚀 Deploying JSE MEP Openings Add-in..." -ForegroundColor Cyan
Write-Host "   Source: $sourcePath" -ForegroundColor Gray
Write-Host "   Target: $targetPath" -ForegroundColor Gray
Write-Host "   Revit Version: $RevitVersion" -ForegroundColor Gray
Write-Host ""

# Create target directory
if (-not (Test-Path $targetPath)) {
    New-Item -ItemType Directory -Force -Path $targetPath | Out-Null
    Write-Host "✅ Created target directory: $targetPath" -ForegroundColor Green
}

# Copy main add-in DLL
$mainDll = "$sourcePath\JSE_RevitAddin_MEP_OPENINGS.dll"
if (Test-Path $mainDll) {
    Copy-Item -Path $mainDll -Destination $targetPath -Force
    Write-Host "✅ Copied main DLL: JSE_RevitAddin_MEP_OPENINGS.dll" -ForegroundColor Green
} else {
    Write-Host "❌ Main DLL not found: $mainDll" -ForegroundColor Red
    exit 1
}

# Copy .addin manifest
$addinManifest = ".\JSE_RevitAddin_MEP_OPENINGS.addin"
if (Test-Path $addinManifest) {
    Copy-Item -Path $addinManifest -Destination $targetPath -Force
    Write-Host "✅ Copied .addin manifest" -ForegroundColor Green
} else {
    Write-Host "⚠️  .addin manifest not found: $addinManifest" -ForegroundColor Yellow
}

# Copy SQLite dependencies (critical for SQLite functionality)
$sqliteFiles = @(
    "System.Data.SQLite.dll",
    "SQLite.Interop.dll"
)

$allSqliteFound = $true
foreach ($file in $sqliteFiles) {
    $sourceFile = "$sourcePath\$file"
    if (Test-Path $sourceFile) {
        Copy-Item -Path $sourceFile -Destination $targetPath -Force
        Write-Host "✅ Copied SQLite dependency: $file" -ForegroundColor Green
    } else {
        Write-Host "⚠️  SQLite dependency not found: $file" -ForegroundColor Yellow
        $allSqliteFound = $false
    }
}

if (-not $allSqliteFound) {
    Write-Host ""
    Write-Host "⚠️  WARNING: Some SQLite dependencies are missing!" -ForegroundColor Yellow
    Write-Host "   SQLite functionality may not work correctly." -ForegroundColor Yellow
    Write-Host "   Ensure System.Data.SQLite.Core NuGet package is restored." -ForegroundColor Yellow
}

# Copy native x64 runtime if present
$sqliteInteropX64 = Join-Path $sourcePath "x64\SQLite.Interop.dll"
if (Test-Path $sqliteInteropX64) {
    $targetRuntimeDir = Join-Path $targetPath "x64"
    if (-not (Test-Path $targetRuntimeDir)) {
        New-Item -ItemType Directory -Path $targetRuntimeDir | Out-Null
    }
    Copy-Item -Path $sqliteInteropX64 -Destination $targetRuntimeDir -Force
    Write-Host "✅ Copied SQLite native runtime: x64\SQLite.Interop.dll" -ForegroundColor Green
} else {
    Write-Host "⚠️  SQLite native runtime not found: x64\SQLite.Interop.dll" -ForegroundColor Yellow
}

# Copy any other dependencies that might be needed
$otherDeps = @(
    "CommunityToolkit.Mvvm.dll",
    "Serilog.dll",
    "Serilog.Sinks.Debug.dll"
)

foreach ($file in $otherDeps) {
    $sourceFile = "$sourcePath\$file"
    if (Test-Path $sourceFile) {
        Copy-Item -Path $sourceFile -Destination $targetPath -Force
        Write-Host "✅ Copied dependency: $file" -ForegroundColor Gray
    }
}

Write-Host ""
Write-Host "✅ Deployment complete!" -ForegroundColor Green
Write-Host "   Add-in location: $targetPath" -ForegroundColor Gray
Write-Host ""
Write-Host "📋 Next steps:" -ForegroundColor Cyan
Write-Host "   1. Restart Revit to load the add-in" -ForegroundColor White
Write-Host "   2. Check Revit Add-Ins tab for 'JSE MEP Openings' ribbon" -ForegroundColor White
Write-Host "   3. If SQLite errors occur, verify all DLLs are in: $targetPath" -ForegroundColor White

