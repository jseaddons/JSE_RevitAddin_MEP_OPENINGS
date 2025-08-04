# PowerShell script to create an installer for the JSE_RevitAddin_MEP_OPENINGS add-in
# This script copies the add-in DLL and .addin manifest to both %APPDATA% and %PROGRAMDATA% for all users

param(
    [string]$BuildConfig = "Debug R24"
)

$ErrorActionPreference = 'Stop'

# Paths
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$binDir = Join-Path $projectRoot "bin/$BuildConfig"
$addinName = "JSE_RevitAddin_MEP_OPENINGS"
$dllName = "$addinName.dll"
$addinManifest = "$addinName.addin"

# Revit Addins folders
$appDataAddins = Join-Path $env:APPDATA "Autodesk/Revit/Addins/2024"
$programDataAddins = Join-Path $env:PROGRAMDATA "Autodesk/Revit/Addins/2024"

# Ensure output folders exist
Write-Host "Ensuring add-in folders exist..."
New-Item -ItemType Directory -Force -Path $appDataAddins | Out-Null
New-Item -ItemType Directory -Force -Path $programDataAddins | Out-Null

# Copy DLL and .addin manifest to both locations
Write-Host "Copying add-in files to user AppData and ProgramData..."
Copy-Item -Path (Join-Path $binDir $dllName) -Destination $appDataAddins -Force
Copy-Item -Path (Join-Path $binDir $addinManifest) -Destination $appDataAddins -Force
Copy-Item -Path (Join-Path $binDir $dllName) -Destination $programDataAddins -Force
Copy-Item -Path (Join-Path $binDir $addinManifest) -Destination $programDataAddins -Force

Write-Host "JSE_RevitAddin_MEP_OPENINGS add-in installed to:"
Write-Host "  $appDataAddins"
Write-Host "  $programDataAddins"
Write-Host "Installation complete."
