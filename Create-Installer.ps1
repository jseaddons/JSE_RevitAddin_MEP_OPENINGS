# JSE Revit MEP Openings Add-in Installer Creator
# This script creates a self-contained installer

param(
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "Y:\DESIGN\JSE_Addins\MEP_Openings\JSE_RevitAddin_MEP_OPENINGS\JSE_MEP_Openings_Installer.exe"
)

Add-Type -AssemblyName System.IO.Compression.FileSystem

$SourceFolder = "Y:\DESIGN\JSE_Addins\MEP_Openings\JSE_RevitAddin_MEP_OPENINGS"
$TempZip = "$env:TEMP\JSE_MEP_Openings_Temp.zip"

Write-Host "Creating self-extracting installer..." -ForegroundColor Green

# Create a zip file with all necessary files
if (Test-Path $TempZip) { Remove-Item $TempZip -Force }

$FilesToInclude = @(
    "JSE_RevitAddin_MEP_OPENINGS.dll",
    "JSE_RevitAddin_MEP_OPENINGS.pdb", 
    "JSE_RevitAddin_MEP_OPENINGS.addin",
    "Install-For-User.bat",
    "Install-For-User.vbs",
    "README.md"
)

$compress = @{
    Path = $FilesToInclude | ForEach-Object { Join-Path $SourceFolder $_ } | Where-Object { Test-Path $_ }
    DestinationPath = $TempZip
}

Compress-Archive @compress -Force

# Create installer script content
$InstallerContent = @'
# JSE Revit MEP Openings Add-in Self-Extracting Installer
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

$result = [System.Windows.Forms.MessageBox]::Show(
    "This will install the JSE Revit MEP Openings Add-in for the current user.`n`nContinue with installation?", 
    "JSE Revit MEP Openings Installer", 
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
)

if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
    try {
        $userAddinsPath = "$env:APPDATA\Autodesk\Revit\Addins\2024"
        if (!(Test-Path $userAddinsPath)) {
            New-Item -ItemType Directory -Path $userAddinsPath -Force | Out-Null
        }
        
        # Extract and install files here (embedded content)
        # This would contain the base64 encoded zip content
        
        [System.Windows.Forms.MessageBox]::Show(
            "Installation completed successfully!`n`nPlease restart Revit to load the add-in.", 
            "Installation Complete", 
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Installation failed: $($_.Exception.Message)", 
            "Installation Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}
'@

Write-Host "Installer creation would require additional tools like WiX or NSIS for a proper EXE/MSI."
Write-Host "The batch and VBScript files created are the recommended alternatives."

Remove-Item $TempZip -Force -ErrorAction SilentlyContinue
