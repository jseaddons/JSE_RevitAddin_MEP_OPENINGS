# KillRevit.ps1
# PowerShell script to kill all running Autodesk Revit processes

$revitProcesses = Get-Process -Name 'Revit' -ErrorAction SilentlyContinue
if ($revitProcesses) {
    Write-Host "Killing Autodesk Revit processes..."
    $revitProcesses | ForEach-Object { $_.Kill() }
    Write-Host "All Revit processes have been terminated."
} else {
    Write-Host "No running Revit processes found."
}
