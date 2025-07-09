# PowerShell script to comment out all DebugLogger.Log statements in a C# file

$inputFile = "DuctSleeveCommand.cs"
$outputFile = "DuctSleeveCommand.cs"  # Overwrite in place; change if you want a backup

(Get-Content $inputFile) | ForEach-Object {
    if ($_ -match "DebugLogger\.Log") {
        # If not already commented, add //
        if ($_ -notmatch "^\s*//") {
            "// $_"
        } else {
            $_
        }
    } else {
        $_
    }
} | Set-Content $outputFile