# PowerShell script to comment out all DamperLogger.Log statements in a C# file

$inputFile = "FireDamperPlaceCommand.cs"
$outputFile = "FireDamperPlaceCommand.cs"  # Overwrite in place; change if you want a backup

(Get-Content $inputFile) | ForEach-Object {
    if ($_ -match "DamperLogger\.Log") {
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
