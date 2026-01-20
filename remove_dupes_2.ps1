
$path = "c:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Data\Repositories\ClashZoneRepository.cs"
$lines = Get-Content -Path $path
$count = $lines.Count
Write-Host "Total lines: $count"

# Lines 7204-7419 contain the duplicate definitions.
# Index 7203 corresponds to Line 7204.
# Index 7418 corresponds to Line 7419.

# We keep 0..7202 (Lines 1..7203)
# We keep 7419..End (Lines 7420..End)

$newContent = $lines[0..7202] + $lines[7419..($count-1)]
$newContent | Set-Content -Path $path -Encoding UTF8
Write-Host "Removed duplicate lines 7204-7419"
