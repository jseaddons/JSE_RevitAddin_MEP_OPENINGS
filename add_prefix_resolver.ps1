# Add initialization line to MarkParameterService constructor
$filePath = "Services\MarkParameterService.cs"
$lines = Get-Content $filePath
$output = @()

foreach ($line in $lines) {
    $output += $line
    if ($line -match '^\s+_combinedService = new CombinedSleeveMarkService\(\);') {
        $output += '            _prefixResolver = new ClusterPrefixResolutionService(_logger);'
    }
}

$output | Set-Content $filePath -Encoding UTF8
Write-Host "✅ Added _prefixResolver initialization"
