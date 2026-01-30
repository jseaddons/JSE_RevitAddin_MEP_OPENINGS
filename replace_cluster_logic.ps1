# Simple line-by-line replacement for cluster prefix logic
$filePath = "Services\MarkParameterService.cs"
$lines = Get-Content $filePath
$output = @()
$i = 0
$skipUntil = -1

while ($i -lt $lines.Count) {
    $line = $lines[$i]
    
    # Detect the start of the cluster rule block
    if ($line -match 'CLUSTER RULE: For clusters, ALWAYS use discipline prefix only') {
        # Replace the entire block with new SOLID implementation
        $output += '                    // ✅ SOLID: Use dedicated service for cluster/combined prefix resolution (implements 3 Prefix Fixes)'
        $output += '                    string prefix;'
        $output += '                    bool isCluster = primaryZone.ClusterInstanceId > 0;'
        $output += '                    '
        $output += '                    if (isCluster)'
        $output += '                    {'
        $output += '                        // Get the sleeve element to check if it''s combined'
        $output += '                        var sleeveElement = doc.GetElement(new ElementId(instanceId)) as FamilyInstance;'
        $output += '                        var settings = markPrefixes ?? new MarkPrefixSettings();'
        $output += '                        '
        $output += '                        // Use the SOLID service for sophisticated prefix resolution'
        $output += '                        using var context = new SleeveDbContext(doc);'
        $output += '                        prefix = _prefixResolver.ResolvePrefixForClusterOrCombined('
        $output += '                            sleeveElement, category, settings.GetDisciplinePrefix(category), settings, doc, context);'
        $output += '                        '
        $output += '                        NumberingDebugLogger.LogInfo($"[MarkParameterService] Cluster {instanceId}: Using resolved prefix ''{prefix}'' (via ClusterPrefixResolutionService)");'
        $output += '                    }'
        $output += '                    else // Individual sleeve'
        $output += '                    {'
        $output += '                        prefix = prefixStrategy.ResolvePrefix(category, markPrefixes ?? new MarkPrefixSettings(), primaryZone);'
        $output += '                        NumberingDebugLogger.LogInfo($"[MarkParameterService] Individual Sleeve {instanceId}: Using resolved prefix ''{prefix}''");'
        $output += '                    }'
        
        # Skip the old implementation lines (find the closing brace of the else block)
        $i++
        $braceCount = 0
        $inBlock = $false
        while ($i -lt $lines.Count) {
            if ($lines[$i] -match 'if \(isCluster\)') {
                $inBlock = $true
            }
            if ($inBlock) {
                if ($lines[$i] -match '\{') { $braceCount++ }
                if ($lines[$i] -match '\}') { 
                    $braceCount--
                    if ($braceCount -eq 0) {
                        $i++
                        break
                    }
                }
            }
            $i++
        }
        continue
    }
    
    $output += $line
    $i++
}

$output | Set-Content $filePath -Encoding UTF8
Write-Host "✅ Replaced cluster prefix logic with SOLID service call"
