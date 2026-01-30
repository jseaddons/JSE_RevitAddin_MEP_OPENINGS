# Update cluster prefix resolution to use the new SOLID service
$filePath = "Services\MarkParameterService.cs"
$content = Get-Content $filePath -Raw

# Replace the simple cluster rule with a call to the new service
$oldPattern = @'
                    // Γ£à CLUSTER RULE: For clusters, ALWAYS use discipline prefix only \(ignore system type overrides\)
                    // For individual sleeves, use full prefix resolution \(including system type overrides\)
                    // Γ£à FIX: Check ClusterInstanceId > 0 to detect clusters \(each cluster has unique ID, so Count is always 1\)
                    string prefix;
                    bool isCluster = primaryZone\.ClusterInstanceId > 0;
                    if \(isCluster\)
                    \{
                        var settings = markPrefixes \?\? new MarkPrefixSettings\(\);
                        prefix = settings\.GetDisciplinePrefix\(category\);
\s+NumberingDebugLogger\.LogInfo\(\$"\[MarkParameterService\] Cluster \{instanceId\}: Using discipline prefix '\{prefix\}' \(ClusterInstanceId=\{primaryZone\.ClusterInstanceId\}\)"\);
                    \}
                    else // Individual sleeve
                    \{
                        prefix = prefixStrategy\.ResolvePrefix\(category, markPrefixes \?\? new MarkPrefixSettings\(\), primaryZone\);
\s+NumberingDebugLogger\.LogInfo\(\$"\[MarkParameterService\] Individual Sleeve \{instanceId\}: Using resolved prefix '\{prefix\}'"\);
                    \}
'@

$newCode = @'
                    // ✅ SOLID: Use dedicated service for cluster/combined prefix resolution (implements 3 Prefix Fixes)
                    string prefix;
                    bool isCluster = primaryZone.ClusterInstanceId > 0;
                    
                    if (isCluster)
                    {
                        // Get the sleeve element to check if it's combined
                        var sleeveElement = doc.GetElement(new ElementId(instanceId)) as FamilyInstance;
                        var settings = markPrefixes ?? new MarkPrefixSettings();
                        
                        // Use the SOLID service for sophisticated prefix resolution
                        using var context = new SleeveDbContext(doc);
                        prefix = _prefixResolver.ResolvePrefixForClusterOrCombined(
                            sleeveElement, category, settings.GetDisciplinePrefix(category), settings, doc, context);
                        
                        NumberingDebugLogger.LogInfo($"[MarkParameterService] Cluster {instanceId}: Using resolved prefix '{prefix}' (via ClusterPrefixResolutionService)");
                    }
                    else // Individual sleeve
                    {
                        prefix = prefixStrategy.ResolvePrefix(category, markPrefixes ?? new MarkPrefixSettings(), primaryZone);
                        NumberingDebugLogger.LogInfo($"[MarkParameterService] Individual Sleeve {instanceId}: Using resolved prefix '{prefix}'");
                    }
'@

if ($content -match [regex]::Escape($oldPattern.Substring(0, 50))) {
    $content = $content -replace $oldPattern, $newCode
    Set-Content $filePath -Value $content -NoNewline -Encoding UTF8
    Write-Host "✅ Updated cluster prefix resolution to use SOLID service"
} else {
    Write-Host "⚠️ Pattern not found - manual update required"
}
