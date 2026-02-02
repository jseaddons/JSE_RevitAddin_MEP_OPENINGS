$path = "c:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Data\Repositories\ClashZoneRepository.cs"
$content = Get-Content $path -Raw
$target = 'ReadyForPlacementFlag = \(SELECT IsCurrentClashFlag FROM TempFlagUpdates WHERE TempFlagUpdates.ClashZoneId = ClashZones.ClashZoneId\),'
$replacement = 'ReadyForPlacementFlag = CASE WHEN (SELECT IsResolvedFlag FROM TempFlagUpdates WHERE TempFlagUpdates.ClashZoneId = ClashZones.ClashZoneId) = 1 OR (SELECT IsClusterResolvedFlag FROM TempFlagUpdates WHERE TempFlagUpdates.ClashZoneId = ClashZones.ClashZoneId) = 1 OR (SELECT IsCombinedResolved FROM TempFlagUpdates WHERE TempFlagUpdates.ClashZoneId = ClashZones.ClashZoneId) = 1 THEN 0 ELSE (SELECT IsCurrentClashFlag FROM TempFlagUpdates WHERE TempFlagUpdates.ClashZoneId = ClashZones.ClashZoneId) END,'
$newContent = $content -replace $target, $replacement
$newContent | Set-Content $path -NoNewline

$path2 = "c:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Services\OpeningCommandOrchestrator.cs"
$content2 = Get-Content $path2 -Raw

# Add logging
$targetLog = 'BulkPlacementResult bulkResult = null;'
$replacementLog = 'BulkPlacementResult bulkResult = null;
                                 SafeFileLogger.SafeAppendTextAlways("placement_debug.log", "`n[" + (Get-Date -Format "HH:mm:ss") + "] [BULK-PLACEMENT-START] allBulkTaskItems.Count = " + $allBulkTaskItems.Count + "`n");'
$content2 = $content2 -replace $targetLog, $replacementLog

# Add FlagManager call
$targetFlag = 'var placedZones = bulkResult.PlacedItems.Select\(p => p.Zone\).ToList\(\);\s+using \(var dbContext = new JSE_RevitAddin_MEP_OPENINGS.Data.SleeveDbContext\(_document\)\)'
$replacementFlag = 'var placedZones = bulkResult.PlacedItems.Select(p => p.Zone).ToList();
                                             
                                             // ✅ FIX: Update flags in database after bulk placement
                                             if (placedZones.Any())
                                             {
                                                 var flagUpdates = bulkResult.PlacedItems
                                                     .Select(item => (item.Zone.Id, item.ElementId.IntegerValue, false))
                                                     .ToList();
                                                 
                                                 flagManager.UpdateFlagsAfterPlacement(flagUpdates);
                                             }

                                             using (var dbContext = new JSE_RevitAddin_MEP_OPENINGS.Data.SleeveDbContext(_document))'
$content2 = $content2 -replace $targetFlag, $replacementFlag

$content2 | Set-Content $path2 -NoNewline
