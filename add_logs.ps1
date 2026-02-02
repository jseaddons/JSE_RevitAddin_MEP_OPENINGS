$path = "c:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Services\OpeningCommandOrchestrator.cs"
$content = Get-Content $path -Raw
$target = 'SafeFileLogger.SafeAppendTextAlways\("placement_debug.log", \$\"\\n\[\{DateTime.Now:HH:mm:ss\}\] \[BULK-PLACEMENT-START\] allBulkTaskItems.Count = \{allBulkTaskItems.Count\}\\n\"\);'
$replacement = 'SafeFileLogger.SafeAppendTextAlways("placement_debug.log", $"\n[{DateTime.Now:HH:mm:ss}] [BULK-PLACEMENT-START] allBulkTaskItems.Count = {allBulkTaskItems.Count}\n");

                                 // Check for duplicates
                                 var duplicateZones = allBulkTaskItems
                                     .GroupBy(x => x.Zone.ClashZoneGuid)
                                     .Where(g => !string.IsNullOrEmpty(g.Key) && g.Count() > 1)
                                     .ToList();

                                 if (duplicateZones.Count > 0)
                                 {
                                     SafeFileLogger.SafeAppendTextAlways("placement_debug.log", 
                                         $"[{DateTime.Now:HH:mm:ss}] [WARNING] Found {duplicateZones.Count} duplicate zones in allBulkTaskItems!\n");
                                     foreach (var dup in duplicateZones)
                                     {
                                         SafeFileLogger.SafeAppendTextAlways("placement_debug.log", 
                                             $"  - Zone {dup.Key} appears {dup.Count()} times\n");
                                     }
                                 }'
$newContent = $content -replace $target, $replacement
$newContent | Set-Content $path -NoNewline

$path2 = "c:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Services\BulkPlacementService.cs"
$content2 = Get-Content $path2 -Raw
$target2 = 'createdIds = doc.Create.NewFamilyInstances2\(creationDataList\);'
$replacement2 = 'createdIds = doc.Create.NewFamilyInstances2(creationDataList);
                        
                        // ✅ DIAGNOSTIC: Log actual Revit creation count (User Request)
                        var idListDiag = createdIds.ToList();
                        SafeFileLogger.SafeAppendTextAlways("placement_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] [BULK-PLACEMENT-RESULT] createdIds.Count = {idListDiag.Count}\n");'
$newContent2 = $content2 -replace $target2, $replacement2
$newContent2 | Set-Content $path2 -NoNewline
