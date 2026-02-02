$path = "c:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Services\OpeningCommandOrchestrator.cs"
$content = Get-Content $path -Raw
$target = 'SafeFileLogger.SafeAppendTextAlways\("placement_debug.log", "`n\[" \+ \(Get-Date -Format "HH:mm:ss"\) \+ "\] \[BULK-PLACEMENT-START\] allBulkTaskItems.Count = " \+ \$allBulkTaskItems.Count \+ "`n"\);'
$replacement = 'SafeFileLogger.SafeAppendTextAlways("placement_debug.log", $"\n[{DateTime.Now:HH:mm:ss}] [BULK-PLACEMENT-START] allBulkTaskItems.Count = {allBulkTaskItems.Count}\n");'
$newContent = $content -replace $target, $replacement
$newContent | Set-Content $path -NoNewline
