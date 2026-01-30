# FINAL FIX INSTRUCTIONS

## ✅ ISSUE #1: ALREADY FIXED
Your `BulkPlacementService.cs` already has the transaction wrapper. Step 5 should be working.

## ⚠️ ISSUE #2: USE THIS POWERSHELL SCRIPT TO FIND THE DUPLICATE

Save this as `find_duplicate.ps1` and run it:

```powershell
# Navigate to project directory
cd "C:\JSE_CSharp_Projects\JSE_MEPOPENING_23"

Write-Host "=" * 80
Write-Host "SEARCHING FOR CLUSTER PLACEMENT CALLS..."
Write-Host "=" * 80

# Search for all occurrences of placement methods
$results = @()

Get-ChildItem -Recurse -Include *.cs | ForEach-Object {
    $file = $_.FullName
    $content = Get-Content $file -Raw
    
    # Check for ClusterSleeves calls
    if ($content -match "\.ClusterSleeves\(") {
        $lineNum = (Get-Content $file | Select-String "\.ClusterSleeves\(" | Select-Object -First 1).LineNumber
        $results += [PSCustomObject]@{
            File = $file.Replace("C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\", "")
            Line = $lineNum
            Method = "ClusterSleeves (OLD)"
        }
    }
    
    # Check for ClusterSleevesV2 calls
    if ($content -match "\.ClusterSleevesV2\(") {
        $lineNum = (Get-Content $file | Select-String "\.ClusterSleevesV2\(" | Select-Object -First 1).LineNumber
        $results += [PSCustomObject]@{
            File = $file.Replace("C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\", "")
            Line = $lineNum
            Method = "ClusterSleevesV2 (NEW)"
        }
    }
    
    # Check for PlaceClusterSleeve calls (from placement service)
    $matches = Select-String -Path $file -Pattern "PlaceClusterSleeve\(" -AllMatches
    if ($matches) {
        foreach ($match in $matches) {
            $results += [PSCustomObject]@{
                File = $file.Replace("C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\", "")
                Line = $match.LineNumber
                Method = "PlaceClusterSleeve (PLACEMENT)"
            }
        }
    }
}

# Display results
$results | Format-Table -AutoSize

Write-Host "`n" + "=" * 80
Write-Host "ANALYSIS:"
Write-Host "=" * 80

$clusterSleevesCount = ($results | Where-Object { $_.Method -eq "ClusterSleeves (OLD)" }).Count
$clusterSleevesV2Count = ($results | Where-Object { $_.Method -eq "ClusterSleevesV2 (NEW)" }).Count
$placementCount = ($results | Where-Object { $_.Method -eq "PlaceClusterSleeve (PLACEMENT)" }).Count

Write-Host "ClusterSleeves (OLD) calls: $clusterSleevesCount"
Write-Host "ClusterSleevesV2 (NEW) calls: $clusterSleevesV2Count"
Write-Host "PlaceClusterSleeve (PLACEMENT) calls: $placementCount"

if ($placementCount -gt 1) {
    Write-Host "`n⚠️  FOUND DUPLICATE: $placementCount calls to PlaceClusterSleeve!" -ForegroundColor Red
    Write-Host "Check the files listed above - one of them has a duplicate call."
}

if ($clusterSleevesCount -gt 0 -and $clusterSleevesV2Count -gt 0) {
    Write-Host "`n⚠️  BOTH OLD AND NEW methods are being called!" -ForegroundColor Yellow
    Write-Host "This might be the cause of duplicate placement."
}

Write-Host "`n" + "=" * 80
```

## 🎯 HOW TO USE THE SCRIPT

1. **Save the script** as `find_duplicate.ps1` in your project root
2. **Run PowerShell** as Administrator
3. **Execute**: `.\find_duplicate.ps1`
4. **Review the output** - it will show you exactly which files call which methods

## 📋 WHAT TO DO AFTER RUNNING THE SCRIPT

### Scenario A: If it shows PlaceClusterSleeve called TWICE
```
File: Services\Clustering\RefactoredClusterService.cs, Line: 850
File: Services\Clustering\RefactoredClusterService.cs, Line: 1200
```
→ Open that file and comment out the SECOND call

### Scenario B: If it shows BOTH ClusterSleeves and ClusterSleevesV2
```
File: Commands\SomeCommand.cs, Line: 100, Method: ClusterSleeves (OLD)
File: Commands\SomeCommand.cs, Line: 150, Method: ClusterSleevesV2 (NEW)
```
→ Open that file and comment out ONE of them (keep V2, remove OLD)

### Scenario C: If it shows multiple PlaceClusterSleeve calls in same file
```
File: Services\Clustering\RefactoredClusterService.cs, Line: 500
File: Services\Clustering\RefactoredClusterService.cs, Line: 1500
```
→ One is in a loop (correct), one is a duplicate call (wrong) - comment out the duplicate

## 🔧 THE FIX TEMPLATE

Once you find the duplicate, use this template:

```csharp
// ❌ DUPLICATE PLACEMENT - DISABLED TO FIX PERFORMANCE ISSUE
// This was causing clusters to be placed twice (880ms overhead)
// Discovered using diagnostic script on [DATE]
// [PASTE THE LINE YOU'RE COMMENTING OUT HERE]
```

## 📊 VERIFICATION

After fixing:
1. Rebuild solution (Ctrl+Shift+B)
2. Run tool in Revit
3. Check logs - should see ONE placement operation (~413ms)
4. Cluster rate should be 12+/sec (not 4.3/sec)

## 🆘 IF SCRIPT DOESN'T WORK

If PowerShell script fails or you can't run it:

### Manual Search (Visual Studio)
1. Open Visual Studio
2. Press Ctrl+Shift+F (Find in Files)
3. Search for: `PlaceClusterSleeve(`
4. Look at results - if same file appears TWICE with different line numbers, that's your duplicate

### Alternative: Add Stack Trace Logging
Add this to line ~300 in `ClusterPlacementService.cs` (right before `NewFamilyInstance`):

```csharp
// ADD DIAGNOSTIC
var stackTrace = new System.Diagnostics.StackTrace(true);
SafeFileLogger.SafeAppendText("cluster_debug.log", 
    $"[{DateTime.Now:HH:mm:ss}] 🔥 PLACEMENT #{++_callCount}:\n{stackTrace}\n");

// Then the existing line:
inst = doc.Create.NewFamilyInstance(...);
```

Add this field at top of class:
```csharp
private static int _callCount = 0;
```

This will show you the call stack every time placement happens.

---

## SUMMARY

Run the PowerShell script above - it will pinpoint exactly where the duplicate is.

Then comment out one of the duplicate calls and rebuild.

Your performance should immediately improve from 4.3 to 12+ sleeves/sec!
