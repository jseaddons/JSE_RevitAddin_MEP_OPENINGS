# 🚨 CRITICAL BUG: Sleeves Disappear When Opening Main UI

## Problem Report

**Issue:** All existing sleeves in Revit **disappear the moment the Main UI dialog appears** after spinning/loading.

**Timeline:**
1. User clicks "Open Main Interface" button
2. Dialog spins/loads for a while
3. **The moment the Main UI appears → ALL SLEEVES VANISH**
4. No refresh or OK button click needed - just opening the dialog deletes sleeves

## Investigation Findings

### What We Checked:

1. ✅ **MainUI Log** - Only shows initialization header, then stops (dialog loads successfully)
2. ✅ **Dialog Constructor** - No Revit API calls that could delete sleeves
3. ✅ **InitializeComponent()** - Only creates UI controls, no Revit operations
4. ✅ **LoadRealLinkedFiles()** - Only loads linked file information, no deletions
5. ✅ **RepopulateSectionsAfterLoad()** - Only rebuilds UI layout, no Revit API calls
6. ✅ **No automatic refresh** - Refresh only happens when user clicks button
7. ✅ **No automatic external event** - External event only raised by OK button

### What We Know:

1. **Sleeves physically disappear from Revit model** - not just hidden or filtered
2. **Deletion happens when dialog becomes visible** - after initialization completes
3. **No error messages in logs** - silent deletion
4. **placement_debug.log doesn't capture it** - only logs during OK button click

## Hypothesis

The issue is likely related to **one of these mechanisms**:

### Hypothesis 1: ResetResolvedFlagForDeletedSleeves Bug
**Problem:** The `ResetResolvedFlagForDeletedSleeves` method might be incorrectly **reporting that sleeves don't exist** due to:
- **Tolerance mismatch** - Using 1mm tolerance but stored placement points are slightly different
- **LocationCurve vs LocationPoint** - Not handling all location types correctly
- **Family name matching** - Opening families might have different names than expected

**Evidence:**
- Method checks if sleeves exist at stored placement points
- If not found, sets `IsResolved=false`
- But this shouldn't **DELETE** sleeves, only reset flags

### Hypothesis 2: Automatic Transaction Rollback
**Problem:** When dialog opens, Revit might be:
- **Rolling back uncommitted transactions** that placed sleeves
- **Validating elements** and marking invalid sleeves for deletion
- **Reloading families** which invalidates existing instances

**Evidence:**
- Sleeves disappear when UI becomes visible (Revit might process pending operations)
- No explicit deletion code found in dialog initialization

### Hypothesis 3: External Event Auto-Execution
**Problem:** The `_sleevePlacementEvent` might be **automatically raised** when the dialog opens, causing:
- Placement of new sleeves
- **Cluster command execution** → Deletes individual sleeves
- Silent failure leaving no sleeves

**Evidence:**
- Line 242: `_sleevePlacementEvent = ExternalEvent.Create(_sleevePlacementHandler);`
- ExternalEvent might auto-execute if context/categories are already set

### Hypothesis 4: Document.Regenerate() Side Effect
**Problem:** Even though we don't see explicit `Document.Regenerate()` calls in dialog initialization, Revit might be:
- **Auto-regenerating** when dialog becomes visible
- **Validating elements** during regeneration
- **Removing invalid sleeves** that have parameter issues (like the "Level" parameter error in logs)

**Evidence:**
- Log shows: `[HOST-PARAM-MISSING] Sleeve 835376: Parameter 'Level' not found or read-only ✗`
- Sleeves with parameter errors might be auto-deleted during document validation

## Proposed Solutions

### Solution 1: Add Logging to Dialog.Shown Event
Add comprehensive logging to capture what's happening when dialog becomes visible:

```csharp
this.Shown += (s, e) => {
    DebugLogger.Info("=== DIALOG SHOWN EVENT TRIGGERED ===");
    
    // Count sleeves BEFORE LoadRealLinkedFiles
    var sleevesBeforeCount = new FilteredElementCollector(document)
        .OfClass(typeof(FamilyInstance))
        .Cast<FamilyInstance>()
        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
        .Count();
    DebugLogger.Info($"Sleeves BEFORE LoadRealLinkedFiles: {sleevesBeforeCount}");
    
    LoadRealLinkedFiles(document);
    
    // Count sleeves AFTER LoadRealLinkedFiles
    var sleevesAfterCount = new FilteredElementCollector(document)
        .OfClass(typeof(FamilyInstance))
        .Cast<FamilyInstance>()
        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
        .Count();
    DebugLogger.Info($"Sleeves AFTER LoadRealLinkedFiles: {sleevesAfterCount}");
    DebugLogger.Info($"Sleeves DELETED: {sleevesBeforeCount - sleevesAfterCount}");
};
```

### Solution 2: Disable ResetResolvedFlagForDeletedSleeves Temporarily
Comment out the call to `ResetResolvedFlagForDeletedSleeves` in `ClashZoneService.cs` to see if this is the culprit.

### Solution 3: Check ExternalEvent Status
Add logging to check if ExternalEvent is being raised automatically:

```csharp
_sleevePlacementEvent = ExternalEvent.Create(_sleevePlacementHandler);
DebugLogger.Info($"External Event created - IsPending: {_sleevePlacementEvent.IsPending}");
```

### Solution 4: Fix Parameter Issues
Ensure all sleeves have valid "Level" parameters to prevent auto-deletion during document validation.

## Next Steps

1. **Add logging to Dialog.Shown event** to capture sleeve counts before/after
2. **Check if ExternalEvent is auto-executing** by adding logging
3. **Temporarily disable ResetResolvedFlagForDeletedSleeves** to isolate the issue
4. **Check Revit's Journal file** for any auto-deletion commands
5. **Test with a single sleeve** to see if it's a bulk operation or per-sleeve issue

## Files Modified

- `Services/ClashZoneService.cs` - Contains `ResetResolvedFlagForDeletedSleeves` method
- `Views/EmergencyMainDialog.cs` - Dialog initialization and Shown event handlers

## Status

🔴 **CRITICAL BUG** - Needs immediate investigation and fix

