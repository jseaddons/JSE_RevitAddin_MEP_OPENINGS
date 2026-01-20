# Ventilation Filter UI Persistence Bug Fix

## Problem Summary

UI persistence for the **Ventilation** filter (and potentially other filters) was not working properly, while Plumbing and Electrical filters worked correctly.

**Symptoms:**
- Ventilation filter UI state not saved to database
- When switching away and back to Ventilation filter, previous selections were lost
- Other filters (Plumbing, Electrical) worked correctly

## Root Cause Analysis

**Location:** `Services/FilterManagementService.cs`

The bug was in four locations where `SaveFilterUIState` was called:
1. Line 241: `CreateNewFilter()` method
2. Line 391: `CopyFilter()` method  
3. Line 756: `SaveFilter()` method (PRIMARY - most frequently used)
4. Line 1797: `SaveFilterToXmlFile()` method (fallback when XML creation disabled)

### The Issue

All four calls to `SaveFilterUIState` were **missing the last three optional parameters**:
- `List<string> selectedMepCategoryNames`
- `List<string> selectedReferenceFiles`
- `List<string> selectedHostFiles`

**Method Signature (FilterRepository.cs, line 247):**
```csharp
public void SaveFilterUIState(
    string filterName, 
    string category, 
    List<string> selectedHostCategories, 
    OpeningSettings openingSettings, 
    List<string> selectedMepCategoryNames = null,      // ← MISSING in calls!
    List<string> selectedReferenceFiles = null,        // ← MISSING in calls!
    List<string> selectedHostFiles = null              // ← MISSING in calls!
)
```

### What Was Happening

When a user saved a filter (especially Ventilation), the FilterManagementService was:
1. ✅ Collecting current UI selections into local variables
2. ✅ Passing `selectedHostCategories` to database
3. ✅ Passing `openingSettings` to database
4. ❌ **NOT passing** MEP category names, reference files, and host files to database

This caused these three critical UI state fields to not be persisted in the database, resulting in loss of UI state when the filter was loaded again.

**Why did Plumbing/Electrical work but not Ventilation?**
- The primary persistence mechanism still worked (categoryDisplay logic)
- But the **secondary UI state fields** (MEP categories, reference files, host files) were being silently dropped
- Different users may have had different configurations that masked the issue

## Solution

Fixed all four locations to pass the complete UI state to `SaveFilterUIState`:

### Fix in SaveFilter() method (line 554-761)

**Before:**
```csharp
var currentHostCategories = FilterUiStateProvider.GetSelectedHostCategories?.Invoke() ?? selectedFilter.SelectedHostCategories ?? new List<string>();

// ... missing collection of other state variables ...

UseFilterRepository(repo =>
{
    repo.SaveFilterUIState(
        selectedFilter.Name,
        categoryDisplay,
        currentHostCategories ?? new List<string>(),
        currentOpeningSettings
        // ❌ Missing last 3 parameters!
    );
});
```

**After:**
```csharp
var currentHostCategories = FilterUiStateProvider.GetSelectedHostCategories?.Invoke() ?? selectedFilter.SelectedHostCategories ?? new List<string>();
var currentMepCategoryNames = FilterUiStateProvider.GetSelectedMepCategoryNames?.Invoke() ?? new List<string>();
var currentReferenceFiles = FilterUiStateProvider.GetSelectedReferenceFiles?.Invoke() ?? new List<string>();
var currentHostFiles = FilterUiStateProvider.GetSelectedHostFiles?.Invoke() ?? new List<string>();

// ...

UseFilterRepository(repo =>
{
    repo.SaveFilterUIState(
        selectedFilter.Name,
        categoryDisplay,
        currentHostCategories ?? new List<string>(),
        currentOpeningSettings,
        currentMepCategoryNames,        // ✅ FIXED: Pass MEP category names
        currentReferenceFiles,          // ✅ FIXED: Pass reference files
        currentHostFiles                // ✅ FIXED: Pass host files
    );
});
_log($"[FILTER_MGMT] ✅ Saved UI state for filter '{selectedFilter.Name}' to database (HostCategories: {currentHostCategories?.Count ?? 0}, MepCategories: {currentMepCategoryNames?.Count ?? 0}, RefFiles: {currentReferenceFiles?.Count ?? 0}, HostFiles: {currentHostFiles?.Count ?? 0})");
```

### Similar Fixes Applied to

1. **CreateNewFilter()** (line 241): Uses properties from `newFilter` object
2. **CopyFilter()** (line 391): Uses properties from `copiedFilter` object  
3. **SaveFilterToXmlFile()** (line 1797): Uses properties from `filter` object

## Files Modified

- `Services/FilterManagementService.cs`

## Database Impact

**Table:** `Filters`

**Columns affected (now properly persisted):**
- `SelectedMepCategoryNames` - JSON array of selected MEP categories
- `SelectedReferenceFiles` - JSON array of selected reference files
- `SelectedHostFiles` - JSON array of selected host files

These columns already existed in the database schema but were receiving NULL values because the code wasn't passing the data.

## Testing

### Before Fix
1. Open Ventilation filter
2. Select specific MEP categories, reference files, and host files
3. Switch to another filter
4. Switch back to Ventilation filter
5. **Result:** UI state lost (options not checked)

### After Fix
1. Open Ventilation filter
2. Select specific MEP categories, reference files, and host files
3. Switch to another filter
4. Switch back to Ventilation filter
5. **Result:** UI state preserved (options still checked)

## Logging

Enhanced logging added to track what's being saved:

**Before (minimal logging):**
```
✅ Saved UI state for filter 'Ventilation' to database
```

**After (detailed logging):**
```
✅ Saved UI state for filter 'Ventilation' to database (HostCategories: 2, MepCategories: 1, RefFiles: 1, HostFiles: 1)
```

This helps diagnose future persistence issues by showing exactly what was persisted.

## Related Code

- **FilterRepository.SaveFilterUIState()** - Database persistence (line 247)
- **FilterRepository.LoadFilterUIState()** - Database loading (line 415)  
- **FilterUiStateProvider** - UI state provider interface
- **EmergencyMainDialog.cs** - Already had correct implementation (line 8343)

## Backward Compatibility

✅ **No breaking changes** - All parameters are optional with NULL defaults
✅ **Existing data unaffected** - NULL values in new fields don't break existing filters
✅ **Gradual migration** - Filters created/updated with fix will have full state persistence

## Impact

**Low Risk** - Only affects filter UI state persistence, which is non-blocking functionality
**High Value** - Fixes user experience for all filter save/load operations
**Performance** - No performance impact (same database columns, just proper data now)

---

**Date Fixed:** December 5, 2025
**Commit:** [To be filled with commit hash]
**Build Status:** ✅ Succeeded (2782 warnings, 0 errors)
