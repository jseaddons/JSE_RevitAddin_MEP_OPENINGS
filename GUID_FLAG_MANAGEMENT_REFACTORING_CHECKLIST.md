# 📋 GUID & Flag Management - Refactoring Checklist

## Overview
This document lists ALL existing files and methods that need to be modified/removed/replaced when implementing the OOP refactoring plan.

**⚠️ CRITICAL: This checklist was updated to include the performance optimization that was missing from the original plan:**
- **Step 3a:** Pre-filter intersections using XML data before calling IntersectionDetectionService
- **Step 4:** Only process NEW/unmatched intersections (skips already-known clashes from XML)
- **Performance gain:** Reduces O(N × L × F) to O(M × L × F) where M << N (only new/changed zones)

---

## 🗂️ Files That Need Changes

### **1. Services/RefreshService.cs**

#### **🔴 REMOVE/REPLACE: Flag Sync Logic (Lines 855-914)**
**Current Code Location:** Lines 855-914
**Method:** Inline code block in `ExecuteRefreshInternal()`
**Action:** **REMOVE** this entire block and **REPLACE** with call to `FlagManager.SyncFlagsFromGlobal()`

**Code to Remove:**
```855:914:Services/RefreshService.cs
// ✅ CRITICAL FIX: Sync flags from Global XML but DO NOT REMOVE clash zones
// Clash zones must remain in Filter XML even if resolved (for OK button logic and history)
// Only update their flags to match Global XML state
if (existingClashZones?.ClashZones != null && existingClashZones.ClashZones.Count > 0)
{
    var syncedFromGlobalCount = 0;
    
    foreach (var cz in existingClashZones.ClashZones)
    {
        // Check Global XML for this clash zone's resolution status
        try
        {
            var globalIndex = GlobalIndexService.LoadOrCreate(_document, cz.MepElementCategory);
            var entry = globalIndex.Entries?.FirstOrDefault(e => e.Id == cz.Id.ToString());
            if (entry != null)
            {
                // ✅ CRITICAL: Sync flags from Global XML but KEEP the clash zone in Filter XML
                // This ensures OK button logic can see all clash zones (including resolved ones)
                bool flagChanged = false;
                
                if (cz.IsResolved != entry.IsResolved)
                {
                    cz.IsResolved = entry.IsResolved;
                    flagChanged = true;
                }
                
                if (cz.IsClusterResolved != entry.IsClusterResolved)
                {
                    cz.IsClusterResolved = entry.IsClusterResolved;
                    flagChanged = true;
                }
                
                if (entry.IsResolved && cz.SleeveInstanceId != entry.SleeveInstanceId)
                {
                    cz.SleeveInstanceId = entry.SleeveInstanceId;
                    flagChanged = true;
                }
                
                if (entry.IsClusterResolved && cz.ClusterSleeveInstanceId != entry.ClusterSleeveInstanceId)
                {
                    cz.ClusterSleeveInstanceId = entry.ClusterSleeveInstanceId;
                    flagChanged = true;
                }
                
                if (flagChanged)
                {
                    syncedFromGlobalCount++;
                    DebugLogger.Info($"[REFRESH-GLOBAL-SYNC] Synced ClashZone {cz.Id} from Global XML (IsResolved={entry.IsResolved}, IsClusterResolved={entry.IsClusterResolved})");
                }
            }
        }
        catch (Exception globalEx)
        {
            DebugLogger.Warning($"[REFRESH-GLOBAL-SYNC] Error syncing Global XML for ClashZone {cz.Id}: {globalEx.Message}");
        }
    }
    
    DebugLogger.Info($"[REFRESH-GLOBAL-SYNC] Synced flags from Global XML for {syncedFromGlobalCount} clash zones (total: {existingClashZones.ClashZones.Count})");
    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [REFRESH-GLOBAL-SYNC] Synced {syncedFromGlobalCount} clash zones from Global XML (total: {existingClashZones.ClashZones.Count})\n");
}
```

**Replace With:**
```csharp
// ✅ Use FlagManager for flag syncing (OOP refactoring)
if (existingClashZones?.ClashZones != null && existingClashZones.ClashZones.Count > 0)
{
    foreach (var category in existingClashZones.ClashZones.Select(cz => cz.MepElementCategory).Distinct())
    {
        var categoryClashZones = existingClashZones.ClashZones.Where(cz => cz.MepElementCategory == category).ToList();
        _flagManager.SyncFlagsFromGlobal(categoryClashZones, category);
    }
}
```

---

#### **🟡 MODIFY: 3-Point Validation Global XML Cleanup (Lines 1030-1072)**
**Current Code Location:** Lines 1030-1072
**Action:** **REPLACE** inline Global XML removal logic with `GuidManager.RemoveFromGlobalXml()`

**Code to Replace:**
```1030:1072:Services/RefreshService.cs
// Clear Global XML entries for invalid clash zones
var invalidByCategory = invalidClashZones.GroupBy(cz => cz.MepElementCategory);
foreach (var categoryGroup in invalidByCategory)
{
    var categoryName = categoryGroup.Key;
    var invalidIds = categoryGroup.Select(cz => cz.Id).ToList();
    
    try
    {
        var globalIndex = GlobalIndexService.LoadOrCreate(_document, categoryName);
        var beforeCount = globalIndex.Entries.Count;
        
        // Remove entries for invalid clash zones
        globalIndex.Entries.RemoveAll(e => 
        {
            Guid guid;
            return Guid.TryParse(e.Id, out guid) && invalidIds.Contains(guid);
        });
        
        var afterCount = globalIndex.Entries.Count;
        var removedFromGlobal = beforeCount - afterCount;
        
        if (removedFromGlobal > 0)
        {
            // ✅ Force save the modified global index using GlobalIndexService
            try
            {
                // Use GlobalIndexService.Save method to ensure consistent serialization
                GlobalIndexService.Save(_document, globalIndex);
                DebugLogger.Info($"[3-POINT-VALIDATION] Cleared {removedFromGlobal} entries from Global XML for category {categoryName}");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [3-POINT-VALIDATION] Cleared {removedFromGlobal} Global XML entries for {categoryName}\n");
            }
            catch (Exception saveEx)
            {
                DebugLogger.Warning($"[3-POINT-VALIDATION] Error saving Global XML for {categoryName}: {saveEx.Message}");
            }
        }
    }
    catch (Exception globalEx)
    {
        DebugLogger.Warning($"[3-POINT-VALIDATION] Error clearing Global XML for {categoryName}: {globalEx.Message}");
    }
}
```

**Replace With:**
```csharp
// ✅ Use GuidManager for Global XML cleanup (OOP refactoring)
foreach (var invalidClashZone in invalidClashZones)
{
    _guidManager.RemoveFromGlobalXml(invalidClashZone.Id, invalidClashZone.MepElementCategory);
}
```

---

#### **🟡 MODIFY: EnsureEntries Call (Line 2113)**
**Current Code Location:** Line 2113
**Action:** **REPLACE** with `GuidManager.EnsureGlobalXmlEntry()`

**Code to Replace:**
```2113:2113:Services/RefreshService.cs
GlobalIndexService.EnsureEntries(_document, category, categoryClashZones.Select(cz => cz.Id));
```

**Replace With:**
```csharp
// ✅ Use GuidManager for ensuring Global XML entries (OOP refactoring)
foreach (var clashZone in categoryClashZones)
{
    _guidManager.EnsureGlobalXmlEntry(clashZone, category);
}
```

---

### **2. Services/ClashZoneService.cs**

#### **🔴 REMOVE: ResetResolvedFlagForDeletedSleeves Method (Lines 1418-1642)**
**Current Code Location:** Lines 1418-1642
**Method:** `ResetResolvedFlagForDeletedSleeves()`
**Action:** **REMOVE** entire method - logic moves to `FlagManager.ResetFlagsForDeletedSleeves()`

**Code to Remove:**
```1418:1642:Services/ClashZoneService.cs
private void ResetResolvedFlagForDeletedSleeves(Document document, List<string> selectedCategories = null)
{
    // ... entire method implementation (225 lines) ...
}
```

**Replace Call Site (Line 753):**
```750:753:Services/ClashZoneService.cs
// CRITICAL: Check for existing sleeves and reset IsResolved flag if sleeves were deleted
// This allows re-placement of sleeves after manual deletion
// ⚠️ CRITICAL FIX: Only reset flags for selected categories to avoid affecting other filters/categories
ResetResolvedFlagForDeletedSleeves(document, selectedCategories);
```

**Replace With:**
```csharp
// ✅ Use FlagManager for flag reset (OOP refactoring)
// Note: FlagManager will be injected into ClashZoneService or called from orchestrator
// For now, keep call but move implementation to FlagManager
```

**⚠️ NOTE:** The call at line 753 should remain but implementation moves to `FlagManager`.

---

#### **🔴 REMOVE: SaveResetFlagsToBothXmls Method (Lines 1644-1689)**
**Current Code Location:** Lines 1644-1689
**Method:** `SaveResetFlagsToBothXmls()`
**Action:** **REMOVE** entire method - logic already in `FlagManager.ResetFlagsForDeletedSleeves()`

**Code to Remove:**
```1644:1689:Services/ClashZoneService.cs
/// <summary>
/// ✅ CRITICAL: Save reset flags to both Global XML and Filter XML
/// </summary>
private void SaveResetFlagsToBothXmls(Document document, List<string> selectedCategories)
{
    // ... entire method implementation (46 lines) ...
}
```

**Also Remove Call at Line 1634:**
```1634:1634:Services/ClashZoneService.cs
SaveResetFlagsToBothXmls(document, selectedCategories);
```

---

#### **🟢 KEEP: FindExistingClashZone Method (Lines 1320-1364)**
**Current Code Location:** Lines 1320-1364
**Method:** `FindExistingClashZone()`
**Action:** **KEEP** but **CONSIDER** moving to `GuidManager.FindByMepAndHost()` (optional refactoring)

**Current Code:**
```1320:1364:Services/ClashZoneService.cs
private ClashZone? FindExistingClashZone(ElementId mepElementId, ElementId structuralElementId, XYZ intersectionPoint)
{
    // ✅ CRITICAL FIX: Compare by IntegerValue to handle XML deserialization cases
    // ... implementation ...
}
```

**Option 1 (Recommended):** Keep in ClashZoneService (it accesses `_clashZoneStorage`)
**Option 2:** Move to GuidManager and pass `_clashZoneStorage.ClashZones` as parameter

---

### **3. Services/UniversalSleevePlacerService.cs**

#### **🟡 MODIFY: Global XML Update After Placement (Lines 1198-1235)**
**Current Code Location:** Lines 1198-1235
**Action:** **REPLACE** inline Global XML update with `FlagManager.UpdateFlagsForPlacement()`

**Code to Replace:**
```1198:1235:Services/UniversalSleevePlacerService.cs
// ✅ STEP 3: GLOBAL XML UPDATE - Use captured data to ensure correct state
// Per methodology document line 242-245: "Update Global XML" after placement
// This is the ONLY system for Global XML (GlobalFlagManager removed)
try
{
    if (placedClashZonesForGlobal.Count > 0)
    {
        var updatesByCategory = placedClashZonesForGlobal
        .GroupBy(cz => cz.MepElementCategory)
        .ToDictionary(g => g.Key, g => g.Select(cz => (cz.Id, cz.IsResolved, cz.IsClusterResolved, cz.SleeveInstanceId, cz.ClusterSleeveInstanceId)));

        foreach (var kvp in updatesByCategory)
        {
            var categoryName = kvp.Key;
                var updates = kvp.Value.ToList(); // Materialize to get count
                DebugLogger.Info($"[GLOBAL_INDEX] Updating Global XML for {updates.Count} placed clash zones in category '{categoryName}'");
                foreach (var update in updates.Take(5))
                {
                    DebugLogger.Info($"[GLOBAL_INDEX] Updating ClashZone {update.Id}: IsResolved={update.IsResolved}, SleeveInstanceId={update.SleeveInstanceId}");
                }
            GlobalIndexService.UpsertFlagsWithIds(_doc, categoryName, updates);
                DebugLogger.Info($"[GLOBAL_INDEX] ✅ Successfully updated Global XML for category '{categoryName}' with {updates.Count} entries");
            }
            
            DebugLogger.Info($"[GLOBAL_INDEX] ✅ Successfully updated Global XML for {placedClashZonesForGlobal.Count} placed sleeves (BEFORE clustering)");
        }
        else
        {
            DebugLogger.Warning($"[GLOBAL_INDEX] ⚠️ No placed clash zones found to update in Global XML (PlacedCount={PlacedCount}, but no clash zones with IsResolved=true AND SleeveInstanceId > 0)");
        }
    }
    catch (Exception upEx)
    {
        DebugLogger.Error($"[GLOBAL_INDEX] ❌ CRITICAL ERROR: Upsert after individual placement failed: {upEx.Message}");
        DebugLogger.Error($"[GLOBAL_INDEX] Stack trace: {upEx.StackTrace}");
        // Don't throw - Global XML failure shouldn't stop placement, but log it clearly
    }
}
```

**Replace With:**
```csharp
// ✅ Use FlagManager for flag updates (OOP refactoring)
try
{
    foreach (var clashZone in placedClashZonesForGlobal)
    {
        _flagManager.UpdateFlagsForPlacement(
            clashZone, 
            clashZone.SleeveInstanceId, 
            isCluster: false, 
            clashZone.MepElementCategory
        );
    }
}
catch (Exception upEx)
{
    DebugLogger.Error($"[FLAG-MANAGER] ❌ CRITICAL ERROR: Flag update after individual placement failed: {upEx.Message}");
}
```

---

#### **🟡 MODIFY: Global XML Check Before Placement (Lines 540-609)**
**Current Code Location:** Lines 540-609
**Action:** **REPLACE** inline Global XML check with `FlagManager` check (or keep as-is if it's pre-placement validation)

**Code to Review:**
```540:609:Services/UniversalSleevePlacerService.cs
// Global XML check logic before placement
// ... checks if sleeve already exists ...
```

**Decision:** 
- **Option A:** Keep as-is (it's pre-placement validation, not flag management)
- **Option B:** Move to `FlagManager.CheckIfPlacementNeeded()` (new method)

**Recommendation:** Keep as-is - it's placement logic, not flag management.

---

#### **🟡 MODIFY: Individual Flag Update Calls (Lines 572, 598)**
**Current Code Location:** Lines 572, 598
**Action:** **REPLACE** inline `GlobalIndexService.UpsertFlagsWithIds()` calls with `FlagManager.UpdateFlagsForPlacement()`

**Code to Replace:**
```572:572:Services/UniversalSleevePlacerService.cs
GlobalIndexService.UpsertFlagsWithIds(_doc, categoryName, new[] { (clashZone.Id, entry.IsResolved, false, entry.SleeveInstanceId, 0) });
```

```598:598:Services/UniversalSleevePlacerService.cs
GlobalIndexService.UpsertFlagsWithIds(_doc, categoryName, new[] { (clashZone.Id, false, entry.IsClusterResolved, 0, entry.ClusterSleeveInstanceId) });
```

**Replace With:**
```csharp
// ✅ Use FlagManager for flag updates (OOP refactoring)
_flagManager.UpdateFlagsForPlacement(clashZone, entry.SleeveInstanceId, isCluster: false, categoryName);
// or for cluster:
_flagManager.UpdateFlagsForPlacement(clashZone, entry.ClusterSleeveInstanceId, isCluster: true, categoryName);
```

---

### **4. Services/UniversalClusterService.cs**

#### **🟡 MODIFY: Global XML Update After Clustering (Lines 2666-2728)**
**Current Code Location:** Lines 2666-2728
**Action:** **REPLACE** inline Global XML update with `FlagManager.UpdateFlagsForPlacement()`

**Code to Replace:**
```2666:2728:Services/UniversalClusterService.cs
// ✅ GLOBAL FLAGS: Upsert IsResolved/IsClusterResolved into per-category global index after clustering
// ✅ CRITICAL: Use clash zones from MarkClashZonesAsClusterResolvedWithSleeveId (already updated in Filter XML)
// Load updated clash zones from XML to get correct flag values
try
{
    var updates = new List<(Guid Id, bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterSleeveInstanceId)>();
    
    // ✅ FIX: Load clash zones from the same XML file that was just updated
    if (!string.IsNullOrEmpty(xmlFilePath) && File.Exists(xmlFilePath))
    {
        // ... XML loading and updates collection ...
    }
    
    // Fallback: Use cache if XML not available
    if (updates.Count == 0)
    {
        // ... fallback logic ...
    }

    if (updates.Count > 0)
    {
        GlobalIndexService.UpsertFlagsWithIds(_doc, targetCategory, updates);
        DebugLogger.Info($"[GLOBAL_INDEX] Updated Global XML for {updates.Count} clash zones after clustering");
    }
}
catch (Exception upEx)
{
    DebugLogger.Warning($"[GLOBAL_INDEX] Upsert after clustering failed: {upEx.Message}");
}
```

**Replace With:**
```csharp
// ✅ Use FlagManager for flag updates (OOP refactoring)
try
{
    // Load updated clash zones from XML to get correct flag values
    if (!string.IsNullOrEmpty(xmlFilePath) && File.Exists(xmlFilePath))
    {
        var filter = _filterService.LoadFilterFromXmlFile(xmlFilePath);
        if (filter?.ClashZoneStorage?.ClashZones != null)
        {
            foreach (var s in cluster)
            {
                var originalSleeveId = s.SleeveInstanceId;
                var clashZone = filter.ClashZoneStorage.ClashZones.FirstOrDefault(cz => 
                    cz.SleeveInstanceId == originalSleeveId || 
                    cz.AfterClusterSleevePlacedSleeveInstanceId == originalSleeveId);
                
                if (clashZone != null)
                {
                    _flagManager.UpdateFlagsForPlacement(
                        clashZone, 
                        inst.Id.IntegerValue, 
                        isCluster: true, 
                        targetCategory
                    );
                }
            }
        }
    }
}
catch (Exception upEx)
{
    DebugLogger.Warning($"[FLAG-MANAGER] Flag update after clustering failed: {upEx.Message}");
}
```

**⚠️ NOTE:** Need to inject `FlagManager` and `FilterManagementService` into `UniversalClusterService`.

---

## 📝 New Files to Create

### **1. Services/FlagManager.cs** (NEW FILE)
**Action:** **CREATE** new file with `FlagManager` class
**Location:** `Services/FlagManager.cs`
**Lines:** ~280 lines (based on implementation plan)

### **2. Services/GuidManager.cs** (NEW FILE)
**Action:** **CREATE** new file with `GuidManager` class
**Location:** `Services/GuidManager.cs`
**Lines:** ~40 lines (based on implementation plan)

### **3. Services/RefreshOrchestrator.cs** (NEW FILE - Optional)
**Action:** **CREATE** new file with `RefreshOrchestrator` class
**Location:** `Services/RefreshOrchestrator.cs`
**Lines:** ~200 lines (based on implementation plan)

**⚠️ NOTE:** RefreshOrchestrator is optional - you can integrate the sequence directly into `RefreshService.ExecuteRefreshInternal()` if preferred.

---

## 🔧 Dependency Injection Changes

### **RefreshService.cs**
**Add Fields:**
```csharp
private readonly FlagManager _flagManager;
private readonly GuidManager _guidManager;
```

**Update Constructor:**
```csharp
public RefreshService(Document document, UIDocument uiDocument, ApplicationProfileService appProfileService)
{
    // ... existing initialization ...
    _flagManager = new FlagManager(document);
    _guidManager = new GuidManager(document);
}
```

### **ClashZoneService.cs**
**Add Fields:**
```csharp
private readonly FlagManager _flagManager;  // Optional - for flag reset
```

**Update Constructor (if needed):**
```csharp
public ClashZoneService(ClashZoneStorage existingClashZones, Action<string> logAction, FlagManager flagManager = null)
{
    // ... existing initialization ...
    _flagManager = flagManager;  // Optional dependency
}
```

**Update Method Call (Line 753):**
```csharp
// ✅ Use FlagManager if available, otherwise use existing method
if (_flagManager != null)
{
    _flagManager.ResetFlagsForDeletedSleeves(_clashZoneStorage.ClashZones, selectedCategories);
}
else
{
    // Fallback to existing implementation (for backward compatibility)
    ResetResolvedFlagForDeletedSleeves(document, selectedCategories);
}
```

### **UniversalSleevePlacerService.cs**
**Add Fields:**
```csharp
private readonly FlagManager _flagManager;
```

**Update Constructor:**
```csharp
// Add FlagManager parameter
public UniversalSleevePlacerService(/* existing params */, FlagManager flagManager)
{
    // ... existing initialization ...
    _flagManager = flagManager;
}
```

### **UniversalClusterService.cs**
**Add Fields:**
```csharp
private readonly FlagManager _flagManager;
private readonly FilterManagementService _filterService;  // For loading XML
```

**Update Constructor:**
```csharp
// Add FlagManager and FilterManagementService parameters
public UniversalClusterService(/* existing params */, FlagManager flagManager, FilterManagementService filterService)
{
    // ... existing initialization ...
    _flagManager = flagManager;
    _filterService = filterService;
}
```

---

## 📊 Summary of Changes

### **Files to Modify:**
1. ✅ **RefreshService.cs** - 3 locations (lines 855-914, 1030-1072, 2113)
2. ✅ **ClashZoneService.cs** - 3 methods (lines 1418-1642, 1644-1689, keep 1320-1364)
3. ✅ **UniversalSleevePlacerService.cs** - 3 locations (lines 1198-1235, 572, 598)
4. ✅ **UniversalClusterService.cs** - 1 location (lines 2666-2728)

### **Files to Create:**
1. ✅ **Services/FlagManager.cs** - NEW (280 lines)
2. ✅ **Services/GuidManager.cs** - NEW (40 lines)
3. ✅ **Services/RefreshOrchestrator.cs** - NEW (200 lines, optional)

### **Methods to Remove:**
1. ❌ `RefreshService`: Inline flag sync (lines 855-914) → Replace with `FlagManager.SyncFlagsFromGlobal()`
2. ❌ `RefreshService`: Inline Global XML cleanup (lines 1030-1072) → Replace with `GuidManager.RemoveFromGlobalXml()`
3. ❌ `ClashZoneService`: `ResetResolvedFlagForDeletedSleeves()` (lines 1418-1642) → Replace with `FlagManager.ResetFlagsForDeletedSleeves()`
4. ❌ `ClashZoneService`: `SaveResetFlagsToBothXmls()` (lines 1644-1689) → Logic already in FlagManager

### **Methods to Modify:**
1. 🔄 `RefreshService.ExecuteRefreshInternal()` - Replace inline flag sync and Global XML cleanup
2. 🔄 `UniversalSleevePlacerService` placement method - Replace Global XML updates with FlagManager calls
3. 🔄 `UniversalClusterService.PlaceClusterSleeve()` - Replace Global XML updates with FlagManager calls

### **Methods to Keep (Unchanged):**
1. ✅ `ClashZoneService.FindExistingClashZone()` (lines 1320-1364) - Keep or optionally move to GuidManager
2. ✅ `GlobalIndexService` - ALL methods stay unchanged
3. ✅ `FilterManagementService` - ALL methods stay unchanged
4. ✅ `ProjectPathService` - ALL methods stay unchanged

---

## ✅ Implementation Order

1. **Step 1:** Create `FlagManager.cs` (NEW)
   - Implement `SyncFlagsFromGlobal()`
   - Implement `ResetFlagsForDeletedSleeves()`
   - Implement `UpdateFlagsForPlacement()`

2. **Step 2:** Create `GuidManager.cs` (NEW)
   - Implement `GenerateNewGuid()`
   - Implement `FindByMepAndHost()`
   - Implement `EnsureGlobalXmlEntry()`
   - Implement `RemoveFromGlobalXml()`

3. **Step 3:** Modify `RefreshService.cs`
   - Remove inline flag sync (lines 855-914)
   - Add FlagManager field and initialization
   - Replace with `FlagManager.SyncFlagsFromGlobal()`
   - Replace Global XML cleanup with `GuidManager.RemoveFromGlobalXml()`
   - Replace `EnsureEntries` with `GuidManager.EnsureGlobalXmlEntry()`

4. **Step 4:** Modify `ClashZoneService.cs`
   - Add FlagManager field (optional dependency)
   - Remove `ResetResolvedFlagForDeletedSleeves()` method (lines 1418-1642)
   - Remove `SaveResetFlagsToBothXmls()` method (lines 1644-1689)
   - Replace call at line 753 with `FlagManager.ResetFlagsForDeletedSleeves()`

5. **Step 5:** Modify `UniversalSleevePlacerService.cs`
   - Add FlagManager field
   - Replace Global XML update at lines 1198-1235 with `FlagManager.UpdateFlagsForPlacement()`
   - Replace inline Global XML calls at lines 572, 598 with FlagManager calls

6. **Step 6:** Modify `UniversalClusterService.cs`
   - Add FlagManager and FilterManagementService fields
   - Replace Global XML update at lines 2666-2728 with `FlagManager.UpdateFlagsForPlacement()`

7. **Step 7:** (Optional) Create `RefreshOrchestrator.cs`
   - Implement orchestration sequence
   - Integrate into `RefreshService.ExecuteRefreshInternal()`

---

## ⚠️ Important Notes

1. **NO XML I/O code duplication** - All XML operations use existing services
2. **Backward compatibility** - Consider making FlagManager optional in ClashZoneService initially
3. **Testing** - Test each step before moving to the next
4. **Line numbers may shift** - After removing code, line numbers in this document will change
5. **Logging** - Preserve all existing logging or migrate to FlagManager/GuidManager

---

## 📍 Quick Reference: Line Number Summary

| File | Lines | Action | New Method |
|------|-------|--------|------------|
| RefreshService.cs | 855-914 | REMOVE | FlagManager.SyncFlagsFromGlobal() |
| RefreshService.cs | 1030-1072 | REPLACE | GuidManager.RemoveFromGlobalXml() |
| RefreshService.cs | 2113 | REPLACE | GuidManager.EnsureGlobalXmlEntry() |
| ClashZoneService.cs | 1418-1642 | REMOVE | FlagManager.ResetFlagsForDeletedSleeves() |
| ClashZoneService.cs | 1644-1689 | REMOVE | (Logic in FlagManager) |
| UniversalSleevePlacerService.cs | 1198-1235 | REPLACE | FlagManager.UpdateFlagsForPlacement() |
| UniversalSleevePlacerService.cs | 572 | REPLACE | FlagManager.UpdateFlagsForPlacement() |
| UniversalSleevePlacerService.cs | 598 | REPLACE | FlagManager.UpdateFlagsForPlacement() |
| UniversalClusterService.cs | 2666-2728 | REPLACE | FlagManager.UpdateFlagsForPlacement() |

---

## ✅ Verification Checklist

After implementation, verify:
- [ ] All flag sync logic removed from RefreshService
- [ ] All flag reset logic removed from ClashZoneService
- [ ] All Global XML updates use FlagManager
- [ ] All GUID operations use GuidManager
- [ ] No duplicate XML I/O code
- [ ] All existing tests pass
- [ ] New FlagManager tests written
- [ ] New GuidManager tests written

