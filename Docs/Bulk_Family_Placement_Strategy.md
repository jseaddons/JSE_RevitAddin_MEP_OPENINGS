# Bulk Family Placement Strategy

**Document Version:** 2.0  
**Date:** January 2026  
**Status:** Implementation Guide

---

## Table of Contents

1. [Overview](#1-overview)
2. [Complete Placement Flow](#2-complete-placement-flow)
3. [Database Query Strategy](#3-database-query-strategy)
4. [Size & Family Calculation](#4-size--family-calculation)
5. [Bulk Placement with NewFamilyInstances2](#5-bulk-placement-with-newfamilyinstances2)
6. [SOLID Architecture & Crash Recovery](#6-solid-architecture--crash-recovery)
7. [Performance Optimization](#7-performance-optimization)
8. [Implementation Checklist](#8-implementation-checklist)

---

## 1. Overview

### 1.1 Goal

Replace per-sleeve placement (slow, many API calls) with **bulk family placement** using Revit's `NewFamilyInstances2` API for maximum performance.

### 1.2 Before vs After

| Aspect | Before (Per-Sleeve) | After (Bulk) |
|--------|---------------------|--------------|
| **Revit API Calls** | N × NewFamilyInstance | 1 × NewFamilyInstances2 |
| **Transactions** | N per category | 1 for ALL |
| **Regeneration** | After each sleeve | 1 at the end |
| **DB Queries** | N queries | 1 bulk query |
| **Symbol Activation** | Per-zone | Once upfront |
| **200 sleeves** | ~2-5 minutes | ~10-30 seconds |

### 1.3 Prerequisites

- ✅ `IsCurrentClashFlag` and `ReadyForPlacementFlag` set during Refresh
- ✅ `SleevePlacementPointX/Y/Z` pre-calculated during Refresh
- ✅ MEP element data stored in DB: `MepElementWidth/Height`, `MepElementRotationAngle`
- ✅ `NewFamilyInstances2` available (Revit 2015+, confirmed for 2023)

### 1.4 Scope & Safety Flag

> **⚠️ IMPORTANT: This implementation covers INDIVIDUAL SLEEVES ONLY.**
>
> - **Cluster sleeves:** Will be implemented in Phase 2 after testing individual sleeves
> - **Post-placement parameter capture:** Handled by `ParameterSnapshotService` (separate process)
> - **Post-placement marking:** Handled by `MarkParameterService` (separate process)

**Safety Flag (default=false for safe rollout):**

```csharp
// OptimizationFlags.cs
public static bool UseBulkIndividualSleevePlacement { get; set; } = false;
```

**To enable:** Set to `true` after testing.

---

## 2. Complete Placement Flow

```
USER CLICKS "PLACE ALL SLEEVES"
    │
    ├─ STEP 1: Query DB
    │     WHERE IsCurrentClashFlag=1 AND ReadyForPlacementFlag=1
    │     (ReadyForPlacement guarantees unresolved - no redundant checks needed)
    │
    ├─ STEP 2: Calculate sizes & family names using CURRENT UI settings
    │     For each zone (sequential, ~10-20ms total for 200 zones):
    │       • SleeveWidth = MepWidth + (2 × Clearance) + InsulationThickness
    │       • SleeveHeight = MepHeight + (2 × Clearance) + InsulationThickness  
    │       • isCircular = DetermineOpeningType(zone) ← Complex rules!
    │       • SleeveFamilyName = GetSleeveFamilyName(zone, isCircular)
    │
    ├─ STEP 3: SAVE calculated sizes to DB
    │     UPDATE ClashZones SET SleeveWidth, SleeveHeight, SleeveDiameter, SleeveFamilyName
    │     WHERE ClashZoneId IN (@ids)
    │
    ├─ STEP 4: Pre-activate all family symbols (single transaction)
    │     Find unique family names → Load/Activate each once → Cache
    │
    ├─ STEP 5: Build FamilyInstanceCreationData list
    │     For each zone: new FamilyInstanceCreationData(point, symbol, StructuralType.NonStructural)
    │     ⚠️ NO LEVEL parameter for absolute Z coordinate placement!
    │
    ├─ STEP 6: Single Transaction - Bulk Placement
    │     createdIds = doc.Create.NewFamilyInstances2(creationDataList)
    │     For each created instance:
    │       • Set Width/Height/Diameter parameters
    │       • Apply Rotation (for floors)
    │       • Set Level parameter via INSTANCE_REFERENCE_LEVEL_PARAM
    │     doc.Regenerate() once
    │     tx.Commit()
    │
    └─ STEP 7: Batch update DB
          UPDATE ClashZones SET SleeveInstanceId, IsResolvedFlag=1, ReadyForPlacementFlag=0
```

---

## 3. Database Query Strategy

### 3.1 Unified Query for ALL Categories

**Repository Method:** `GetZonesReadyForPlacement()`

```sql
SELECT 
    ClashZoneGuid, ClashZoneId, SleeveFamilyName,
    StructuralElementType, HostOrientation, DuctShape, MepElementCategory,
    SleevePlacementPointX, SleevePlacementPointY, SleevePlacementPointZ,
    SleeveWidth, SleeveHeight, SleeveDiameter, StructuralElementThickness,
    MepElementWidth, MepElementHeight, MepElementOuterDiameter,
    MepElementRotationAngle, IsInsulated, InsulationThickness,
    MepElementLevelName, MepElementLevelElevation
FROM ClashZones
WHERE IsCurrentClashFlag = 1        -- Active in current session (section box + filters)
  AND ReadyForPlacementFlag = 1     -- Guaranteed unresolved
ORDER BY MepElementCategory, SleeveFamilyName
```

**Why only 2 flags?**
- `SetReadyForPlacementForUnresolvedZonesInSectionBox` already applies: `IsResolvedFlag=0 AND IsClusterResolvedFlag=0 AND IsCombinedResolved=0`
- If `ReadyForPlacementFlag = 1`, the zone is **guaranteed** unresolved
- ORDER BY groups by category for optimal symbol caching

---

## 4. Size & Family Calculation

### 4.1 Size Calculation

Uses **CURRENT UI settings** at placement time (not stale values from Refresh):

```csharp
var uiSettings = OpeningSettingsHelper.GetClearanceSettings();

foreach (var zone in zones)
{
    double clearance = GetClearance(zone.MepElementCategory, zone.IsInsulated, uiSettings);
    zone.SleeveWidth = zone.MepElementWidth + (2 * clearance) + (zone.IsInsulated ? zone.InsulationThickness * 2 : 0);
    zone.SleeveHeight = zone.MepElementHeight + (2 * clearance) + (zone.IsInsulated ? zone.InsulationThickness * 2 : 0);
    
    var (_, _, diameter, isCircular) = CalculateSleeveDimensions(zone);
    zone.SleeveDiameter = diameter;
    zone.SleeveFamilyName = GetSleeveFamilyName(zone, isCircular);
}
```

### 4.2 Family Name Determination (DetermineOpeningType)

**Critical:** The `isCircular` boolean involves complex rules, not just diameter check.

```
┌─────────────────────────────────────────────────────────────────┐
│                    DetermineOpeningType()                        │
└─────────────────────────────────────────────────────────────────┘
                              │
        ┌─────────────────────┼─────────────────────┐
        ▼                     ▼                     ▼
┌───────────────┐    ┌───────────────┐    ┌───────────────┐
│    PIPES      │    │    DUCTS      │    │    OTHER      │
│ • UI Pref     │    │ • DuctShape   │    │ (Cable Trays) │
│ • GetResolved │    │ • UI Pref     │    │ → FALSE       │
│   OpeningType │    │ • RoundDucts  │    │               │
│ • Threshold!  │    │   preference  │    │               │
└───────────────┘    └───────────────┘    └───────────────┘
```

**PIPES Rules:**
1. Get UI preference from CONDITIONS XML
2. Call `PipePlacementStrategy.GetResolvedOpeningType()`:
   - **Round → Rectangular threshold** (e.g., >200mm from UI setting)
   - Structural framing rule: pipes on framing → always circular

**DUCTS Rules:**
1. Check `zone.DuctShape` == "Round" or "Circular"
2. If round: Use `OpeningTypePreferences.RoundDucts` from UI

**OTHER (Cable Trays):** Always rectangular

### 4.3 Family Name Selection

```csharp
private string GetSleeveFamilyName(ClashZone zone, bool isCircular)
{
    bool isWallOrFraming = zone.StructuralElementType == "Wall" || 
                          zone.StructuralElementType == "Structural Framing";
    
    if (isWallOrFraming)
        return isCircular ? "CircularOpeningOnWall" : "RectangularOpeningOnWall";
    else // Floor
        return isCircular ? "CircularOpeningOnSlab" : "RectangularOpeningOnSlab";
}
```

---

## 5. Bulk Placement with NewFamilyInstances2

### 5.0 Pre-Flight Validation (BEFORE any transaction)

**Validate ALL data before starting Revit transactions:**

```csharp
private void ValidateBeforePlacement(List<ClashZone> zones)
{
    var errors = new List<string>();
    
    // 1. Check for invalid coordinates
    var invalidCoords = zones.Where(z => 
        double.IsNaN(z.SleevePlacementPointX) ||
        double.IsNaN(z.SleevePlacementPointY) ||
        double.IsNaN(z.SleevePlacementPointZ) ||
        double.IsInfinity(z.SleevePlacementPointX)
    ).ToList();
    
    if (invalidCoords.Any())
        errors.Add($"Invalid coordinates in {invalidCoords.Count} zones");
    
    // 2. Check for duplicate zones
    var duplicates = zones.GroupBy(z => z.ClashZoneGuid)
        .Where(g => g.Count() > 1).Count();
    if (duplicates > 0)
        errors.Add($"Duplicate zones detected: {duplicates}");
    
    // 3. Check for null/empty family names
    var missingFamilyNames = zones.Count(z => string.IsNullOrWhiteSpace(z.SleeveFamilyName));
    if (missingFamilyNames > 0)
        errors.Add($"Missing family names in {missingFamilyNames} zones");
    
    // 4. Check for invalid dimensions (sanity check)
    var invalidDims = zones.Count(z => z.SleeveWidth <= 0 || z.SleeveHeight <= 0);
    if (invalidDims > 0)
        errors.Add($"Invalid dimensions in {invalidDims} zones");
    
    // Throw combined error if any issues
    if (errors.Any())
        throw new InvalidOperationException($"Pre-flight validation failed:\n• {string.Join("\n• ", errors)}");
}
```

### 5.1 The Official API

```csharp
ICollection<ElementId> NewFamilyInstances2(IList<FamilyInstanceCreationData> dataList)
```

- Available since **Revit 2015** (confirmed for 2023)
- Single API call creates ALL instances at once
- Returns ElementIds in **same order** as input list

### 5.2 FamilyInstanceCreationData for Non-Hosted Sleeves

**Pattern 1: Individual Sleeves (ABSOLUTE Z placement)**
```csharp
// NO Level parameter - Z is absolute world coordinate
var creationData = new FamilyInstanceCreationData(
    point,                             // ABSOLUTE world coordinates
    symbol,                            // Pre-activated FamilySymbol
    StructuralType.NonStructural       // Always NonStructural for Generic Models
);
// Set Level parameter AFTER via INSTANCE_REFERENCE_LEVEL_PARAM
```

> **⚠️ DO NOT pass Level** to placement. Your database stores **absolute** Z coordinates. Passing Level would make Z relative to level elevation.

### 5.3 Level Parameter: NOT Required

> **✅ Generic Model sleeves do NOT require level association.**
> 
> - Placement uses **absolute Z coordinates** from database
> - No `INSTANCE_REFERENCE_LEVEL_PARAM` needed
> - Sleeves appear correctly in schedules without level parameter
> - Keeps implementation simple and avoids level lookup overhead


### 5.4 Symbol Pre-Activation (Validate-First Pattern)

> **⚠️ Critical:** Validate ALL families exist BEFORE starting transaction to avoid partial activation.

```csharp
private Dictionary<string, FamilySymbol> PreActivateSymbols(Document doc, List<ClashZone> zones)
{
    var uniqueFamilyNames = zones.Select(z => z.SleeveFamilyName).Distinct().ToList();
    var cache = new Dictionary<string, FamilySymbol>();
    var missingFamilies = new List<string>();
    
    // ═══════════════════════════════════════════════════════════
    // PHASE 1: VALIDATE all families exist BEFORE any transaction
    // ═══════════════════════════════════════════════════════════
    foreach (var familyName in uniqueFamilyNames)
    {
        var family = new FilteredElementCollector(doc)
            .OfClass(typeof(Family))
            .Cast<Family>()
            .FirstOrDefault(f => f.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase));
        
        if (family == null)
        {
            missingFamilies.Add(familyName);
            continue;
        }
        
        var symbolIds = family.GetFamilySymbolIds();
        if (symbolIds == null || symbolIds.Count == 0)
            missingFamilies.Add(familyName);
    }
    
    // FAIL FAST if any families missing - BEFORE any transaction
    if (missingFamilies.Count > 0)
    {
        throw new InvalidOperationException(
            $"Cannot proceed: Missing families [{string.Join(", ", missingFamilies)}]. " +
            "Load these families before placement.");
    }
    
    // ═══════════════════════════════════════════════════════════
    // PHASE 2: ACTIVATE (all validated, safe to proceed)
    // ═══════════════════════════════════════════════════════════
    using (var tx = new Transaction(doc, "Activate Family Symbols"))
    {
        tx.Start();
        
        foreach (var familyName in uniqueFamilyNames)
        {
            var family = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .First(f => f.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase));
            
            var symbolId = family.GetFamilySymbolIds().First();
            var symbol = doc.GetElement(symbolId) as FamilySymbol;
            
            if (!symbol.IsActive)
                symbol.Activate();
            
            cache[familyName] = symbol;
        }
        
        tx.Commit();
    }
    
    return cache;
}
```

### 5.5 Rotation Strategy

| Host Type | Rotation Axis | When to Apply | Source |
|-----------|---------------|---------------|--------|
| **Floor** | Z-axis (vertical) | Rectangular ducts/cable trays | `zone.MepElementRotationAngle` (radians) |
| **Wall** | Not needed | Wall sleeves align to wall | N/A - placement already correct |
| **Structural Framing** | Complex | Case-by-case | May need transformation matrix |

```csharp
private void ApplyRotation(Document doc, FamilyInstance sleeve, ClashZone zone)
{
    // Only apply rotation for floors with non-zero angle
    if (!string.Equals(zone.StructuralElementType, "Floor", StringComparison.OrdinalIgnoreCase))
        return;
    
    if (Math.Abs(zone.MepElementRotationAngle) < 0.001)
        return;  // No rotation needed
    
    try
    {
        var locationPoint = sleeve.Location as LocationPoint;
        if (locationPoint == null) return;
        
        var point = locationPoint.Point;
        var axis = Line.CreateBound(point, point.Add(XYZ.BasisZ));
        
        // MepElementRotationAngle is in RADIANS (from Revit API)
        ElementTransformUtils.RotateElement(doc, sleeve.Id, axis, zone.MepElementRotationAngle);
    }
    catch (Exception ex)
    {
        SafeFileLogger.SafeAppendText("bulk_placement.log",
            $"[{DateTime.Now:HH:mm:ss}] ⚠️ Rotation failed for zone {zone.Id}: {ex.Message}\n");
    }
}
```

### 5.6 Complete Implementation

```csharp
public void ExecuteBulkPlacement(Document doc)
{
    // STEP 1: Query DB
    var zones = _repository.GetZonesReadyForPlacement();
    if (zones.Count == 0) return;
    
    // STEP 2: Calculate sizes using CURRENT UI settings
    CalculateSizesAndFamilyNames(zones);
    
    // STEP 3: Save to DB
    _repository.BatchUpdateSleeveSizes(zones);
    
    // STEP 4: Pre-activate symbols
    var symbolCache = PreActivateSymbols(doc, zones);
    
    // STEP 5: Build creation data
    var creationDataList = zones.Select(z => new FamilyInstanceCreationData(
        new XYZ(z.SleevePlacementPointX, z.SleevePlacementPointY, z.SleevePlacementPointZ),
        symbolCache[z.SleeveFamilyName],
        StructuralType.NonStructural
    )).ToList();
    
    // STEP 6: Single Transaction - Bulk Placement
    using (var tx = new Transaction(doc, "Bulk Sleeve Placement"))
    {
        tx.Start();
        
        var createdIds = doc.Create.NewFamilyInstances2(creationDataList);
        
        var idList = createdIds.ToList();
        for (int i = 0; i < idList.Count; i++)
        {
            var sleeve = doc.GetElement(idList[i]) as FamilyInstance;
            var zone = zones[i];
            
            // Set parameters
            sleeve.LookupParameter("Width")?.Set(zone.SleeveWidth);
            sleeve.LookupParameter("Height")?.Set(zone.SleeveHeight);
            if (zone.SleeveDiameter > 0)
                sleeve.LookupParameter("Diameter")?.Set(zone.SleeveDiameter);
            
            // Apply rotation for floors
            if (zone.StructuralElementType == "Floor" && Math.Abs(zone.MepElementRotationAngle) > 0.001)
            {
                var point = (sleeve.Location as LocationPoint).Point;
                var axis = Line.CreateBound(point, point.Add(XYZ.BasisZ));
                ElementTransformUtils.RotateElement(doc, sleeve.Id, axis, zone.MepElementRotationAngle);
            }
            
            zone.SleeveInstanceId = sleeve.Id.IntegerValue;
        }
        
        doc.Regenerate();
        tx.Commit();
    }
    
    // STEP 7: Batch update DB
    _repository.BatchUpdateAfterPlacement(zones);
}
```

---

## 6. SOLID Architecture & Crash Recovery

### 6.1 Service Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│               IBulkPlacementService (Orchestrator)              │
└─────────────────────────────────────────────────────────────────┘
        ┌─────────────────────┼─────────────────────┐
        ▼                     ▼                     ▼
┌───────────────┐    ┌───────────────┐    ┌───────────────┐
│ISizeCalculator│    │ISymbolManager │    │IPlacementExec │
└───────────────┘    └───────────────┘    └───────────────┘
```

### 6.2 Transaction Safety

Use `TransactionGroup` for atomic rollback:

```csharp
using (var transactionGroup = new TransactionGroup(doc, "Bulk Sleeve Placement"))
{
    transactionGroup.Start();
    
    try
    {
        using (var tx = new Transaction(doc, "Place All Sleeves"))
        {
            tx.Start();
            // ... bulk placement logic ...
            tx.Commit();
        }
        transactionGroup.Assimilate();  // Success
    }
    catch (Exception)
    {
        transactionGroup.RollBack();    // Undo ALL Revit changes
        throw;
    }
}
```

### 6.3 Comprehensive Crash Recovery Table

| Phase | Failure Type | Revit State | DB State | Recovery Action |
|-------|-------------|-------------|----------|-----------------|
| **Pre-flight** | Validation fails | Unchanged | Unchanged | Fix data, retry |
| **Step 2** | Calculation fails | Unchanged | Unchanged | Fix logic, retry |
| **Step 3** | DB save fails | Unchanged | Rolled back | Check DB, retry |
| **Step 4** | Family missing | Unchanged | Sizes saved | Load family, retry |
| **Step 4** | Activation fails | **Partial** | Sizes saved | TransactionGroup rollback |
| **Step 6** | Placement fails | **Partial sleeves** | Sizes saved | TransactionGroup rollback → clean |
| **Step 6** | Parameter fails | Sleeves exist | Sizes saved | Log warning, continue |
| **Step 7** | DB update fails (retry 1-2) | Sleeves exist | Stale | Automatic retry |
| **Step 7** | DB update fails (retry 3) | Sleeves exist | Stale | Recovery JSON saved |
| **Hard crash** | Revit crashes | Unknown | Sizes saved | Reopen, check, retry |

### 6.4 Error Handling Matrix

| Exception Type | Phase | Recovery | User Message |
|----------------|-------|----------|--------------|
| `InvalidOperationException` | Symbol activation | Skip symbol, continue | "Family X not loaded" |
| `ArgumentNullException` | Placement data | Fail transaction | "Invalid data for zone Y" |
| `Autodesk.Revit.Exceptions.ArgumentException` | Parameter setting | Log, continue | "Could not set parameters" |
| Database `SqliteException` | DB operations | Retry 3x, then fail | "Database error" |

### 6.5 DB Update Retry Logic with Recovery File

```csharp
private void BatchUpdateDatabaseWithRetry(List<ClashZone> placedZones)
{
    int retryCount = 0;
    const int maxRetries = 3;
    
    while (retryCount < maxRetries)
    {
        try
        {
            _repository.BatchUpdateAfterPlacement(placedZones);
            return;  // Success
        }
        catch (Exception ex)
        {
            retryCount++;
            SafeFileLogger.SafeAppendText("bulk_placement.log",
                $"[{DateTime.Now:HH:mm:ss}] ⚠️ DB update attempt {retryCount} failed: {ex.Message}\n");
            
            if (retryCount >= maxRetries)
            {
                // Save recovery file for manual intervention
                var recoveryData = placedZones.Select(z => new {
                    ZoneGuid = z.ClashZoneGuid,
                    SleeveInstanceId = z.SleeveInstanceId
                });
                
                string recoveryPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "JSE_MEP_Openings", "Logs",
                    $"placement_recovery_{DateTime.Now:yyyyMMdd_HHmmss}.json"
                );
                
                File.WriteAllText(recoveryPath, JsonConvert.SerializeObject(recoveryData, Formatting.Indented));
                
                throw new Exception(
                    $"Database update failed after {maxRetries} attempts. " +
                    $"Recovery file saved to: {recoveryPath}", ex);
            }
            
            Thread.Sleep(500 * retryCount);  // Exponential backoff
        }
    }
}
```

### 6.6 Partial Success Handling

```csharp
public class BulkPlacementResult
{
    public bool OverallSuccess { get; set; }
    public int TotalZones { get; set; }
    public int PlacedCount { get; set; }
    public int SkippedCount { get; set; }
    public int FailedCount { get; set; }
    public string? Error { get; set; }
    public string? Warning { get; set; }
    public List<(Guid ZoneId, string Reason)> Failures { get; set; } = new();
    
    public string GetSummary()
    {
        if (OverallSuccess)
            return $"✅ Successfully placed {PlacedCount}/{TotalZones} sleeves";
        else if (PlacedCount > 0)
            return $"⚠️ Partial success: {PlacedCount} placed, {FailedCount} failed. {Error}";
        else
            return $"❌ Failed: {Error}";
    }
}
```

---

## 7. Performance Optimization

### 7.1 Performance Targets

| Metric | Target |
|--------|--------|
| **200 sleeves** | < 30 seconds |
| **500 sleeves** | < 1 minute |
| **Size calculation** | < 500ms (sequential) |

### 7.2 Optimization Strategies

1. **Single Transaction** - All placements in one transaction
2. **Deferred Regeneration** - Only regenerate once at the end
3. **Symbol Caching** - Activate each unique symbol once
4. **Batch DB Operations** - Single query, single update
5. **ORDER BY in query** - Group by category for symbol reuse

### 7.3 Multi-Threading (Optional, >500 zones only)

```csharp
// Only parallelize for large batches (>500 zones)
if (OptimizationFlags.UseSOLIDRefactoredIndividualPreCalculation && zones.Count > 500)
{
    Parallel.ForEach(
        Partitioner.Create(0, zones.Count),
        new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount / 2 },
        range => {
            for (int i = range.Item1; i < range.Item2; i++)
                CalculateZoneSizeAndFamily(zones[i], settings);
        });
}
```

**Safety:** Each thread processes its own index range - no shared state mutation.

---

## 8. Implementation Checklist

### Phase 1: Repository Methods
- [ ] Add `GetZonesReadyForPlacement()` to `ClashZoneRepository`
- [ ] Add `BatchUpdateSleeveSizes(zones)` for Step 3
- [ ] Add `BatchUpdateAfterPlacement(zones)` for Step 7

### Phase 2: Bulk Placement Service
- [ ] Create `IBulkPlacementService` interface
- [ ] Implement `SafeBulkPlacementService` with 7-step flow
- [ ] Add `BulkPlacementResult` result class

### Phase 3: Integration
- [ ] Add "Place All Sleeves" button/command
- [ ] Wire up to existing UI
- [ ] Add feature flag: `UseBulkPlacement`

### Phase 4: Testing
- [ ] Test with 50 zones (quick)
- [ ] Test with 200 zones (typical)
- [ ] Test with 500+ zones (large project)
- [ ] Test BIM 360 performance
- [ ] Test rollback on failure

---

## End of Document
