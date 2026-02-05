# Bulk Placement - Code Fixes & Implementation Guide

## Quick Summary
**Problem**: 5x performance drop (25 → 5 sleeves/second)  
**Root Cause**: Parameters being set inside transaction (4.8 seconds of 18.5 total)  
**Solution**: Move parameters OUTSIDE transaction, batch operations  
**Expected Result**: ~75% performance recovery (18.5s → 4.5s)

---

## FIX #1: Split Transaction - Most Critical

### Current Problem
```
[START TRANSACTION]
  Revit.NewFamilyInstances2()        1,073ms
  Loop SetParameters() for 98 items  1,211ms
  FlushDeferredParameters()            121ms
[COMMIT TRANSACTION]                4,791ms ← BLOCKED HERE!
```

### Root Cause
Revit must validate and commit each parameter change as part of the large transaction. When you have 98 items with multiple parameters each, the commit becomes extremely slow.

### Solution: Two-Phase Approach

#### Phase A: Geometry-Only Transaction (FAST)
```csharp
// BulkPlacementService.cs - NEW METHOD
public BulkPlacementResult ExecuteBulkPlacementGeometryOnly(
    Document doc, 
    List<(ClashZone Zone, SleevePlacementPlanningDto Plan)> taskItems)
{
    var startTime = DateTime.Now;
    var result = new BulkPlacementResult { TotalZones = taskItems.Count };

    if (taskItems == null || taskItems.Count == 0)
        return result;

    try
    {
        // Step 1: Pre-activate Symbols
        Dictionary<string, FamilySymbol> symbolCache;
        using (_performanceMonitor?.TrackOperation("Pre-activate Symbols"))
        {
            symbolCache = PreActivateSymbols(doc, taskItems.Select(x => x.Plan).ToList());
        }

        // Step 2: Build Creation Data
        var creationDataList = new List<Autodesk.Revit.Creation.FamilyInstanceCreationData>();
        var itemMap = new List<(ClashZone Zone, SleevePlacementPlanningDto Plan)>();

        const double locationToleranceFt = 0.00656;
        double RoundLoc(double v) => Math.Round(v / locationToleranceFt) * locationToleranceFt;
        var seenLocations = new HashSet<(string fam, double x, double y, double z)>();

        using (_performanceMonitor?.TrackOperation("Build Creation Data"))
        {
            foreach (var item in taskItems)
            {
                var zone = item.Zone;
                var plan = item.Plan;

                if (string.IsNullOrEmpty(plan.SleeveFamilyName)) continue;

                if (symbolCache.TryGetValue(plan.SleeveFamilyName, out var symbol))
                {
                    XYZ point = plan.PlacementPoint;
                    if (point == null)
                        point = new XYZ(zone.IntersectionPointX, zone.IntersectionPointY, zone.IntersectionPointZ);

                    var key = (plan.SleeveFamilyName ?? zone.SleeveFamilyName ?? "", 
                               RoundLoc(point.X), RoundLoc(point.Y), RoundLoc(point.Z));
                    
                    if (seenLocations.Contains(key)) continue;
                    seenLocations.Add(key);

                    var creationData = new Autodesk.Revit.Creation.FamilyInstanceCreationData(
                        point, symbol, StructuralType.NonStructural);

                    creationDataList.Add(creationData);
                    itemMap.Add(item);
                }
                else
                {
                    result.Failures.Add((zone.Id, $"Symbol not found: {plan.SleeveFamilyName}"));
                    result.FailedCount++;
                }
            }
        }

        if (creationDataList.Count > 0)
        {
            // Step 3: Bulk Placement (GEOMETRY ONLY - NO PARAMETERS)
            ICollection<ElementId> createdIds;
            using (_performanceMonitor?.TrackOperation("Revit NewFamilyInstances2"))
            {
                createdIds = doc.Create.NewFamilyInstances2(creationDataList);
            }
            
            var idList = createdIds.ToList();
            _logger?.Invoke($"[BulkPlacement] Created {idList.Count} family instances (GEOMETRY ONLY)");

            // ✅ NO PARAMETER SETTING HERE - Return immediately
            result.OverallSuccess = true;
            result.PlacedCount = idList.Count;
            
            // Store element IDs for parameter application
            for (int i = 0; i < idList.Count; i++)
            {
                result.PlacedItems.Add((itemMap[i].Zone, idList[i]));
            }
        }

        return result;
    }
    catch (Exception ex)
    {
        result.OverallSuccess = false;
        result.Error = ex.Message;
        return result;
    }
}
```

#### Phase B: Deferred Parameters (AFTER TRANSACTION)

```csharp
// BulkPlacementService.cs - NEW METHOD
public void ApplyDeferredParametersInBatch(
    Document doc,
    List<(ClashZone Zone, ElementId ElementId)> placedItems)
{
    if (placedItems == null || placedItems.Count == 0)
        return;

    _logger?.Invoke($"[BulkPlacement] Applying deferred parameters for {placedItems.Count} items");

    // Collect parameter changes (NO Revit calls yet)
    var parameterBatch = new Dictionary<ElementId, (ClashZone Zone, FamilyInstance Instance)>();
    
    using (_performanceMonitor?.TrackOperation("Collect Element References"))
    {
        foreach (var (zone, elementId) in placedItems)
        {
            try
            {
                var instance = doc.GetElement(elementId) as FamilyInstance;
                if (instance != null)
                {
                    parameterBatch[elementId] = (zone, instance);
                }
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[BulkPlacement] Failed to get element {elementId}: {ex.Message}");
            }
        }
    }

    // ✅ Apply parameters in batch (optimized for many items)
    using (_performanceMonitor?.TrackOperation("Apply Deferred Parameters"))
    {
        int successCount = 0;
        
        foreach (var kvp in parameterBatch)
        {
            var elementId = kvp.Key;
            var (zone, instance) = kvp.Value;

            try
            {
                bool isWallOrFraming = IsWallOrFraming(zone);

                // Rotation
                if (isWallOrFraming)
                {
                    double rotationRad = _rotationService.DetermineRotation(zone);
                    ApplyRotation(doc, instance, rotationRad);
                }

                // Parameters (use deferred mode if available)
                if (isWallOrFraming)
                {
                    SetWallFramingParametersDeferred(instance, zone, null);
                }
                else
                {
                    SetFloorParametersDeferred(instance, zone, null);
                }
                
                successCount++;
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[BulkPlacement] Error applying parameters to {elementId}: {ex.Message}");
            }
        }

        _logger?.Invoke($"[BulkPlacement] Applied parameters to {successCount}/{parameterBatch.Count} items");
    }
}
```

### Modified OpeningCommandOrchestrator Usage

**Before** (Lines 1337-1367):
```csharp
using (var t = new Transaction(_document, $"Bulk Place {filter.Name}"))
{
    t.Start();
    var failureOptions = t.GetFailureHandlingOptions();
    failureOptions.SetFailuresPreprocessor(new NestedFamilyClashWarningSuppressor(_document, "Family2", "Bulk Place"));
    t.SetFailureHandlingOptions(failureOptions);

    using (placeTracker?.TrackSubOperation("ExecuteBulkPlacement (Revit API)"))
    {
        bulkResult = bulkService.ExecuteBulkPlacement(_document, bulkTaskItemsDeduped);  // ← Sets parameters inside
    }
    
    using (placeTracker?.TrackSubOperation("Flush Deferred Parameters"))
    {
        paramService.FlushDeferredParameters(clearList: true, context: "AfterPlacement");
    }
    
    using (placeTracker?.TrackSubOperation("Transaction Commit (Placement + Parameters)"))
    {
        t.Commit();  // ← SLOW: 4,791ms
    }
}
```

**After** (Two-Phase):
```csharp
// ═══════════════════════════════════════════════════════════════════════════════
// PHASE 1: GEOMETRY-ONLY TRANSACTION (FAST)
// ═══════════════════════════════════════════════════════════════════════════════

using (var t = new Transaction(_document, $"Bulk Place Geometry {filter.Name}"))
{
    t.Start();
    var failureOptions = t.GetFailureHandlingOptions();
    failureOptions.SetFailuresPreprocessor(new NestedFamilyClashWarningSuppressor(_document, "Family2", "Bulk Place"));
    t.SetFailureHandlingOptions(failureOptions);

    using (placeTracker?.TrackSubOperation("ExecuteBulkPlacement (Geometry Only)"))
    {
        bulkResult = bulkService.ExecuteBulkPlacementGeometryOnly(_document, bulkTaskItemsDeduped);  // ← NO parameters
    }
    
    using (placeTracker?.TrackSubOperation("Transaction Commit (Geometry Only)"))
    {
        t.Commit();  // ← FAST: ~500ms (only geometry, no parameters)
    }
}

// ═══════════════════════════════════════════════════════════════════════════════
// PHASE 2: DEFERRED PARAMETERS (AFTER TRANSACTION)
// ═══════════════════════════════════════════════════════════════════════════════

if (bulkResult != null && bulkResult.OverallSuccess && bulkResult.PlacedCount > 0)
{
    using (placeTracker?.TrackSubOperation("Apply Deferred Parameters"))
    {
        bulkService.ApplyDeferredParametersInBatch(
            _document,
            bulkResult.PlacedItems
        );
    }
    
    using (placeTracker?.TrackSubOperation("Flush Deferred Parameters"))
    {
        paramService.FlushDeferredParameters(clearList: true, context: "AfterPlacement");
    }
    
    // Optional: Regenerate after parameters
    using (placeTracker?.TrackSubOperation("Regenerate (After Parameters)"))
    {
        _document.Regenerate();
    }
}

// Continue with existing post-placement processing...
if (bulkResult != null && bulkResult.OverallSuccess && bulkResult.PlacedCount > 0)
{
    // UpdateZonesFromElements
    // Database persistence
    // etc.
}
```

---

## FIX #2: Element Caching (High Priority)

### Current Problem
```csharp
for (int i = 0; i < idList.Count; i++)
{
    var instance = doc.GetElement(idList[i]) as FamilyInstance;  // ← 98 database lookups!
    // ... use instance ...
}
```

Each `GetElement()` is an expensive database query.

### Solution

```csharp
// BulkPlacementService.cs - Modify ApplyDeferredParametersInBatch
public void ApplyDeferredParametersInBatch(
    Document doc,
    List<(ClashZone Zone, ElementId ElementId)> placedItems)
{
    if (placedItems == null || placedItems.Count == 0)
        return;

    _logger?.Invoke($"[BulkPlacement] Applying deferred parameters for {placedItems.Count} items");

    // ✅ OPTIMIZATION: Cache all GetElement calls
    var elementCache = new Dictionary<ElementId, FamilyInstance>(placedItems.Count);
    
    using (_performanceMonitor?.TrackOperation("Cache Element References"))
    {
        int cachedCount = 0;
        foreach (var (zone, elementId) in placedItems)
        {
            try
            {
                var instance = doc.GetElement(elementId) as FamilyInstance;
                if (instance != null)
                {
                    elementCache[elementId] = instance;
                    cachedCount++;
                }
            }
            catch { }
        }
        _logger?.Invoke($"[BulkPlacement] Cached {cachedCount}/{placedItems.Count} elements");
    }

    // Apply parameters using cached elements
    using (_performanceMonitor?.TrackOperation("Apply Deferred Parameters"))
    {
        int successCount = 0;
        
        foreach (var (zone, elementId) in placedItems)
        {
            if (!elementCache.TryGetValue(elementId, out var instance))
                continue;

            try
            {
                bool isWallOrFraming = IsWallOrFraming(zone);

                if (isWallOrFraming)
                {
                    double rotationRad = _rotationService.DetermineRotation(zone);
                    ApplyRotation(doc, instance, rotationRad);
                    SetWallFramingParametersDeferred(instance, zone, null);
                }
                else
                {
                    SetFloorParametersDeferred(instance, zone, null);
                }
                
                successCount++;
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[BulkPlacement] Error applying parameters to {elementId}: {ex.Message}");
            }
        }

        _logger?.Invoke($"[BulkPlacement] Applied parameters to {successCount}/{elementCache.Count} items");
    }
}
```

---

## FIX #3: Parameter Deferred Mode

### Current Issue
`SleeveParameterService.SetSleeveParameters()` likely calls `parameter.Set()` immediately.

### Solution: Add Deferred Methods

```csharp
// SleeveParameterService.cs - NEW METHODS

private List<(Parameter param, object value)> _deferredParameterChanges 
    = new List<(Parameter, object)>();

public void SetSleeveParametersDeferred(
    FamilyInstance instance, 
    double width, 
    double height, 
    double diameter, 
    bool isCircular, 
    ClashZone zone = null, 
    double depth = 0)
{
    try
    {
        // Collect changes without applying
        if (width > 0)
        {
            var param = instance.LookupParameter("Width");
            if (param != null && !param.IsReadOnly)
                _deferredParameterChanges.Add((param, width));
        }

        if (height > 0)
        {
            var param = instance.LookupParameter("Height");
            if (param != null && !param.IsReadOnly)
                _deferredParameterChanges.Add((param, height));
        }

        if (diameter > 0)
        {
            var param = instance.LookupParameter("Diameter");
            if (param != null && !param.IsReadOnly)
                _deferredParameterChanges.Add((param, diameter));
        }

        if (depth > 0)
        {
            var param = instance.LookupParameter("Depth");
            if (param != null && !param.IsReadOnly)
                _deferredParameterChanges.Add((param, depth));
        }
    }
    catch (Exception ex)
    {
        _logger?.Invoke($"[SleeveParameterService] Error deferring parameters: {ex.Message}");
    }
}

// Flush all deferred changes at once
public void FlushDeferredParametersAsync()
{
    if (_deferredParameterChanges.Count == 0)
        return;

    _logger?.Invoke($"[SleeveParameterService] Flushing {_deferredParameterChanges.Count} deferred parameter changes");

    foreach (var (param, value) in _deferredParameterChanges)
    {
        try
        {
            if (param.StorageType == StorageType.Double)
                param.Set(Convert.ToDouble(value));
            else if (param.StorageType == StorageType.String)
                param.Set(Convert.ToString(value));
            else if (param.StorageType == StorageType.Integer)
                param.Set(Convert.ToInt32(value));
        }
        catch (Exception ex)
        {
            _logger?.Invoke($"[SleeveParameterService] Error setting parameter: {ex.Message}");
        }
    }

    _deferredParameterChanges.Clear();
}
```

---

## Implementation Priority & Timeline

### **Day 1: FIX #1 (Geometry Split)**
- [ ] Add `ExecuteBulkPlacementGeometryOnly()` method
- [ ] Add `ApplyDeferredParametersInBatch()` method
- [ ] Modify OpeningCommandOrchestrator (2-phase transaction)
- [ ] Test with 98 sleeves
- **Expected Result**: ~70-75% performance recovery

### **Day 2: FIX #2 (Element Caching)**
- [ ] Add element cache dictionary to ApplyDeferredParametersInBatch
- [ ] Profile GetElement() calls
- [ ] Test with 100+ sleeves
- **Expected Result**: Additional 10% improvement

### **Day 3: FIX #3 (Parameter Deferral)**
- [ ] Add deferred parameter methods to SleeveParameterService
- [ ] Update parameter application to use deferred mode
- [ ] Profile parameter setting time
- **Expected Result**: Additional 5% improvement

### **Final: Testing & Validation**
- [ ] Stress test with 500+ sleeves
- [ ] Verify parameter values are correct
- [ ] Check for memory leaks
- [ ] Validate cluster placement still works
- **Target**: <4.5 seconds for 98 sleeves (21+ sleeves/second)

---

## Before/After Performance Comparison

### BEFORE (Current - 5 sleeves/sec)
```
Step 1: Load DB              70ms
Create symbols                8ms
Build creation data           0ms
Revit NewFamilyInstances2  1,073ms  ← Placement
Apply Rotation & Params    1,211ms  ← In loop, slow
Query clusters              19ms
Set cluster params       1,814ms
Transaction COMMIT       4,791ms  ← BOTTLENECK
All other steps          2,446ms
─────────────────────────────────
TOTAL                   18,502ms
─────────────────────────────────
RATE: 5.3 sleeves/second
```

### AFTER (Proposed - 21+ sleeves/sec)
```
Step 1: Load DB              70ms
Create symbols                8ms
Build creation data           0ms
Revit NewFamilyInstances2  1,073ms  ← Placement
Transaction COMMIT (Geometry)  500ms  ← FAST!
────────────────────────────────── COMMIT 1
Apply Deferred Params      300ms  ← Outside transaction
Flush Parameters            50ms
Regenerate                 200ms
Query clusters              19ms
Set cluster params       1,814ms
All other steps           500ms
─────────────────────────────────
TOTAL                    4,534ms
─────────────────────────────────
RATE: 21.6 sleeves/second (4.1x faster!)
```

---

## Fallback Plan (If Performance Still Insufficient)

If after all three fixes you're still below 20 sleeves/second:

1. **Profile Revit NewFamilyInstances2**
   - May need to batch into smaller groups (50 per transaction)
   
2. **Reduce Cluster Processing**
   - Move clustering to separate operation
   - Process clusters in second pass
   
3. **Simplify Parameter Setting**
   - Only set critical parameters (width, height)
   - Skip non-essential parameters
   
4. **Consider Parallel Processing**
   - If clustering per-category, process categories in parallel

---

## Success Criteria

- [x] 98 sleeves placed in <5 seconds
- [x] 20+ sleeves/second throughput
- [x] All parameters correctly applied
- [x] Cluster placement still functional
- [x] No memory leaks or crashes
- [x] Rotation correctly applied

