# Multi-Floor BIM360 Optimization Plan

## Executive Summary

**Current Issue:** The multi-story feature processes each floor in a **separate transaction**, causing N transactions for N floors. In BIM360, each transaction requires a cloud sync round-trip, creating significant latency.

**Solution:** Aggregate all floors' sleeve placement data and execute in **ONE transaction** with **ONE batch API call**.

---

## Current Implementation Analysis

### 1. FloorBatchProcessor (Current)
```csharp
// Line 114-143: Processes floors sequentially, ONE transaction per floor
foreach (var level in chunk)
{
    var floorProcessor = new SingleFloorProcessor(_doc, cache, _monitor);
    var res = floorProcessor.ProcessFloor(level, filter); // ← ONE transaction per floor
}
```

### 2. SingleFloorProcessor.ProcessFloor (Current)
```csharp
// Line 39-162: Creates a transaction for EACH floor
public FloorProcessingResult ProcessFloor(Level level, OpeningFilter filter)
{
    using (var transaction = new Transaction(_doc, $"Process Floor {level.Name}"))
    {
        transaction.Start();
        // Detect clashes → Place sleeves → Commit
        transaction.Commit(); // ← BIM360 sync happens here
    }
}
```

### 3. BulkPlacementService (Already Optimal)
```csharp
// Line 551-558: Uses NewFamilyInstances2 (BATCH API)
ICollection<ElementId> placedIds = doc.Create.NewFamilyInstances2(creationDataList);
```

---

## Optimization Strategy

### Phase 1: Multi-Floor Batch Placement (HIGH PRIORITY)

**Change from:**
```
FOR each floor:
  START transaction
  detect clashes
  plan sleeves
  place sleeves (batch per floor)
  COMMIT transaction
```

**Change to:**
```
START transaction
FOR each floor:
  detect clashes (no transaction needed - read-only)
  plan sleeves (no transaction needed - calculation only)
  
AGGREGATE all planned sleeves from all floors
place ALL sleeves in ONE batch API call
COMMIT transaction
```

### Implementation Code

```csharp
public class MultiFloorBatchPlacementService
{
    private readonly Document _doc;
    private readonly IPerformanceMonitor _monitor;
    
    /// <summary>
    /// 🚀 OPTIMIZED: Places sleeves for ALL floors in ONE transaction
    /// </summary>
    public MultiFloorResult PlaceAllFloorsInSingleTransaction(
        List<Level> levels, 
        OpeningFilter filter,
        SharedResourceCache cache)
    {
        var result = new MultiFloorResult();
        var allPlannedItems = new List<(ClashZone Zone, SleevePlacementPlanningDto Plan)>();
        
        // ===== PHASE 1: DETECTION & PLANNING (NO TRANSACTION) =====
        // Read-only operations don't need Revit transaction
        foreach (var level in levels)
        {
            try
            {
                // Detect clashes (read-only)
                var detector = new FloorIsolatedClashDetector(_doc, cache);
                var categories = GetSelectedCategories(filter);
                var clashes = detector.DetectClashesForFloor(level, categories);
                
                if (clashes.Count == 0) continue;
                
                // Plan sleeves (calculation only, no Revit API)
                var plannedItems = PlanSleevesForFloor(clashes, filter);
                allPlannedItems.AddRange(plannedItems);
                
                result.SuccessfulFloors.Add(level.Name);
            }
            catch (Exception ex)
            {
                result.FailedFloors.Add($"{level.Name}: {ex.Message}");
            }
        }
        
        if (allPlannedItems.Count == 0)
        {
            return result;
        }
        
        // ===== PHASE 2: SINGLE TRANSACTION PLACEMENT =====
        using (var transaction = new Transaction(_doc, "Place Sleeves - All Floors"))
        {
            transaction.Start();
            
            try
            {
                // Place ALL sleeves across ALL floors in ONE batch
                var placementService = new BulkPlacementService(
                    _doc,
                    contextFactory: () => new SleeveDbContext(_doc),
                    logger: msg => SafeFileLogger.SafeAppendText("multifloor.log", msg),
                    performanceMonitor: _monitor);
                
                // ✅ SINGLE API CALL for all floors
                var placementResult = placementService.ExecuteBulkPlacement(
                    _doc, 
                    allPlannedItems, 
                    skipSpatialFiltering: false);
                
                result.TotalSleevesPlaced = placementResult.PlacedCount;
                
                transaction.Commit();
                
                SafeFileLogger.SafeAppendText("multifloor.log",
                    $"[{DateTime.Now}] ✅ Multi-floor batch complete: {placementResult.PlacedCount} sleeves placed across {levels.Count} floors in ONE transaction");
            }
            catch (Exception ex)
            {
                transaction.RollBack();
                throw;
            }
        }
        
        return result;
    }
}
```

---

## Phase 2: Global Clustering After Multi-Floor Placement

**Current:** Clustering runs per-floor (if enabled)
**Optimized:** Run clustering ONCE after all floors are placed

```csharp
public class GlobalClusteringOrchestrator
{
    /// <summary>
    /// Runs clustering for ALL floors in single pass
    /// </summary>
    public void ExecuteGlobalClustering(Document doc, List<string> floorNames)
    {
        using (var transaction = new Transaction(doc, "Global Clustering - All Floors"))
        {
            transaction.Start();
            
            // 1. Load ALL resolved zones across ALL floors
            var allZones = repo.GetZonesReadyForProximityCheck(floorNames);
            
            // 2. Run proximity check globally (cross-floor sleeves won't cluster due to Z-distance)
            var proximityMarker = new ClusterProximityMarker(repo, proximityTolerance, _logger);
            int marked = proximityMarker.MarkZonesForClustering(doc, null); // null = all categories
            
            // 3. Calculate clusters for marked zones
            var zonesForClustering = repo.GetZonesForClustering();
            var results = calcService.CalculateOnly(zonesForClustering, "All", 0, 0, doc);
            
            // 4. Place all clusters in ONE batch
            var (placed, failed, cleaned) = placeService.PlaceAllCategoriesAndCleanup(doc, cleanup);
            
            transaction.Commit();
        }
    }
}
```

---

## Performance Comparison

| Scenario | Current (Per-Floor) | Optimized (Batch) | Improvement |
|----------|---------------------|-------------------|-------------|
| **20 Floors, 50 sleeves each** | 20 transactions, 20 batch calls | 1 transaction, 1 batch call | **20x fewer syncs** |
| **BIM360 Latency** | 20 × 500ms = 10 seconds | 1 × 500ms = 0.5 seconds | **95% faster** |
| **Revit API Calls** | 20 transactions | 1 transaction | **95% reduction** |
| **Memory Usage** | Similar | Similar | No change |

---

## Implementation Checklist

### Week 1: Core Multi-Floor Batch Service
- [ ] Create `MultiFloorBatchPlacementService`
- [ ] Separate detection/planning from placement
- [ ] Modify `FloorBatchProcessor` to use single transaction
- [ ] Test with 10+ floors

### Week 2: Global Clustering
- [ ] Create `GlobalClusteringOrchestrator`
- [ ] Run clustering after all floors placed
- [ ] Update database flags for multi-floor
- [ ] Test clustering across floors

### Week 3: BIM360 Testing
- [ ] Test in BIM360 environment
- [ ] Measure transaction timing
- [ ] Verify cloud sync behavior
- [ ] Optimize chunk sizes if needed

---

## Code Changes Required

### Files to Modify:

1. **Services/MultiFloor/FloorBatchProcessor.cs**
   - Replace per-floor transaction loop
   - Aggregate planned items before placement
   - Single transaction for placement phase

2. **Services/MultiFloor/SingleFloorProcessor.cs**
   - Split into detection/planning (no transaction) and placement (transaction)
   - Or create new `FloorDetectionService` for read-only operations

3. **Services/Placement/PlacementWorkflowOrchestrator.cs**
   - Support multi-floor detection → single placement flow
   - Global clustering after multi-floor placement

4. **Services/Clustering/BatchClusterPlacementService.cs**
   - Ensure clustering works with multi-floor data
   - Single transaction for all cluster placement

---

## Risk Mitigation

| Risk | Mitigation |
|------|------------|
| **Too many sleeves in one transaction** | Add chunking: 500 sleeves max per batch |
| **Memory pressure** | Stream clash detection, don't hold all in memory |
| **One failure fails all floors** | Try-catch per floor during detection, continue on error |
| **Transaction timeout** | Add 5-minute timeout, commit partial if needed |

---

## Success Metrics

- [ ] 20 floors processed in < 1 transaction (vs 20)
- [ ] BIM360 sync time reduced by 90%+
- [ ] No regression in sleeve placement accuracy
- [ ] Clustering works correctly across floors
