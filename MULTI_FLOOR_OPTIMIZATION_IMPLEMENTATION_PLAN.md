# Multi-Floor Optimization Implementation Plan

## Objective
Reduce transaction time by minimizing the number of transaction commits (each commit = implicit regen + Revit overhead + BIM360 sync).

## Current Architecture (Per-Chunk, Multi-Floor)

### Individual Placement (N chunks = N transactions)
```
Chunk 1: [Transaction Start → NewFamilyInstances2 → Regen → Rotate → Params → Flush → Commit]
Chunk 2: [Transaction Start → NewFamilyInstances2 → Regen → Rotate → Params → Flush → Commit]
Chunk N: [Transaction Start → NewFamilyInstances2 → Regen → Rotate → Params → Flush → Commit]
→ UpdateZonesFromElements (reads placement point, bbox, params)
→ ExtractAndSaveCornersForZones (reads 4 corners to DB)
```
**Result**: N explicit regens + N commit regens = **2N regenerations**

### Clustering (1 transaction — already global)
```
Transaction Start
  → NewFamilyInstances2 (cluster sleeves)
  → Rotation (pre-calculated axis, no regen needed)
  → Params + Flush
  → doc.Delete(individual sleeves)     ← INSIDE transaction (adds weight)
  → Pre-read Revit params
  → Commit                            ← implicit regen (heavy: placement + deletion)
→ UpdateClusterBoundingBoxesAfterPlacement (reads bbox after commit regen)
→ ExtractAndSaveCornersForClusters (reads corners after commit regen)
```
**Result**: 0 explicit + 1 commit regen = **1 regeneration** (but heavy transaction)

**Current total for 10 chunks**: 2×10 + 1 = **21 regenerations**

---

## Proposed Architecture (Optimized)

### Optimization 1: Individual Placement — Consolidate N chunks into 1 transaction

```
ONE Transaction:
  Batch 1: NewFamilyInstances2 → collect instances
  Batch 2: NewFamilyInstances2 → collect instances
  Batch N: NewFamilyInstances2 → collect instances
  → ONE doc.Regenerate()           ← all instances get valid Location at once
  → Apply ALL rotations            ← uses valid Location for axis
  → Set ALL parameters + Flush
  → Commit                         ← 1 implicit regen
→ UpdateZonesFromElements          ← reads geometry after commit (existing code, no changes)
→ ExtractAndSaveCornersForZones    ← reads corners after commit (existing code, no changes)
```
**Result**: 1 explicit + 1 commit = **2 regenerations** (down from 2N)

### Optimization 2: Clustering — Move deletion out of cluster transaction

```
Transaction 1 (Cluster Placement — lighter, commits faster):
  → NewFamilyInstances2 (cluster sleeves)
  → Rotation (pre-calculated axis)
  → Params + Flush
  → Pre-read Revit params
  → Commit                         ← implicit regen (lighter: no deletion in this tx)
→ UpdateClusterBoundingBoxesAfterPlacement  ← reads bbox after commit regen
→ ExtractAndSaveCornersForClusters          ← reads corners after commit regen

Transaction 2 (Deletion — separate, independent):
  → doc.Delete(old individual sleeves)
  → Commit                         ← implicit regen
```
**Result**: 2 commit regens (but each transaction is lighter and commits faster)

**Optimized total**: 2 (individual) + 2 (clustering) = **4 regenerations**
But each regen/commit is faster because transactions are lighter.

---

## Why Each Regen MUST Stay

| Regen | Location | Why it CANNOT be removed |
|-------|----------|--------------------------|
| Explicit after NewFamilyInstances2 | BulkPlacementService.cs:564 | Without this, `instance.Location` returns `(0,0,0)`. Rotation axis is wrong → outlier sleeves. |
| Commit after individual placement | transaction.Commit() | `UpdateZonesFromElements` reads placement point, bbox, params from placed sleeves. `ExtractAndSaveCornersForZones` reads 4 corners for clustering proximity. Both need committed geometry. |
| Commit after cluster placement | transaction.Commit() | `UpdateClusterBoundingBoxesAfterPlacement` reads cluster bbox. `ExtractAndSaveCornersForClusters` reads cluster corners. Both need committed geometry. |
| Commit after deletion | transaction.Commit() | Revit needs to finalize element deletion. |

---

## Implementation Rules

1. **Code reuse only** — no reinventing. Use existing `ExecuteBulkPlacement`, `UpdateZonesFromElements`, `ExtractAndSaveCornersForZones`, `PlaceBulkClusters`, `UpdateClusterBoundingBoxesAfterPlacement` as-is.
2. **No changes to data persistence** — all DB writes, flag management, snapshot transfers stay untouched.
3. **No changes to parameter service** — `SleeveParameterService`, `FlushDeferredParameters`, `QueueRectangularClusterParameters` called the same way.
4. **No changes to rotation logic** — `SleeveRotationService`, `ClusterRotationService`, `ApplyRotation` untouched.
5. **Feature flag controlled** — `UseOptimizedMultiFloorFlow` (default OFF). Existing legacy flow untouched.

## Implementation Steps

### Step 1: Git Checkpoint
```bash
git checkout -b backup/before-multi-floor-optimization
git tag -a v1.0-before-optimization -m "Working state before multi-floor optimization"
```

### Step 2: Optimization 1 — Consolidate Individual Placement Transactions

**File**: `MultiFloorBatchPlacementService.cs`
**Change**: Modify `PlaceChunkInTransaction` loop → wrap ALL chunks in ONE transaction

**Current code (per-chunk transaction):**
```csharp
for (int i = 0; i < chunks.Count; i++)
{
    var chunkResult = PlaceChunkInTransaction(chunks[i], i + 1, chunks.Count);
    totalPlaced += chunkResult.placed;
}
```

**New code (single transaction, multiple batches):**
```csharp
using (var transaction = new Transaction(_doc, "Place Sleeves - All Floors"))
{
    transaction.Start();

    // Batch all NewFamilyInstances2 calls, collect all results
    var allPlacedItems = new List<(ClashZone Zone, ElementId ElementId)>();
    for (int i = 0; i < chunks.Count; i++)
    {
        // Reuse ExecuteBulkPlacement for each chunk (existing code)
        var result = bulkService.ExecuteBulkPlacement(_doc, chunks[i], skipSpatialFiltering: false);
        allPlacedItems.AddRange(result.PlacedItems);
        totalPlaced += result.PlacedCount;
    }
    // ExecuteBulkPlacement already does: NewFamilyInstances2 → Regen → Rotate → Params
    // The Regen inside ExecuteBulkPlacement (line 564) is preserved per batch

    _parameterService.FlushDeferredParameters(clearList: true, context: "AllFloors");
    transaction.Commit();  // ONE commit for all chunks
}

// Post-commit: existing code, no changes
UpdateZonesFromElements(doc, allPlacedItems);
ExtractAndSaveCornersForZones(doc, placedZones);
```

**Note**: Each `ExecuteBulkPlacement` call still does its own `doc.Regenerate()` after `NewFamilyInstances2` internally (line 564). This is preserved — it prevents the outlier bug. The saving is in removing N-1 transaction commits.

### Step 3: Optimization 2 — Move Deletion Out of Cluster Transaction

**File**: `BatchClusterPlacementService.cs` → `PlaceBulkClusters` method
**Change**: Collect sleeves to delete, but don't delete inside the cluster transaction. Return them to caller. Caller deletes in separate transaction after cluster bbox/corner extraction.

**Current code (deletion inside cluster tx):**
```csharp
// Inside PlaceBulkClusters, inside cluster transaction:
if (sleevesToDeleteNow.Any())
{
    doc.Delete(sleevesToDeleteNow);  // line 2089 — adds weight to cluster commit
}
```

**New code (return deletion list, delete separately):**
```csharp
// Inside PlaceBulkClusters: SKIP deletion, return list instead
// persistenceData.SleevesToDelete = sleevesToDeleteNow;  // pass out

// In caller (PlaceFromDatabase), AFTER commit + bbox extraction:
// Transaction 2: Deletion (separate, lighter)
if (sleevesToDelete.Any())
{
    using (var delTx = new Transaction(doc, "Delete Clustered Individuals"))
    {
        delTx.Start();
        doc.Delete(sleevesToDelete);
        delTx.Commit();
    }
}
```

### Step 4: Feature Flag

**File**: `OptimizationFlags.cs` (existing)
```csharp
/// <summary>
/// Consolidates N chunk transactions into 1 for individual placement.
/// Moves deletion out of cluster transaction.
/// Reduces transaction overhead without changing any placement/rotation/parameter logic.
/// </summary>
public static bool UseOptimizedMultiFloorFlow { get; set; } = false; // Default: false (safe)
```

---

## What Is NOT Changed

| Component | Status |
|-----------|--------|
| `BulkPlacementService.ExecuteBulkPlacement` | Untouched — still does NewFamilyInstances2 + Regen + Rotate + Params per batch |
| `SleeveParameterService` | Untouched — same deferred queue + flush |
| `SleeveRotationService` / `ApplyRotation` | Untouched — same rotation logic |
| `UpdateZonesFromElements` | Untouched — same post-commit geometry read |
| `BatchSleeveCornerExtractor` | Untouched — same corner extraction |
| `BatchClusterPlacementService.PlaceBulkClusters` | Minor change only: skip `doc.Delete` call, return list |
| `UpdateClusterBoundingBoxesAfterPlacement` | Untouched — same post-commit bbox read |
| `ClusterRotationService` | Untouched — same pre-calculated rotation |
| All DB persistence / flag management | Untouched |
| Legacy flow (flag = false) | Untouched — zero risk to existing users |

## Rollback Strategy

```bash
# Quick rollback to working state
git checkout v1.0-before-optimization

# Or disable via flag (no code change needed, no rebuild needed if flag is runtime)
OptimizationFlags.UseOptimizedMultiFloorFlow = false;
```

## Testing Strategy
1. **Flag = false** → must behave exactly as current code (regression test)
2. **Flag = true** → optimized flow
3. **Compare**: same sleeves placed, same clusters formed, same DB data

## Expected Performance Gain

| Scenario | Before (commits) | After (commits) | Saved |
|----------|-------------------|------------------|-------|
| 10 chunks individual | 10 commits | 1 commit | 9 commits |
| Clustering + deletion | 1 heavy commit | 1 light + 1 delete commit | Faster cluster commit |
| **Total for 10-floor job** | **11 commits** | **3 commits** | **73% fewer** |

Each saved commit = saved Revit implicit regen + BIM360 sync overhead.

## Files to Modify

1. `MultiFloorBatchPlacementService.cs` — wrap chunk loop in single transaction
2. `BatchClusterPlacementService.cs` — skip deletion in PlaceBulkClusters, return list
3. `BatchClusterPlacementService.cs` — caller deletes in separate transaction after bbox extraction
4. `OptimizationFlags.cs` — flag already exists, just ensure it gates the new paths
