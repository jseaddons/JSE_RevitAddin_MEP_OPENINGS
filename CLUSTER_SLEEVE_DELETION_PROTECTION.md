# Cluster Sleeve Deletion Protection & AfterSleeve ID

## Summary

This document explains:
1. **AfterSleeve ID** - How deleted individual sleeves are tracked
2. **When individual sleeves are deleted** - Only when cluster is successfully formed
3. **Protection mechanisms** - Multiple safeguards to prevent incorrect deletions

---

## 1. AfterSleeve ID (`AfterClusterSleevePlacedSleeveInstanceId`)

### What is it?

**`AfterClusterSleevePlacedSleeveInstanceId`** is a database field in the `ClashZones` table that stores the **original individual sleeve ID** before it was deleted by cluster formation.

### Purpose

- **Track deleted sleeves**: Preserves the ID of individual sleeves that were replaced by a cluster sleeve
- **Database consistency**: Allows tracking which individual sleeves were part of a cluster
- **Recovery**: Can be used to identify which sleeves were deleted if needed for debugging

### When is it set?

**Location**: `Services/Clustering/RefactoredClusterService.cs` (line 1323)

```csharp
// ✅ CRITICAL: Save original SleeveInstanceId to AfterClusterSleevePlacedSleeveInstanceId 
// BEFORE UpdateFlagsForPlacement clears it
if (originalSleeveInstanceId > 0)
{
    clashZone.AfterClusterSleevePlacedSleeveInstanceId = originalSleeveInstanceId;
}
```

**Timing**: Set **BEFORE** individual sleeves are deleted, **AFTER** cluster sleeve is successfully placed.

### Database Schema

**Model**: `Models/ClashZone.cs` (line 284)
```csharp
public int AfterClusterSleevePlacedSleeveInstanceId { get; set; } = -1;
```

**Default Value**: `-1` (indicates no individual sleeve was deleted)

---

## 2. When Are Individual Sleeves Deleted?

### ✅ CORRECT BEHAVIOR: Individual sleeves are deleted ONLY when:

1. **Cluster sleeve is successfully placed** (`placementSuccess == true`)
2. **Cluster sleeve exists in document** (`placedClusterSleeve != null && IsValidObject`)
3. **Cluster sleeve ID is captured** (`capturedClusterSleeveId.HasValue`)

### ❌ PROTECTION: Individual sleeves are NOT deleted when:

1. **Cluster placement fails** (`placementSuccess == false`)
   - **Location**: `Services/Clustering/RefactoredClusterService.cs` (line 1266)
   - **Action**: Returns `(false, 0, 0, null, null)` - **NO deletion occurs**

2. **Cluster sleeve is null** (`placedClusterSleeve == null`)
   - **Location**: `Services/Clustering/RefactoredClusterService.cs` (line 1371)
   - **Action**: Returns `(false, 0, 0, null, null)` - **NO deletion occurs**

3. **Cluster sleeve not found in document** (`verifyBeforeDelete == null`)
   - **Location**: `Services/Clustering/RefactoredClusterService.cs` (line 1391)
   - **Action**: Returns `(false, 0, 0, null, null)` - **NO deletion occurs**

4. **No clusters placed** (`placedClusters.Count == 0`)
   - **Location**: `Services/Clustering/RefactoredClusterService.cs` (line 638)
   - **Action**: Cleanup service is **SKIPPED** - **NO deletion occurs**

---

## 3. Protection Mechanisms

### Protection Layer 1: Early Return on Placement Failure

**Location**: `Services/Clustering/RefactoredClusterService.cs` (line 1253-1266)

```csharp
if (!placementSuccess || placedClusterSleeve == null)
{
    return (false, 0, 0, null, null); // NO deletion - early return
}
```

**Protection**: If cluster placement fails, method returns immediately without deleting any sleeves.

### Protection Layer 2: Cluster Sleeve Validation

**Location**: `Services/Clustering/RefactoredClusterService.cs` (line 1371-1397)

```csharp
// ✅ CRITICAL PROTECTION: Only delete if BOTH conditions are met
if (!capturedClusterSleeveId.HasValue || placedClusterSleeve == null)
{
    return (false, 0, 0, null, null); // NO deletion
}

// Verify cluster sleeve exists in document
var verifyBeforeDelete = doc.GetElement(new ElementId(capturedClusterSleeveId.Value));
if (verifyBeforeDelete == null || !verifyBeforeDelete.IsValidObject)
{
    return (false, 0, 0, null, null); // NO deletion - cluster sleeve lost
}
```

**Protection**: Triple-checks that cluster sleeve exists before proceeding with deletion.

### Protection Layer 3: Individual Sleeve Validation

**Location**: `Services/Clustering/RefactoredClusterService.cs` (line 1412-1426)

```csharp
// ✅ PROTECTION: Don't delete the cluster sleeve itself
if (placedClusterSleeve != null && sleeveElementId == placedClusterSleeve.Id)
{
    continue; // Skip cluster sleeve
}

// ✅ DOUBLE-CHECK: Verify this is NOT the cluster sleeve by ID
if (capturedClusterSleeveId.HasValue && sleeveInstanceId == capturedClusterSleeveId.Value)
{
    continue; // Skip cluster sleeve
}
```

**Protection**: Prevents cluster sleeve from being added to deletion list.

### Protection Layer 4: Final Deletion Check

**Location**: `Services/Clustering/RefactoredClusterService.cs` (line 1451-1457)

```csharp
// ✅ TRIPLE-CHECK: Verify this is NOT the cluster sleeve
if (capturedClusterSleeveId.HasValue && id.IntegerValue == capturedClusterSleeveId.Value)
{
    continue; // Skip cluster sleeve
}
```

**Protection**: Final check before actual deletion to prevent cluster sleeve deletion.

### Protection Layer 5: Cleanup Service Protection

**Location**: `Services/Clustering/Cleanup/ClusterCleanupService.cs` (line 59-75)

```csharp
// Build protection set of cluster sleeve IDs
var clusterSleeveIds = new HashSet<int>(placedClusters.Where(c => c != null).Select(c => c.Id.IntegerValue));

// Skip cluster sleeves
if (clusterSleeveIds.Contains(id))
{
    continue; // protected
}

// Check parameters
bool isClusterSleeve = (sleeveInstanceValue == -1) || (clusterValue > 0 && clusterValue == id);
if (isClusterSleeve)
{
    continue; // Skip cluster sleeve
}
```

**Protection**: Cleanup service uses protection set AND parameter checks to identify cluster sleeves.

---

## 4. Code Flow for Individual Sleeve Deletion

```
1. ClusterSleeves() called
   ↓
2. FormClusters() - Groups sleeves into clusters
   ↓
3. PlaceClusterForGroup() - For each cluster:
   ├─ Calculate bounding box
   ├─ PlaceClusterSleeve() - Place cluster sleeve
   │  ├─ ✅ SUCCESS → Continue
   │  └─ ❌ FAILURE → Return (false, 0, 0, null, null) [NO DELETION]
   ↓
4. Set AfterClusterSleevePlacedSleeveInstanceId (preserve original sleeve ID)
   ↓
5. Verify cluster sleeve exists in document
   ├─ ✅ EXISTS → Continue
   └─ ❌ NOT FOUND → Return (false, 0, 0, null, null) [NO DELETION]
   ↓
6. Queue individual sleeves for deletion (with triple-checks)
   ↓
7. Delete individual sleeves in batch
   ↓
8. Verify cluster sleeve still exists after deletion
   ↓
9. CleanupSleevesWithinClusters() - Additional cleanup (only if clusters placed)
   ├─ ✅ Has valid clusters → Run cleanup
   └─ ❌ No valid clusters → SKIP cleanup [NO DELETION]
```

---

## 5. If Cluster Path is Not Entered

### Question: "If the cluster path is not entered, how do individual sleeves get deleted?"

**Answer**: Individual sleeves are **NOT deleted** if the cluster path is not entered.

### Scenarios:

#### Scenario 1: No Clusters Formed (0 clusters)
- **Result**: `placedClusters.Count == 0`
- **Action**: Cleanup service is **SKIPPED** (line 648)
- **Individual Sleeves**: **NOT DELETED** ✅

#### Scenario 2: Cluster Placement Failed
- **Result**: `placementSuccess == false`
- **Action**: Early return `(false, 0, 0, null, null)` (line 1266)
- **Individual Sleeves**: **NOT DELETED** ✅

#### Scenario 3: Cluster Sleeve Lost After Placement
- **Result**: `verifyBeforeDelete == null`
- **Action**: Return `(false, 0, 0, null, null)` (line 1391)
- **Individual Sleeves**: **NOT DELETED** ✅

#### Scenario 4: Cluster Successfully Placed
- **Result**: All validations pass
- **Action**: Individual sleeves are deleted (line 1459)
- **Individual Sleeves**: **DELETED** ✅ (expected behavior)

---

## 6. Recent Fixes (2025-12-04)

### Fix 1: Cluster Sleeve Identification Parameters

**Problem**: When `UseBatchedParameterWrites` is enabled, cluster sleeve identification parameters (`Sleeve Instance ID = -1`, `Cluster Sleeve Instance ID`) were deferred, causing cleanup service to misidentify cluster sleeves.

**Fix**: Changed to **always set these parameters immediately** (not deferred), regardless of batching settings.

**Location**: `Services/Clustering/Placement/ClusterPlacementService.cs` (line 790-820)

### Fix 2: Enhanced Deletion Protection

**Problem**: If cluster sleeve was not found before deletion, method returned `(true, 1, 0, ...)` which could cause confusion.

**Fix**: Changed to return `(false, 0, 0, null, null)` when cluster sleeve validation fails, ensuring individual sleeves are never deleted if cluster is not properly formed.

**Location**: `Services/Clustering/RefactoredClusterService.cs` (line 1371-1397)

### Fix 3: Added Diagnostic Logging

**Enhancement**: Added comprehensive logging to track:
- When parameters are set (immediate vs deferred)
- When batch flush runs and how many parameters are flushed
- When individual sleeves are queued for deletion
- When cluster sleeve validation passes/fails

**Location**: Multiple locations in `RefactoredClusterService.cs` and `ClusterPlacementService.cs`

---

## 7. Verification Checklist

To verify individual sleeves are NOT deleted when cluster is not formed:

- [ ] Check logs for `"PLACEMENT FAILED"` - should show `(false, 0, 0, null, null)`
- [ ] Check logs for `"CRITICAL PROTECTION"` - should show deletion skipped
- [ ] Check logs for `"CLEANUP: No valid cluster sleeves"` - cleanup should be skipped
- [ ] Verify `AfterClusterSleevePlacedSleeveInstanceId` is set BEFORE deletion
- [ ] Verify individual sleeves still exist in Revit after failed cluster placement

---

## 8. Database Fields Reference

| Field | Purpose | Default | Set When |
|-------|---------|---------|----------|
| `SleeveInstanceId` | Current individual sleeve ID | -1 | When individual sleeve is placed |
| `ClusterSleeveInstanceId` | Cluster sleeve ID that replaced individual sleeve | -1 | When cluster is formed |
| `AfterClusterSleevePlacedSleeveInstanceId` | Original individual sleeve ID (before deletion) | -1 | BEFORE individual sleeve is deleted |
| `IsClusterResolved` | Flag indicating cluster was formed | false | When cluster is successfully placed |

---

## Conclusion

**Individual sleeves are ONLY deleted when:**
1. ✅ Cluster sleeve is successfully placed
2. ✅ Cluster sleeve exists in document
3. ✅ All validations pass

**Individual sleeves are NEVER deleted when:**
1. ❌ Cluster placement fails
2. ❌ Cluster sleeve is null
3. ❌ Cluster sleeve not found in document
4. ❌ No clusters are placed
5. ❌ Cluster path is not entered

**AfterSleeve ID (`AfterClusterSleevePlacedSleeveInstanceId`) is set BEFORE deletion** to preserve the original individual sleeve ID for tracking purposes.

