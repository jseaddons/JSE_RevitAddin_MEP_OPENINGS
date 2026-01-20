# Cluster Corner Persistence Fix

## Problem
Cluster sleeve corners were being calculated correctly but showing as 0.0 in the database.

## Root Cause
We were trying to use `INSERT OR REPLACE` in `ClusterSleeveRepository.BatchSaveClusterSleevesBulk`, but this doesn't match the pattern used for individual sleeves.

## Solution
Use the same pattern as individual sleeves:

### Individual Sleeves (Working)
`SleevePersistenceService.SaveSleeveCorners()` calls:
```csharp
repository.UpdateSleeveCorners(zone.Id, c1x, c1y, c1z, ...)
```

### Cluster Sleeves (Should Match)
In `RefactoredClusterService.BatchSaveClusterDataToDatabase()`, after calculating corners, call:
```csharp
_clashZoneRepository.UpdateClusterSleeveCorners(
    clusterInstanceId,
    c1x, c1y, c1z,
    c2x, c2y, c2z,
    c3x, c3y, c3z,
    c4x, c4y, c4z);
```

## Implementation
Add this call in `RefactoredClusterService.cs` at line ~2467, right after the corner debug logging.

The method `UpdateClusterSleeveCorners` already exists in `ClashZoneRepository.cs` (line 5722) and does a simple UPDATE on the `ClusterSleeves` table based on `ClusterInstanceId`.

## Files to Modify
1. **RefactoredClusterService.cs** (line ~2467): Add `UpdateClusterSleeveCorners` call after corner calculation
2. **ClusterSleeveRepository.cs**: Fixed compilation errors (added `using System.IO;`)
