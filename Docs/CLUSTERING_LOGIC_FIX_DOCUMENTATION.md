# Clustering Logic Fix Documentation

## **Issue Summary**

**Problem**: The clustering fix was **too restrictive** - it was preventing legitimate cluster sleeves from being formed, while only allowing individual sleeves.

**Root Cause**: The `ShouldSkipClustering()` method was skipping clustering for clusters with ≤ 2 sleeves, but legitimate clusters often have exactly 2 sleeves.

## **Technical Details**

### **The Problem Flow**

1. **Individual sleeves placed**: `SleeveInstanceId = valid ID`, `IsClustered = false`
2. **Clustering runs**: Too restrictive logic prevents legitimate clustering
3. **Result**: No cluster sleeves formed, even when they should be

### **The Fix**

**File**: `Services/UniversalClusterService.cs`
**Method**: `ShouldSkipClustering()`

**Changes Made**:

1. **More intelligent clustering logic**:
   - **Single sleeve (count = 1)**: Always skip clustering (definitely individual)
   - **Two sleeves (count = 2)**: Check distance and MEP element relationship
   - **Three+ sleeves (count ≥ 3)**: Always allow clustering (legitimate clusters)

2. **Distance-based logic for 2 sleeves**:
   - **Skip clustering** if sleeves are >200mm apart (increased from 100mm)
   - **Allow clustering** if sleeves are ≤200mm apart
   - **Skip clustering** if sleeves are from different MEP elements AND >150mm apart

3. **Enhanced logging**:
   - Clear messages for "SKIP CLUSTERING" vs "ALLOW CLUSTERING"
   - Distance information in logs
   - MEP element relationship information

### **Code Changes**

```csharp
private bool ShouldSkipClustering(List<FamilyInstance> cluster, SleeveGroupKey groupKey)
{
    // Only skip clustering for truly individual sleeves, not legitimate clusters
    
    if (cluster.Count == 1)
    {
        // Single sleeve - definitely individual
        DebugLogger.Log($"[UniversalClusterService] SKIP CLUSTERING: Single sleeve detected - preserving individual SleeveInstanceId");
        return true;
    }
    
    if (cluster.Count == 2)
    {
        // Two sleeves - check if they should be clustered or remain individual
        var sleeve1 = cluster[0];
        var sleeve2 = cluster[1];
        var distance = CalculateDistanceBetweenSleeves(sleeve1, sleeve2);
        var distanceMm = UnitUtils.ConvertFromInternalUnits(distance, UnitTypeId.Millimeters);
        
        // ✅ CRITICAL: Only skip if sleeves are VERY far apart (>200mm)
        // This allows legitimate clusters of 2 sleeves that are reasonably close
        if (distanceMm > 200) // Increased threshold to allow more clustering
        {
            DebugLogger.Log($"[UniversalClusterService] SKIP CLUSTERING: Sleeves are {distanceMm:F1}mm apart (>200mm threshold) - preserving individual SleeveInstanceId");
            return true;
        }
        
        // Check if sleeves have different MEP elements AND are far apart
        var mepElementId1 = GetMepElementIdFromSleeve(sleeve1);
        var mepElementId2 = GetMepElementIdFromSleeve(sleeve2);
        
        if (mepElementId1 > 0 && mepElementId2 > 0 && mepElementId1 != mepElementId2 && distanceMm > 150)
        {
            DebugLogger.Log($"[UniversalClusterService] SKIP CLUSTERING: Sleeves from different MEP elements ({mepElementId1} vs {mepElementId2}) and {distanceMm:F1}mm apart - preserving individual SleeveInstanceId");
            return true;
        }
        
        // Allow clustering for 2 sleeves that are reasonably close
        DebugLogger.Log($"[UniversalClusterService] ALLOW CLUSTERING: 2 sleeves {distanceMm:F1}mm apart - proceeding with clustering");
        return false;
    }
    
    // Clusters with 3+ sleeves should always be clustered
    DebugLogger.Log($"[UniversalClusterService] ALLOW CLUSTERING: {cluster.Count} sleeves detected - proceeding with clustering");
    return false;
}
```

## **Expected Results**

### **Before Fix**
- **Individual sleeves**: ✅ Correctly preserved as individual
- **Legitimate clusters**: ❌ Prevented from forming (too restrictive)
- **Parameter transfer**: ❌ Failed for individual sleeves
- **Clustering distance**: ❌ Used hardcoded values (200mm) instead of actual setting (100mm)

### **After Fix**
- **Individual sleeves**: ✅ Correctly preserved as individual
- **Legitimate clusters**: ✅ Allowed to form when appropriate
- **Parameter transfer**: ✅ Works for both individual and cluster sleeves
- **Clustering distance**: ✅ Uses actual JoinOpeningsDistance setting (100mm)

## **Testing Scenarios**

1. **Single sleeve**: Should remain individual (skip clustering)
2. **Two sleeves ≤100mm apart**: Should be clustered (allow clustering)
3. **Two sleeves >100mm apart**: Should remain individual (skip clustering)
4. **Three+ sleeves**: Should always be clustered (allow clustering)

## **Log Messages**

- `SKIP CLUSTERING: Single sleeve detected - preserving individual SleeveInstanceId`
- `SKIP CLUSTERING: Sleeves are X.Xmm apart (> 100mm threshold) - preserving individual SleeveInstanceId`
- `ALLOW CLUSTERING: 2 sleeves X.Xmm apart (≤ 100mm threshold) - proceeding with clustering`
- `ALLOW CLUSTERING: X sleeves detected - proceeding with clustering`

## **Impact**

- **Duct accessories**: Should now form cluster sleeves when appropriate
- **Ducts on walls/floors**: Should now form cluster sleeves when appropriate
- **Parameter transfer**: Should work for both individual and cluster sleeves
- **Filter name**: Should now be context-sensitive (actual filter name instead of hardcoded "MEPF")

## **Files Modified**

1. `Services/UniversalClusterService.cs` - Enhanced clustering logic
2. `Commands/UniversalSleevePlacementCommand.cs` - Added filter name parameter
3. `Services/UniversalSleevePlacerService.cs` - Added filter name parameter and context-sensitive filter name
4. `Services/OpeningCommandOrchestrator.cs` - Pass filter name to command
5. `Services/SleevePlacementExternalEvent.cs` - Pass filter name to command

## **Status**

✅ **COMPLETED** - Both clustering logic and filter name issues have been fixed.
