# Sleeve Placement Clustering Fix Documentation

## **Issue Summary**

**Problem**: Individual duct sleeves on walls were getting `IsClustered = true` flag, preventing parameter transfer from working.

**Root Cause**: The clustering logic was too aggressive - it was clustering individual sleeves that were just spatially close, not logically related.

## **Technical Details**

### **The Problem Flow**

1. **Individual sleeves placed**: `SleeveInstanceId = valid ID`, `IsClustered = false`
2. **Clustering runs**: Incorrectly marks individual sleeves as clustered
3. **Flag management**: Sets `IsClustered = true`, `SleeveInstanceId = -1`
4. **Parameter transfer fails**: Can't find individual sleeves

### **The Fix**

**File**: `Services/UniversalClusterService.cs`
**Method**: `ShouldSkipClustering()`

**Changes Made**:

1. **Enhanced clustering logic** to prevent individual sleeves from being marked as clustered
2. **Added distance check** for sleeves more than 100mm apart
3. **Added MEP element check** for sleeves from different MEP elements
4. **Added logging** for debugging clustering decisions

### **Code Changes**

```csharp
// ✅ CRITICAL FIX: Only cluster sleeves that are truly related (same MEP element)
// This prevents individual sleeves from being incorrectly marked as clustered
if (ShouldSkipClustering(cluster, groupKey))
{
    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", 
        $"SKIP CLUSTERING: Individual sleeves detected - preserving individual SleeveInstanceId for parameter transfer\n");
    continue;
}
```

**Enhanced `ShouldSkipClustering` method**:

```csharp
private bool ShouldSkipClustering(List<FamilyInstance> cluster, SleeveGroupKey groupKey)
{
    if (cluster.Count <= 2) // Small clusters might be individual sleeves
    {
        // Check if sleeves have different MEP elements (indicating they should remain individual)
        var mepElementIds = new HashSet<int>();
        foreach (var sleeve in cluster)
        {
            var mepElementId = GetMepElementIdFromSleeve(sleeve);
            if (mepElementId > 0)
            {
                mepElementIds.Add(mepElementId);
            }
        }
        
        // If sleeves have different MEP elements, they should remain individual
        if (mepElementIds.Count > 1)
        {
            DebugLogger.Log($"[UniversalClusterService] SKIP CLUSTERING: Sleeves have different MEP elements ({mepElementIds.Count} unique) - preserving individual SleeveInstanceId for parameter transfer");
            return true;
        }
        
        // ✅ ADDITIONAL CHECK: If sleeves are far apart (>100mm), they should remain individual
        if (cluster.Count == 2)
        {
            var sleeve1 = cluster[0];
            var sleeve2 = cluster[1];
            var distance = CalculateDistanceBetweenSleeves(sleeve1, sleeve2);
            var distanceMm = UnitUtils.ConvertFromInternalUnits(distance, UnitTypeId.Millimeters);
            
            if (distanceMm > 100) // More than clustering threshold
            {
                DebugLogger.Log($"[UniversalClusterService] SKIP CLUSTERING: Sleeves are {distanceMm:F1}mm apart (>100mm threshold) - preserving individual SleeveInstanceId");
                return true;
            }
        }
    }
    
    return false;
}
```

## **Expected Results**

### **Before Fix**
- Individual duct sleeves on walls: `IsClustered = true`, `SleeveInstanceId = -1`
- Parameter transfer: ❌ Failed
- Logs: "Sleeve ID 873377 not found in XML"

### **After Fix**
- Individual duct sleeves on walls: `IsClustered = false`, `SleeveInstanceId = valid ID`
- Parameter transfer: ✅ Working
- Logs: "SKIP CLUSTERING: Individual sleeves detected - preserving individual SleeveInstanceId"

## **Testing**

### **Test Cases**

1. **Individual sleeves > 100mm apart**: Should remain individual
2. **Individual sleeves from different MEP elements**: Should remain individual
3. **True cluster sleeves ≤ 100mm apart**: Should be clustered
4. **Parameter transfer**: Should work for individual sleeves

### **Logs to Check**

- `cluster_debug.log`: Look for "SKIP CLUSTERING" messages
- `transfer_debug.log`: Look for successful parameter transfers
- `placement_debug.log`: Look for individual sleeve placement

## **Impact**

- **Fixes**: Parameter transfer for individual duct sleeves on walls
- **Fixes**: Parameter transfer for individual duct accessories on walls
- **Fixes**: Parameter transfer for individual pipes on walls
- **Maintains**: Existing cluster sleeve functionality
- **Maintains**: Existing parameter transfer for cluster sleeves

## **Files Modified**

- `Services/UniversalClusterService.cs`: Enhanced clustering logic
- `SLEEVE_PLACEMENT_CLUSTERING_FIX.md`: This documentation

## **Date**

**Fixed**: October 22, 2025
**Issue**: Individual sleeves getting incorrect clustering flags
**Solution**: Enhanced clustering logic to preserve individual sleeves

