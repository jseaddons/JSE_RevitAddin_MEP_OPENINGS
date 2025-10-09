# Sleeve Flag Management Rules - Complete Reference

## Overview
The sleeve placement system uses **two critical flags** in the `ClashZone` XML to manage individual and cluster sleeve placement:
1. **`IsResolved`** - Indicates an individual sleeve has been placed
2. **`IsClusterResolved`** - Indicates a cluster sleeve has been placed

These flags control **duplication suppression** and **MEPMARK application**.

---

## Flag Properties in ClashZone Model

### Individual Sleeve Tracking
```csharp
public bool IsResolved { get; set; } = false;           // Individual sleeve placed
public int SleeveInstanceId { get; set; } = -1;         // Individual sleeve ElementId (serialized to XML)
public string SleeveFamilyName { get; set; } = "";     // Family used for individual sleeve
```

### Cluster Sleeve Tracking
```csharp
public bool IsClusterResolved { get; set; } = false;                    // Cluster sleeve placed
[XmlIgnore] public ElementId? ClusterSleeveId { get; set; }             // In-memory only (NOT serialized)
public int ClusterSleeveInstanceId { get; set; } = -1;                  // Cluster sleeve ElementId (serialized to XML) ✅
```

**⚠️ CRITICAL**: `ClusterSleeveId` has `[XmlIgnore]` and is NOT saved to XML. Always use `ClusterSleeveInstanceId` for persistence!

---

## Flag Management Rules

### 1. Individual Sleeve Placement (UniversalSleevePlacerService)

#### Placement Logic
```csharp
if (clashZone.IsResolved && clashZone.SleeveInstanceId > 0)
{
    // Check if sleeve still exists in Revit
    var existingSleeve = doc.GetElement(new ElementId(clashZone.SleeveInstanceId));
    if (existingSleeve != null)
    {
        SkippedCount++;
        continue; // ✅ Skip - individual sleeve already exists
    }
    else
    {
        // Sleeve was deleted - reset flag and place new one
        clashZone.IsResolved = false;
        clashZone.SleeveInstanceId = -1;
    }
}

// Place individual sleeve
var sleeveInstance = PlaceSleeve(...);
clashZone.IsResolved = true;  // ✅ Mark as resolved
clashZone.SleeveInstanceId = sleeveInstance.Id.IntegerValue;  // ✅ Store ID
clashZone.SleeveFamilyName = familySymbol.Family.Name;
PlacedCount++;
```

**Key Points**:
- ✅ Sets `IsResolved = true` AFTER successful placement
- ✅ Stores actual sleeve `ElementId` as integer in `SleeveInstanceId`
- ✅ Checks if existing sleeve was deleted → resets flag if deleted
- ❌ Does NOT save XML (flags updated in-memory only during placement transaction)

---

### 2. Cluster Sleeve Placement (UniversalClusterService)

#### Pre-Placement Check (Duplication Suppression)
```csharp
// Called at start of ClusterSleeves() - line 68
ResetClusterFlagsForDeletedSleeves(doc);

// For each proposed cluster - line 192
if (IsClusterAlreadyExists(cluster, groupKey))
{
    continue; // ✅ Skip - cluster already exists
}
```

#### Reset Logic (ResetClusterFlagsForDeletedSleeves)
```csharp
foreach (var clashZone in filter.ClashZoneStorage.ClashZones)
{
    if (clashZone.IsClusterResolved)
    {
        // ✅ Check integer version (serialized to XML)
        if (clashZone.ClusterSleeveInstanceId <= 0)
        {
            // Cluster sleeve ID invalid - reset flag
            clashZone.IsClusterResolved = false;
            clashZone.ClusterSleeveInstanceId = -1;
            // ✅ Save XML immediately
        }
        else
        {
            // Check if cluster sleeve still exists in Revit
            var clusterSleeveId = new ElementId(clashZone.ClusterSleeveInstanceId);
            var clusterSleeve = doc.GetElement(clusterSleeveId);
            if (clusterSleeve == null)
            {
                // Cluster sleeve was deleted - reset flag
                clashZone.IsClusterResolved = false;
                clashZone.ClusterSleeveInstanceId = -1;
                // ✅ Save XML immediately
            }
        }
    }
}
```

#### Cluster Existence Check (IsClusterAlreadyExists)
```csharp
foreach (var sleeve in cluster)
{
    var mepElementId = sleeve.LookupParameter("MEP_ElementId")?.AsInteger();
    var clashZone = GetClashZoneByMepElementId(mepElementId);
    
    // ✅ Check both flag AND valid ID
    if (clashZone != null && 
        clashZone.IsClusterResolved && 
        clashZone.ClusterSleeveInstanceId > 0)
    {
        return true; // Cluster already exists
    }
}
return false; // No cluster exists - safe to create
```

#### Placement & Flag Setting
```csharp
// After successfully placing cluster sleeve
var clusterSleeve = PlaceClusterSleeve(...);

// ✅ Set MEP_ElementId on cluster sleeve
clusterSleeve.LookupParameter("MEP_ElementId")?.Set(firstSleeve.LookupParameter("MEP_ElementId").AsInteger());

// ✅ Mark clash zones as cluster-resolved
MarkClashZonesAsClusterResolvedWithSleeveId(cluster, clusterSleeve.Id);

// Delete original individual sleeves
foreach (var s in cluster)
{
    doc.Delete(s.Id);
}
```

#### Flag Marking Logic (MarkClashZonesAsClusterResolvedWithSleeveId)
```csharp
foreach (var sleeve in cluster)
{
    var mepElementId = sleeve.LookupParameter("MEP_ElementId")?.AsInteger();
    var clashZone = FindClashZoneByMepElementId(mepElementId);
    
    if (clashZone != null)
    {
        clashZone.IsClusterResolved = true;  // ✅ Mark as cluster-resolved
        clashZone.ClusterSleeveId = clusterSleeveId;  // In-memory only (NOT saved)
        clashZone.ClusterSleeveInstanceId = clusterSleeveId.IntegerValue;  // ✅ Saved to XML
        clashZone.LastUpdated = DateTime.Now;
    }
}
// ✅ Save XML immediately after marking
```

---

## Flag Combinations & Behavior

| IsResolved | IsClusterResolved | SleeveInstanceId | ClusterSleeveInstanceId | Behavior |
|-----------|------------------|-----------------|------------------------|----------|
| `false` | `false` | `-1` | `-1` | **Fresh clash zone** - No sleeve placed yet → Place individual sleeve |
| `true` | `false` | `> 0` | `-1` | **Individual sleeve exists** - Skip placement (already resolved) |
| `false` | `true` | `-1` | `> 0` | **Cluster sleeve exists** - Skip individual placement, skip clustering |
| `true` | `true` | `> 0` | `> 0` | **Invalid state** - Should never happen (individual deleted when cluster created) |
| `true` | `false` | `-1` | `-1` | **Deleted sleeve** - Reset during next placement, will re-place |
| `false` | `true` | `-1` | `-1` | **Deleted cluster** - Reset during clustering, will re-cluster |

---

## MEPMARK Application Logic

### Cluster Sleeves (GetClusterSleevesForCategory)
```csharp
foreach (var sleeve in allSleeves)
{
    var mepElementId = sleeve.LookupParameter("MEP_ElementId")?.AsInteger();
    var clashZone = GetClashZoneByMepElementId(mepElementId);
    
    // ✅ ALL four conditions must be true
    if (clashZone != null &&                                      // 1. ClashZone exists
        clashZone.IsClusterResolved &&                            // 2. Is cluster-resolved
        clashZone.ClusterSleeveInstanceId > 0 &&                  // 3. Has valid cluster sleeve ID
        sleeve.Id.IntegerValue == clashZone.ClusterSleeveInstanceId && // 4. This IS the cluster sleeve
        clashZone.MepElementCategory == category)                 // 5. Category matches
    {
        categorySleeves.Add(sleeve); // ✅ Apply MEPMARK
    }
}
```

**Rejection Reasons** (logged for debugging):
- `REJECT {id}: no clash-zone for MEPid {X}` → XML missing or MEP_ElementId mismatch
- `REJECT {id}: IsClusterResolved=false` → Individual sleeve, not cluster
- `REJECT {id}: ClusterSleeveInstanceId={X} != {Y}` → This sleeve is not the cluster (it's one of the originals)
- `REJECT {id}: category='{X}' != '{Y}'` → Wrong category

### Individual Sleeves (GetIndividualSleevesForCategory)
```csharp
foreach (var sleeve in allSleeves)
{
    var mepElementId = sleeve.LookupParameter("MEP_ElementId")?.AsInteger();
    var clashZone = GetClashZoneByMepElementId(mepElementId);
    
    // ✅ Must be resolved but NOT cluster-resolved
    if (clashZone != null && 
        clashZone.IsResolved &&           // Individual sleeve placed
        !clashZone.IsClusterResolved &&   // NOT part of cluster
        clashZone.MepElementCategory == category)
    {
        individualSleeves.Add(sleeve); // ✅ Apply MEPMARK
    }
}
```

---

## XML Persistence Rules

### ✅ Flags That ARE Saved to XML
- `IsResolved` (bool)
- `IsClusterResolved` (bool)
- `SleeveInstanceId` (int) - Individual sleeve ElementId
- `ClusterSleeveInstanceId` (int) - Cluster sleeve ElementId ✅
- `SleeveFamilyName` (string)
- `LastUpdated` (DateTime)

### ❌ Flags That are NOT Saved to XML (XmlIgnore)
- `ResolvedSleeveId` (ElementId?) - In-memory only
- `ClusterSleeveId` (ElementId?) - In-memory only

### When XML is Saved
1. **After individual sleeve placement**: ❌ NOT saved (in-memory only during transaction)
2. **After cluster flag reset**: ✅ Saved immediately (ResetClusterFlagsForDeletedSleeves)
3. **After cluster sleeve placement**: ✅ Saved immediately (MarkClashZonesAsClusterResolvedWithSleeveId)
4. **During Refresh**: ✅ Saved by RefreshService/ClashZoneService

---

## Common Scenarios

### Scenario 1: First Run (Fresh Project)
1. **Refresh** → Clash zones created → `IsResolved = false`, `IsClusterResolved = false`
2. **Place Individual Sleeves** → `IsResolved = true`, `SleeveInstanceId = 12345`
3. **Cluster Sleeves** → Some sleeves clustered → `IsClusterResolved = true`, `ClusterSleeveInstanceId = 67890`, individual sleeves deleted
4. **Apply MEPMARK** → Cluster sleeves get marks (e.g., SLEEVE_DCT001), individual sleeves get marks (e.g., SLEEVE_DCT002)

### Scenario 2: Re-Run (No Changes)
1. **Place Individual Sleeves** → All skipped (`IsResolved = true`, sleeve exists)
2. **Cluster Sleeves** → All skipped (`IsClusterResolved = true`, `ClusterSleeveInstanceId > 0`, cluster exists)
3. **Apply MEPMARK** → Re-applies marks (idempotent, continuous numbering)

### Scenario 3: User Deletes Cluster Sleeve
1. **Cluster Sleeves** → `ResetClusterFlagsForDeletedSleeves()` detects missing cluster → `IsClusterResolved = false`, `ClusterSleeveInstanceId = -1`
2. **Individual Sleeves** → Now visible again (`IsResolved = true` but `IsClusterResolved = false`)
3. **Cluster Sleeves** → Re-clusters individual sleeves → `IsClusterResolved = true` again
4. **Apply MEPMARK** → Marks new cluster sleeve

### Scenario 4: User Deletes Individual Sleeve
1. **Place Individual Sleeves** → Detects missing sleeve → `IsResolved = false`, `SleeveInstanceId = -1` → Places new sleeve → `IsResolved = true`
2. **Cluster Sleeves** → Clusters the new individual sleeve
3. **Apply MEPMARK** → Marks cluster sleeve

### Scenario 5: Mixed Individual + Cluster Sleeves
1. Some clash zones have **individual sleeves** (`IsResolved = true`, `IsClusterResolved = false`)
2. Some clash zones have **cluster sleeves** (`IsResolved = false`, `IsClusterResolved = true`, `ClusterSleeveInstanceId > 0`)
3. **Apply MEPMARK** → Both types get marks:
   - `GetClusterSleevesForCategory()` finds cluster sleeves
   - `GetIndividualSleevesForCategory()` finds individual sleeves
   - Both added to `allSleeves` list → MEPMARK applied to all

---

## Debugging Checklist

### Individual Sleeves Not Placing
- [ ] Check `IsResolved` flag in XML (should be `false` for unplaced sleeves)
- [ ] Check if `SleeveInstanceId > 0` (if yes, check if sleeve exists in Revit)
- [ ] Verify placement code sets `IsResolved = true` after placement
- [ ] Check if XML is being saved after placement (currently NOT saved - known limitation)

### Cluster Sleeves Not Placing
- [ ] Check `IsClusterResolved` flag in XML (should be `false` for unclustered areas)
- [ ] Check `ClusterSleeveInstanceId` in XML (should be `-1` or missing for unclustered)
- [ ] Verify `ResetClusterFlagsForDeletedSleeves()` is called at start of clustering
- [ ] Check if cluster sleeve exists in Revit (if yes, duplication suppression is working)
- [ ] Verify `MarkClashZonesAsClusterResolvedWithSleeveId()` sets `ClusterSleeveInstanceId` (NOT just `ClusterSleeveId`)

### MEPMARK Not Applied to Cluster Sleeves
- [ ] Check `MEP_ElementId` parameter on cluster sleeve (should be set during clustering)
- [ ] Check `IsClusterResolved = true` in XML
- [ ] Check `ClusterSleeveInstanceId > 0` in XML
- [ ] Check if cluster sleeve ElementId matches `ClusterSleeveInstanceId` in XML
- [ ] Check category in XML matches requested category
- [ ] Review `mepmark_debug.log` for REJECT messages

### MEPMARK Not Applied to Individual Sleeves
- [ ] Check `IsResolved = true` in XML
- [ ] Check `IsClusterResolved = false` in XML (should NOT be clustered)
- [ ] Check `MEP_ElementId` parameter on individual sleeve
- [ ] Check category in XML matches requested category
- [ ] Review `mepmark_debug.log` for individual sleeve findings

---

## Critical Bug Fixes History

### Bug: ClusterSleeveId Not Persisting to XML
**Cause**: `ClusterSleeveId` had `[XmlIgnore]` attribute → never saved to XML  
**Symptom**: Cluster flags reset on every run → clustering never worked → MEPMARK never applied  
**Fix**: Added `ClusterSleeveInstanceId` integer property (WITHOUT `[XmlIgnore]`) → now saved to XML  
**Date**: 2025-10-09  

### Bug: Individual Sleeve Flags Not Persisting
**Cause**: `UniversalSleevePlacerService` sets `IsResolved = true` in-memory but never saves XML  
**Symptom**: Individual sleeves not found for MEPMARK on subsequent runs  
**Status**: Known limitation - individual sleeve flags only persist if Refresh is run after placement  
**Workaround**: MEPMARK service now checks if sleeve exists in Revit, not just XML flags  

---

## Best Practices

1. **Always use `ClusterSleeveInstanceId`** (integer) for cluster sleeve tracking, NOT `ClusterSleeveId` (ElementId with XmlIgnore)
2. **Reset flags BEFORE placement/clustering** to detect deleted sleeves
3. **Save XML immediately after setting cluster flags** (already implemented in `MarkClashZonesAsClusterResolvedWithSleeveId`)
4. **Check both flag AND ElementId existence** to handle deleted sleeves gracefully
5. **Use diagnostic logging** (`mepmark_debug.log`, `cluster_debug.log`) to trace flag states
6. **Apply MEPMARK to BOTH cluster and individual sleeves** to handle mixed scenarios

---

## Related Files

- **Models**: `Models/ClashZone.cs` - Flag definitions
- **Individual Placement**: `Services/UniversalSleevePlacerService.cs` - Sets `IsResolved`
- **Cluster Placement**: `Services/UniversalClusterService.cs` - Sets `IsClusterResolved`
- **MEPMARK Application**: `Services/MarkParameterService.cs` - Reads flags to find sleeves
- **Flag Reset**: `Services/UniversalClusterService.cs:ResetClusterFlagsForDeletedSleeves()`
- **Cluster Check**: `Services/UniversalClusterService.cs:IsClusterAlreadyExists()`

