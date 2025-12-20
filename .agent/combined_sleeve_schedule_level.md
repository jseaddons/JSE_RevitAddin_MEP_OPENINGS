# Combined Sleeve Schedule Level & Bottom of Opening - Implementation Summary

## Overview
Added "Schedule Level" and "Bottom of Opening" parameter setting for combined sleeves, matching the implementation used for cluster sleeves. This ensures combined sleeves have proper level and elevation parameters for scheduling and coordination.

## What Was Added

### 1. Schedule Level Parameter Setting
**Location:** `Services/Combined/CombinedSleevePlacementService.cs` (after Instance ID setting, before AUTO-JOIN)

**Logic:**
1. Extract level name from first sleeve in proximity group
2. Find the level in the document by name
3. Set "Schedule Level" parameter (tries multiple variations)
4. Supports both ElementId and String storage types

**Code:**
```csharp
// Get first sleeve from proximity group
var firstSleeve = group.Sleeves.FirstOrDefault();

// Extract level name from Individual or Cluster sleeve
if (firstSleeve.SourceData is ClashZone cz)
{
    levelName = cz.MepElementLevelName;
}

// Find level in document
Level mepLevel = new FilteredElementCollector(_doc)
    .OfClass(typeof(Level))
    .Cast<Level>()
    .FirstOrDefault(l => string.Equals(l.Name, levelName, StringComparison.OrdinalIgnoreCase));

// Set Schedule Level parameter (tries multiple variations)
var scheduleLevelParam = placedInstance.LookupParameter("Schedule of Level")
                     ?? placedInstance.LookupParameter("Schedule Level")
                     ?? placedInstance.LookupParameter("ScheduleLevel")
                     ?? placedInstance.Symbol?.LookupParameter("Schedule of Level")
                     ?? placedInstance.Symbol?.LookupParameter("Schedule Level")
                     ?? placedInstance.Symbol?.LookupParameter("ScheduleLevel");

if (scheduleLevelParam != null && !scheduleLevelParam.IsReadOnly)
{
    if (scheduleLevelParam.StorageType == StorageType.ElementId)
    {
        scheduleLevelParam.Set(mepLevel.Id);
    }
    else if (scheduleLevelParam.StorageType == StorageType.String)
    {
        scheduleLevelParam.Set(mepLevel.Name);
    }
}
```

---

### 2. Bottom of Opening Parameter Setting
**Location:** `Services/Combined/CombinedSleevePlacementService.cs` (after Schedule Level setting)

**Formula:**
```
Bottom of Opening = Elevation from Level - (Height / 2)
```

**Critical Sequencing:**
1. **Schedule Level must be set FIRST**
2. Revit automatically calculates "Elevation from Level" after Schedule Level is set
3. Read "Elevation from Level" from parameter
4. Calculate Bottom of Opening using formula
5. Set "Bottom of Opening" parameter

**Code:**
```csharp
// Step 1: Read "Elevation from Level" (Revit calculates this after Schedule Level is set)
var elevationParam = placedInstance.LookupParameter("Elevation from Level");
double? elevationFromLevel = null;

if (elevationParam != null && elevationParam.StorageType == StorageType.Double)
{
    elevationFromLevel = elevationParam.AsDouble();
}

if (elevationFromLevel.HasValue)
{
    // Step 2: Calculate Bottom of Opening
    double bottomOfOpening = elevationFromLevel.Value - (height / 2.0);
    
    // Step 3: Set "Bottom of Opening" parameter (tries multiple variations)
    var bottomParam = placedInstance.LookupParameter("Bottom Of Opening")
                   ?? placedInstance.LookupParameter("Bottom of Opening")
                   ?? placedInstance.LookupParameter("BottomOfOpening")
                   ?? placedInstance.Symbol?.LookupParameter("Bottom Of Opening")
                   ?? placedInstance.Symbol?.LookupParameter("Bottom of Opening")
                   ?? placedInstance.Symbol?.LookupParameter("BottomOfOpening");
    
    if (bottomParam != null && !bottomParam.IsReadOnly)
    {
        bottomParam.Set(bottomOfOpening);
    }
}
```

---

## Parameter Name Variations

### Schedule Level
The code tries these parameter names (in order):
1. "Schedule of Level" (most common)
2. "Schedule Level"
3. "ScheduleLevel"

Checks both **instance** and **symbol** (type) parameters.

### Bottom of Opening
The code tries these parameter names (in order):
1. "Bottom Of Opening" (capital O in "Of")
2. "Bottom of Opening"
3. "BottomOfOpening"

Checks both **instance** and **symbol** (type) parameters.

---

## Comparison with Cluster Sleeves

| Feature | Cluster Sleeves | Combined Sleeves | Status |
|---------|----------------|------------------|--------|
| **Schedule Level** | ✅ Set from first ClashZone's MEP element level | ✅ Set from first sleeve's MEP element level | ✅ **MATCHED** |
| **Elevation from Level** | ✅ Revit calculates automatically | ✅ Revit calculates automatically | ✅ **MATCHED** |
| **Bottom of Opening** | ✅ Formula: Elevation - (Height/2) | ✅ Formula: Elevation - (Height/2) | ✅ **MATCHED** |
| **Parameter Variations** | ✅ Tries multiple names | ✅ Tries multiple names | ✅ **MATCHED** |
| **Error Handling** | ✅ Try-catch with logging | ✅ Try-catch with logging | ✅ **MATCHED** |
| **Sequencing** | ✅ Schedule Level → Bottom of Opening | ✅ Schedule Level → Bottom of Opening | ✅ **MATCHED** |

---

## Example Execution Flow

### Before (Missing Parameters):
```
Combined Sleeve Created:
- Width: 600mm
- Height: 300mm
- Depth: 200mm
- Schedule Level: ❌ NOT SET
- Elevation from Level: ❌ NOT SET
- Bottom of Opening: ❌ NOT SET
```

### After (With Parameters):
```
Combined Sleeve Created:
- Width: 600mm
- Height: 300mm
- Depth: 200mm
- Schedule Level: ✅ "Level 1"
- Elevation from Level: ✅ 2500mm (Revit calculated)
- Bottom of Opening: ✅ 2350mm (2500 - 300/2)
```

---

## Logging

### Success Logs:
```
[CombinedSleevePlacement] ✅ Set Schedule Level to 'Level 1' on combined sleeve 123456
[CombinedSleevePlacement] ✅ Set Bottom of Opening = 2350.0mm (Elevation from Level = 2500.0mm, Height = 300.0mm)
```

### Warning Logs:
```
[CombinedSleevePlacement] ⚠️ 'Bottom of Opening' parameter not found or read-only
[CombinedSleevePlacement] ⚠️ 'Elevation from Level' parameter not available - skipping Bottom of Opening calculation
[CombinedSleevePlacement] ⚠️ Error setting Schedule Level: <error message>
```

---

## Edge Cases Handled

### 1. **No Level Name Available**
- If first sleeve doesn't have `MepElementLevelName`, Schedule Level is not set
- Bottom of Opening calculation is skipped (depends on Schedule Level)

### 2. **Level Not Found in Document**
- If level name doesn't match any level in document, Schedule Level is not set
- Logged as warning

### 3. **Parameter Not Found**
- Tries multiple parameter name variations
- If none found, logs warning and continues

### 4. **Read-Only Parameter**
- Checks `!parameter.IsReadOnly` before setting
- Skips if read-only

### 5. **Elevation from Level Not Available**
- If Revit hasn't calculated "Elevation from Level" yet, Bottom of Opening is skipped
- Logged as warning

### 6. **Cluster Sleeve as First Sleeve**
- Currently returns `null` for level name (TODO: implement cluster level extraction)
- This is a known limitation - cluster sleeves don't store level name directly

---

## Known Limitations

### 1. **Cluster Sleeve Level Extraction**
**Issue:** If the first sleeve in a proximity group is a cluster sleeve, we cannot extract the level name.

**Current Behavior:**
```csharp
else if (firstSleeve.SourceData is ClusterSleeveData cluster)
{
    levelName = null; // TODO: Implement cluster level extraction if needed
}
```

**Workaround:** The code will skip Schedule Level setting for this case.

**Future Enhancement:** Query the ClashZone table to get the level from cluster constituents.

---

## Testing Checklist

- [ ] Verify Schedule Level is set correctly for combined sleeves
- [ ] Verify Bottom of Opening is calculated correctly
- [ ] Test with different level names
- [ ] Test with combined sleeves from individual sleeves
- [ ] Test with combined sleeves from cluster sleeves (known limitation)
- [ ] Test with combined sleeves from mixed (individual + cluster)
- [ ] Verify parameter name variations work
- [ ] Test with opening families that have different parameter names
- [ ] Verify error handling when parameters are missing
- [ ] Check logs for success and warning messages

---

## Benefits

✅ **Consistency:** Combined sleeves now have the same parameters as cluster sleeves
✅ **Scheduling:** Combined sleeves can be scheduled by level
✅ **Coordination:** Bottom of Opening helps with coordination drawings
✅ **Robustness:** Handles multiple parameter name variations
✅ **Error Handling:** Graceful degradation if parameters are missing
✅ **Logging:** Clear logs for debugging

---

## Files Modified

- `Services/Combined/CombinedSleevePlacementService.cs`
  - Added Schedule Level setting logic (lines ~340-397)
  - Added Bottom of Opening calculation logic (lines ~399-445)

---

## Next Steps

1. **Test** the implementation with real combined sleeves
2. **Verify** Schedule Level and Bottom of Opening appear in Revit properties
3. **Implement** cluster level extraction for the TODO item (if needed)
4. **Document** any issues found during testing
