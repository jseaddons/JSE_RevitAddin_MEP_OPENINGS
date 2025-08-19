# Cable Tray Sleeve Placement Fix Summary

## Problem Analysis
Cable tray sleeve placement was failing with coordinate mismatch errors like:
```
intersection=(38.878,-338.788,-18.869), placePoint=(113.620,-34.360,-18.869), mmDistance=95545.3, allowedOffset=105.0
```

This indicated a massive coordinate calculation error causing placement points to be ~95,000mm away from intersection points.

## Root Cause Identified
The fundamental issue was in **wall centerline calculation logic difference** between:

### Cable Tray Placer (BROKEN)
```csharp
// Complex projection logic that was causing coordinate errors
XYZ placePoint = GetWallCenterlinePoint(wall, intersection);
```

### Duct Placer (WORKING) 
```csharp
// Simple offset calculation using wall normal and thickness
XYZ n = normal; // Wall normal vector  
XYZ wallVector = n.Multiply(-wallThickness);
XYZ placePoint = intersection + wallVector.Multiply(0.5);
```

## Solution Applied
**Replaced complex `GetWallCenterlinePoint()` projection with exact duct logic pattern:**

### In `Services\CableTraySleevePlacer.cs`
**BEFORE (Complex projection):**
```csharp
XYZ placePoint = GetWallCenterlinePoint(wall, intersection);
```

**AFTER (Simple duct-style offset):**
```csharp
// Use exact same logic as working duct placer
XYZ wallVector = n.Multiply(-wallThickness);
XYZ placePoint = intersection + wallVector.Multiply(0.5);
```

## Why This Fix Works

1. **Proven Pattern**: Duct sleeve placement works reliably with this exact calculation
2. **Coordinate Consistency**: No complex projections that can introduce coordinate transform errors
3. **Simple Math**: `intersection + offset` is straightforward and predictable
4. **Transaction Safety**: Follows the lesson from README_PIPE.md - placers assume active transaction exists

## Key Lessons Applied from README_PIPE.md

1. **Transaction Management**: Commands manage transactions, placers assume they exist
2. **Keep It Simple**: Don't over-engineer when proven simple patterns exist
3. **Copy Working Logic**: When something works (duct placement), use the exact same pattern

## Testing Status
- ✅ Build successful with warnings only (no errors)
- ⏳ Ready for runtime testing in Revit
- 🎯 Expected: Cable tray sleeves should now place correctly without coordinate mismatches

## Files Modified
1. `Services\CableTraySleevePlacer.cs` - Applied exact duct wall centerline calculation logic
2. Previous GPT-4o changes in `Commands\CableTraySleeveCommand.cs` remain intact (coordinate transforms for structural intersections)

## Next Steps
1. Test cable tray sleeve placement in Revit
2. Verify no more "mmDistance=95545.3" type errors
3. Confirm sleeves place at correct wall centerline positions
