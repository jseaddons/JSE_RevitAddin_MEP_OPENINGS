# RCS (Relative Coordinate System) Implementation Plan for Walls/Framing

## Overview
Store bounding boxes in **wall-aligned RCS** instead of WCS for all walls and framing. This eliminates rotation logic and provides accurate cluster sizes for angled walls.

## Benefits

### 1. **No Rotation Logic Needed**
- Bounding boxes already in wall-aligned coordinates
- Cluster size calculation is direct (no rotation)
- Dimension mapping: Width = RCS_X, Depth = RCS_Y, Height = RCS_Z

### 2. **Works for Angled Walls**
- 45° walls, 30° walls, any angle → accurate bounding boxes
- No oversized clusters for angled walls
- Tighter, more accurate cluster sleeves

### 3. **Simpler Code**
- No need to determine rotation angle for walls
- No need to apply rotation transformations
- Direct dimension mapping

## RCS Coordinate System Definition

For **walls and framing**:
- **RCS X-axis** = Along wall direction (WallDirection vector)
- **RCS Y-axis** = Through wall (perpendicular to wall direction in XY plane)
- **RCS Z-axis** = Vertical (same as WCS Z)

**Transformation:**
```csharp
// Wall direction vector (already in database: ClashZone.WallDirection)
XYZ wallDir = clashZone.WallDirection; // e.g., (0.707, 0.707, 0) for 45° wall

// Wall normal (perpendicular in XY plane)
XYZ wallNormal = new XYZ(-wallDir.Y, wallDir.X, 0).Normalize();

// RCS basis vectors
XYZ rcsX = wallDir;           // Along wall
XYZ rcsY = wallNormal;        // Through wall
XYZ rcsZ = XYZ.BasisZ;        // Vertical

// Transformation matrix: WCS → RCS
// To transform point from WCS to RCS:
//   rcsX = point.DotProduct(rcsX)
//   rcsY = point.DotProduct(rcsY)
//   rcsZ = point.DotProduct(rcsZ)
```

## Implementation Steps

### Step 1: Add RCS Bounding Box Columns to Database
**File:** `Data/SleeveDbContext.cs`

Add new columns to `ClashZones` table:
- `SleeveBoundingBoxRCS_MinX` (REAL)
- `SleeveBoundingBoxRCS_MinY` (REAL)
- `SleeveBoundingBoxRCS_MinZ` (REAL)
- `SleeveBoundingBoxRCS_MaxX` (REAL)
- `SleeveBoundingBoxRCS_MaxY` (REAL)
- `SleeveBoundingBoxRCS_MaxZ` (REAL)

**Migration:** Add to `EnsureSchemaUpgraded` method

### Step 2: Add RCS Properties to ClashZone Model
**File:** `Models/ClashZone.cs`

Add properties:
```csharp
public double SleeveBoundingBoxRCS_MinX { get; set; } = 0.0;
public double SleeveBoundingBoxRCS_MinY { get; set; } = 0.0;
public double SleeveBoundingBoxRCS_MinZ { get; set; } = 0.0;
public double SleeveBoundingBoxRCS_MaxX { get; set; } = 0.0;
public double SleeveBoundingBoxRCS_MaxY { get; set; } = 0.0;
public double SleeveBoundingBoxRCS_MaxZ { get; set; } = 0.0;
```

### Step 3: Create RCS Transformation Service
**New File:** `Services/Clustering/Geometry/WallRcsTransformer.cs`

```csharp
public static class WallRcsTransformer
{
    /// <summary>
    /// Transform WCS bounding box to wall-aligned RCS
    /// </summary>
    public static BoundingBoxXYZ TransformToRcs(BoundingBoxXYZ wcsBbox, XYZ wallDirection)
    {
        // Calculate RCS basis vectors
        XYZ rcsX = wallDirection.Normalize();
        XYZ rcsY = new XYZ(-rcsX.Y, rcsX.X, 0).Normalize();
        XYZ rcsZ = XYZ.BasisZ;
        
        // Transform all 8 corners of bounding box
        XYZ[] wcsCorners = new XYZ[]
        {
            new XYZ(wcsBbox.Min.X, wcsBbox.Min.Y, wcsBbox.Min.Z),
            new XYZ(wcsBbox.Max.X, wcsBbox.Min.Y, wcsBbox.Min.Z),
            new XYZ(wcsBbox.Min.X, wcsBbox.Max.Y, wcsBbox.Min.Z),
            new XYZ(wcsBbox.Max.X, wcsBbox.Max.Y, wcsBbox.Min.Z),
            new XYZ(wcsBbox.Min.X, wcsBbox.Min.Y, wcsBbox.Max.Z),
            new XYZ(wcsBbox.Max.X, wcsBbox.Min.Y, wcsBbox.Max.Z),
            new XYZ(wcsBbox.Min.X, wcsBbox.Max.Y, wcsBbox.Max.Z),
            new XYZ(wcsBbox.Max.X, wcsBbox.Max.Y, wcsBbox.Max.Z)
        };
        
        // Transform to RCS
        double[] rcsX_vals = wcsCorners.Select(c => c.DotProduct(rcsX)).ToArray();
        double[] rcsY_vals = wcsCorners.Select(c => c.DotProduct(rcsY)).ToArray();
        double[] rcsZ_vals = wcsCorners.Select(c => c.DotProduct(rcsZ)).ToArray();
        
        return new BoundingBoxXYZ
        {
            Min = new XYZ(rcsX_vals.Min(), rcsY_vals.Min(), rcsZ_vals.Min()),
            Max = new XYZ(rcsX_vals.Max(), rcsY_vals.Max(), rcsZ_vals.Max()),
            Enabled = true
        };
    }
    
    /// <summary>
    /// Transform RCS point back to WCS
    /// </summary>
    public static XYZ TransformToWcs(XYZ rcsPoint, XYZ wallDirection, XYZ origin)
    {
        XYZ rcsX = wallDirection.Normalize();
        XYZ rcsY = new XYZ(-rcsX.Y, rcsX.X, 0).Normalize();
        XYZ rcsZ = XYZ.BasisZ;
        
        // Inverse transformation: RCS → WCS
        return origin + rcsPoint.X * rcsX + rcsPoint.Y * rcsY + rcsPoint.Z * rcsZ;
    }
}
```

### Step 4: Update Individual Sleeve Placement to Store RCS Bounding Box
**File:** `Services/UniversalSleevePlacerService.cs`

**Location:** After `sleeve.get_BoundingBox(null)` (around line 1760)

```csharp
var actualBbox = sleeve.get_BoundingBox(null);
if (actualBbox != null)
{
    // Store WCS bounding box (keep for backward compatibility)
    zone.SetSleeveBoundingBox(actualBbox);
    
    // ✅ NEW: Store RCS bounding box for walls/framing
    if (clashZone.StructuralElementType == "Wall" || 
        clashZone.StructuralElementType == "Walls" ||
        clashZone.StructuralElementType == "Structural Framing")
    {
        if (clashZone.WallDirection != null && clashZone.WallDirection != XYZ.Zero)
        {
            var rcsBbox = WallRcsTransformer.TransformToRcs(actualBbox, clashZone.WallDirection);
            zone.SleeveBoundingBoxRCS_MinX = rcsBbox.Min.X;
            zone.SleeveBoundingBoxRCS_MinY = rcsBbox.Min.Y;
            zone.SleeveBoundingBoxRCS_MinZ = rcsBbox.Min.Z;
            zone.SleeveBoundingBoxRCS_MaxX = rcsBbox.Max.X;
            zone.SleeveBoundingBoxRCS_MaxY = rcsBbox.Max.Y;
            zone.SleeveBoundingBoxRCS_MaxZ = rcsBbox.Max.Z;
        }
    }
}
```

### Step 5: Update Cluster Size Calculation to Use RCS
**File:** `Services/Clustering/Rotation/ClusterRotationService.cs`

**Method:** `CalculateRotatedBoundingBox`

**Change:** For walls/framing, use RCS bounding boxes directly (no rotation needed)

```csharp
// For walls/framing: Use RCS bounding boxes (already wall-aligned)
if (isWallHost || isFramingHost)
{
    var rcsBboxes = new List<(XYZ min, XYZ max)>();
    foreach (var sleeveData in cluster)
    {
        var cz = _getClashZoneFunc(sleeveData.SleeveInstanceId, xmlFilePath) as ClashZone;
        if (cz == null) continue;
        
        // Check if RCS bounding box is available
        if (cz.SleeveBoundingBoxRCS_MinX != 0 || cz.SleeveBoundingBoxRCS_MaxX != 0 ||
            cz.SleeveBoundingBoxRCS_MinY != 0 || cz.SleeveBoundingBoxRCS_MaxY != 0 ||
            cz.SleeveBoundingBoxRCS_MinZ != 0 || cz.SleeveBoundingBoxRCS_MaxZ != 0)
        {
            rcsBboxes.Add((
                new XYZ(cz.SleeveBoundingBoxRCS_MinX, cz.SleeveBoundingBoxRCS_MinY, cz.SleeveBoundingBoxRCS_MinZ),
                new XYZ(cz.SleeveBoundingBoxRCS_MaxX, cz.SleeveBoundingBoxRCS_MaxY, cz.SleeveBoundingBoxRCS_MaxZ)
            ));
        }
    }
    
    if (rcsBboxes.Count > 0)
    {
        // Simple union in RCS (no rotation needed!)
        double minX = rcsBboxes.Min(b => b.min.X);
        double minY = rcsBboxes.Min(b => b.min.Y);
        double minZ = rcsBboxes.Min(b => b.min.Z);
        double maxX = rcsBboxes.Max(b => b.max.X);
        double maxY = rcsBboxes.Max(b => b.max.Y);
        double maxZ = rcsBboxes.Max(b => b.max.Z);
        
        double width = maxX - minX;   // RCS X = along wall = Width
        double height = maxZ - minZ;  // RCS Z = vertical = Height
        double depth = maxY - minY;  // RCS Y = through wall = Depth
        
        // Midpoint in RCS
        XYZ rcsMid = new XYZ((minX + maxX) / 2, (minY + maxY) / 2, (minZ + maxZ) / 2);
        
        // Transform midpoint back to WCS for placement
        var firstClashZone = cluster[0].ClashZone as ClashZone;
        if (firstClashZone?.WallDirection != null)
        {
            // Use first sleeve's center as origin for transformation
            XYZ origin = new XYZ(
                firstClashZone.SleevePlacementPointActiveDocumentX,
                firstClashZone.SleevePlacementPointActiveDocumentY,
                firstClashZone.SleevePlacementPointActiveDocumentZ
            );
            XYZ wcsMid = WallRcsTransformer.TransformToWcs(rcsMid, firstClashZone.WallDirection, origin);
            return (width, height, depth, wcsMid, null, null, null, null, null, null);
        }
    }
}
```

### Step 6: Update Dimension Mapping (No Rotation Needed!)
**File:** `Services/Clustering/Placement/ClusterPlacementService.cs`

**Method:** `SetSizeParameters`

**Change:** For walls, use RCS dimensions directly (no rotation, no swapping)

```csharp
if (groupKey.hostType == "Wall" || groupKey.hostType == "Structural Framing")
{
    // ✅ RCS: Dimensions are already in wall-aligned coordinates
    // Width = RCS X (along wall)
    // Depth = RCS Y (through wall) - override with wall thickness
    // Height = RCS Z (vertical)
    openingWidth = width;   // RCS X → Width parameter
    openingHeight = depth;   // RCS Z → Height parameter
    openingDepth = height;   // RCS Y → Depth parameter (will be overridden)
    
    // Override depth with wall thickness
    if (wallThickness > 0)
    {
        openingDepth = wallThickness;
    }
}
```

### Step 7: Remove Rotation Logic for Walls
**File:** `Services/Clustering/Rotation/ClusterRotationService.cs`

**Method:** `DetermineRotationAngle`

**Change:** For walls/framing, return 0.0 (no rotation needed - RCS handles it)

```csharp
if (isWallHost || isFramingHost)
{
    // ✅ RCS: No rotation needed - bounding boxes are already in wall-aligned coordinates
    return 0.0;
}
```

**File:** `Services/Clustering/Placement/ClusterPlacementService.cs`

**Change:** Don't apply rotation for walls (rotationAngle will be 0.0)

```csharp
// ✅ RCS: No rotation needed for walls - bounding boxes already in wall-aligned coordinates
// Only apply rotation for floors (rotated axis/non-straight)
if (Math.Abs(rotationAngle) > 1e-6 && groupKey.hostType != "Wall" && groupKey.hostType != "Structural Framing")
{
    ApplyRotation(doc, inst, placementPoint, rotationAngle);
}
```

### Step 8: Update Database Repository
**File:** `Data/Repositories/ClashZoneRepository.cs`

**Method:** `MapClashZone`

Add RCS bounding box loading:
```csharp
clashZone.SleeveBoundingBoxRCS_MinX = GetDouble(reader, "SleeveBoundingBoxRCS_MinX");
clashZone.SleeveBoundingBoxRCS_MinY = GetDouble(reader, "SleeveBoundingBoxRCS_MinY");
clashZone.SleeveBoundingBoxRCS_MinZ = GetDouble(reader, "SleeveBoundingBoxRCS_MinZ");
clashZone.SleeveBoundingBoxRCS_MaxX = GetDouble(reader, "SleeveBoundingBoxRCS_MaxX");
clashZone.SleeveBoundingBoxRCS_MaxY = GetDouble(reader, "SleeveBoundingBoxRCS_MaxY");
clashZone.SleeveBoundingBoxRCS_MaxZ = GetDouble(reader, "SleeveBoundingBoxRCS_MaxZ");
```

**Method:** `UpdateSleeveBoundingBoxes` (or create new method)

Add RCS bounding box saving:
```csharp
public void UpdateSleeveBoundingBoxesRcs(Guid clashZoneId, 
    double minX, double minY, double minZ, 
    double maxX, double maxY, double maxZ)
{
    // SQL UPDATE with RCS columns
}
```

## Testing Checklist

- [ ] Straight X-wall: Cluster size accurate
- [ ] Straight Y-wall: Cluster size accurate
- [ ] 45° angled wall: Cluster size accurate (not oversized)
- [ ] 30° angled wall: Cluster size accurate
- [ ] Multiple angled walls: Each wall clusters correctly
- [ ] Floors: Still use WCS (unchanged)
- [ ] Backward compatibility: WCS bounding boxes still available

## Migration Strategy

1. **Phase 1:** Add RCS columns to database (backward compatible)
2. **Phase 2:** Start storing RCS bounding boxes for new sleeves
3. **Phase 3:** Update cluster calculation to prefer RCS, fallback to WCS
4. **Phase 4:** Remove rotation logic for walls (after testing)

## Benefits Summary

✅ **No rotation logic** - RCS handles wall alignment  
✅ **Accurate for angled walls** - No oversized clusters  
✅ **Simpler code** - Direct dimension mapping  
✅ **Works for all wall angles** - 0°, 45°, 30°, any angle  
✅ **Backward compatible** - WCS still stored for floors/fallback

