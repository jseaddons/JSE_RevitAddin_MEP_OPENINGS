# Clustering Angle Improvement Discussion

## Current Problem

The current clustering algorithm has a fundamental limitation:

1. **Axis-Aligned Bounding Box Calculation**: The cluster bounding box is calculated by taking the min/max of all individual sleeve bounding boxes in the **model coordinate system** (X/Y/Z axes). This creates a box that's always aligned to 0° or 90°.

2. **Post-Placement Rotation**: The cluster sleeve is rotated AFTER the bounding box dimensions are set. This means:
   - If individual sleeves are at 45°, the axis-aligned bounding box will be much larger than needed
   - Rotating the cluster sleeve afterward doesn't help because the width/height are already set based on the oversized axis-aligned box

3. **Result**: Cluster sleeves don't follow the actual outline of individual sleeves, especially when sleeves are at non-orthogonal angles (e.g., 30°, 45°, 60°).

## Example Scenario

**Individual Sleeves:**
- Sleeve 1: 200mm × 200mm at 45° angle
- Sleeve 2: 200mm × 200mm at 45° angle, 500mm away

**Current Approach:**
- Axis-aligned bounding box: ~700mm × ~700mm (oversized)
- Cluster sleeve created: 700mm × 700mm
- Then rotated to 45°
- **Result**: Cluster sleeve is much larger than needed

**Desired Approach:**
- Calculate bounding box in rotated coordinate system (45°)
- Cluster sleeve: ~500mm × ~200mm (fits actual outline)
- **Result**: Cluster sleeve follows the actual outline of individual sleeves

## Proposed Solution

### Option 1: Rotated Coordinate System Bounding Box (Recommended)

1. **Determine Dominant Angle**:
   - Analyze all sleeves in the cluster to find the dominant rotation angle
   - Use the average or most common `MepElementRotationAngle` from individual sleeves
   - For wall/framing sleeves, use the orientation-based angle (X=90°, Y=0°)

2. **Transform to Rotated Coordinate System**:
   - Create a transformation matrix for the dominant angle
   - Transform all sleeve bounding box corners to the rotated coordinate system
   - Calculate min/max in the rotated coordinate system

3. **Calculate Cluster Dimensions**:
   - Width = max(transformed X) - min(transformed X)
   - Height = max(transformed Y) - min(transformed Y)
   - Depth = max(transformed Z) - min(transformed Z)

4. **Place Cluster Sleeve**:
   - Use calculated dimensions (already in rotated coordinate system)
   - Place at midpoint
   - Apply rotation (if needed for alignment)

### Option 2: Convex Hull Approach

1. **Collect All Sleeve Corner Points**:
   - Get all 8 corners of each sleeve's bounding box
   - Transform to rotated coordinate system if needed

2. **Calculate Convex Hull**:
   - Find the convex hull of all points
   - This gives the actual outline of the cluster

3. **Calculate Bounding Box from Convex Hull**:
   - Find the oriented bounding box (OBB) of the convex hull
   - Use OBB dimensions for cluster sleeve

### Option 3: Average Angle with Tolerance Grouping

1. **Group Sleeves by Similar Angles**:
   - Group sleeves with angles within ±5° tolerance
   - Use the average angle for each group

2. **Calculate Bounding Box per Group**:
   - Use rotated coordinate system for each group
   - Merge groups if they're close enough

## Implementation Details

### Key Changes Needed

1. **`ClusterBoundingBoxServices.GetClusterBoundingBox()`**:
   - Add parameter for rotation angle
   - Transform bounding boxes to rotated coordinate system before calculating min/max

2. **`UniversalClusterService.GetClusterBoundingBoxFromXml()`**:
   - Determine dominant angle from individual sleeves
   - Transform bounding boxes before calculating union

3. **`UniversalClusterService.PlaceClusterSleeve()`**:
   - Calculate rotation angle BEFORE calculating bounding box
   - Pass rotation angle to bounding box calculation methods

### Code Structure

```csharp
// 1. Determine dominant angle
double dominantAngle = DetermineDominantAngle(cluster);

// 2. Transform bounding boxes to rotated coordinate system
var transformedBoxes = cluster.Select(s => 
    TransformBoundingBox(s.BoundingBox, dominantAngle)
).ToList();

// 3. Calculate bounding box in rotated coordinate system
var (width, height, depth, mid) = CalculateRotatedBoundingBox(transformedBoxes);

// 4. Place cluster sleeve with correct dimensions
// (rotation already accounted for in dimensions)
```

## Benefits

1. **Accurate Sizing**: Cluster sleeves match the actual outline of individual sleeves
2. **Better Space Utilization**: No oversized cluster sleeves
3. **Respects Individual Sleeve Angles**: Works for any angle (0°, 30°, 45°, 60°, 90°, etc.)
4. **More Professional Results**: Cluster sleeves look like they were designed to fit the actual sleeve layout

## Considerations

1. **Angle Consistency**: What if sleeves in a cluster have different angles?
   - Use average angle
   - Use most common angle
   - Group by angle and create separate clusters

2. **Performance**: Transformation calculations add overhead
   - Cache transformation matrices
   - Only calculate for clusters with >1 sleeve

3. **Backward Compatibility**: Existing clusters may have different dimensions
   - This is acceptable - new clusters will be more accurate
   - Old clusters can be regenerated if needed

## Next Steps

1. **Discuss Approach**: Choose between Option 1, 2, or 3
2. **Implement Prototype**: Start with Option 1 (simplest)
3. **Test with Real Projects**: Verify accuracy with various angles
4. **Refine as Needed**: Adjust based on results

