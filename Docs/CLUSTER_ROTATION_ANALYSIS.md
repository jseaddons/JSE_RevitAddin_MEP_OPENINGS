# Cluster Rotation Logic Analysis

## Problem Statement

Cluster sleeves are not aligning correctly with individual sleeves for rotated-axis cable trays on floors. The cluster rotation logic is more complex than individual sleeve rotation, causing misalignment.

## Current Behavior

### Individual Sleeves (Working Correctly)
**Location**: `Services/UniversalSleevePlacerService.cs` lines 6109-6206

For **cable trays on floors**:
1. Uses `clashZone.MepElementRotationAngle` **directly** from database (line 6157)
2. Applies rotation: `ElementTransformUtils.RotateElement(_doc, sleeveInstance.Id, rotationAxis, rotationAngle)` (line 6192)
3. **NO 90° offset** - uses MEP rotation angle as-is
4. **NO conditional logic** - same behavior for all angles (0°, 45°, 135°, etc.)

**Example from logs**:
- Sleeve 1086234: `MepElementRotationAngle = -135.0°` → Applied as `-135.0°`
- Sleeve 1086212: `MepElementRotationAngle = 45.0°` → Applied as `45.0°`

### Cluster Sleeves (Current Implementation)
**Location**: `Services/Clustering/Rotation/ClusterRotationService.cs` lines 198-290

1. **Determines rotation angle**:
   - Gets `MepElementRotationAngle` from each clash zone (line 212)
   - Calculates **average** of all rotation angles (line 225)
   - For cluster 1086312: `(-135° + 45°) / 2 = -45° = 315°` → normalized to `135°`

2. **Applies conditional logic**:
   - Checks if angle is "straight axis" (0°, 90°, 180°, 270°) vs "rotated axis" (45°, 135°, etc.)
   - For cable trays + rotated axis: **skips 90° offset**
   - For cable trays + straight axis: **applies 90° offset**

**Location**: `Services/Clustering/Placement/ClusterPlacementService.cs` lines 442-458

3. **Rotation application**:
   - Cluster 1086312: Rotation = 135° (rotated axis)
   - Category = Cable Trays
   - Logic: `isRotatedAxis = true` → `skipOffset = true`
   - Applied rotation: `135.0°` (no 90° offset)

## The Problem

**Individual sleeves don't use conditional logic** - they always use `MepElementRotationAngle` directly for cable trays on floors, regardless of whether it's 0°, 45°, 135°, etc.

**Cluster sleeves use conditional logic** - they check if the angle is "straight axis" vs "rotated axis" and apply different behavior.

This creates a **mismatch**:
- Individual sleeve at 0°: Uses 0° rotation (no offset)
- Cluster with average 0°: Detects "straight axis" → applies 90° offset → uses 90° rotation
- **Result**: Cluster is rotated 90° more than individual sleeves!

## Root Cause

The cluster rotation logic was designed to handle a special case (straight-axis cable trays needing 90° offset), but this logic doesn't match how individual sleeves actually work.

**Individual sleeves** use `MepElementRotationAngle` directly because:
- The rotation angle is already calculated correctly during refresh
- It represents the actual MEP element orientation in world coordinates
- No additional offset is needed

**Cluster logic** tries to be "smart" by:
- Averaging rotation angles (correct)
- But then applying conditional 90° offset based on axis type (incorrect - doesn't match individual behavior)

## Proposed Solution

### Option 1: Match Individual Sleeve Logic Exactly (Recommended)

**For cable trays on floors**: Always use the average `MepElementRotationAngle` directly, **no conditional 90° offset logic**.

**For ducts/pipes on floors**: Use average `MepElementRotationAngle + 90°` (always apply 90° offset).

**Implementation**:
1. Remove the `IsStraightAxisAlignedAngle()` check for cable trays
2. For cable trays: Always use rotation angle directly (same as individual sleeves)
3. For ducts/pipes: Always add 90° offset (same as individual sleeves would if they had this logic)

### Option 2: Use Individual Sleeve Rotation Logic Directly

Instead of averaging rotation angles, use the same rotation logic that individual sleeves use:
- For each sleeve in cluster, get its `MepElementRotationAngle`
- Apply the same rotation logic as individual sleeves
- For cluster, use the average of these angles

But this still requires averaging, so Option 1 is simpler.

## Sizing Issue

The cluster bounding box calculation (741.4mm x 341.4mm) seems correct based on union of rotated bounding boxes. However, the rotation misalignment may make it appear incorrect visually.

**Next Steps**:
1. Fix rotation logic to match individual sleeves
2. Verify sizing after rotation is corrected
3. If sizing is still incorrect, investigate bounding box calculation

## Code Changes Implemented (2025-11-28)

1. **`ClusterPlacementService.cs` lines 442-458**: ✅ **FIXED** - Removed conditional `IsStraightAxisAlignedAngle()` logic for cable trays. Now always uses rotation angle directly for cable trays (matches individual sleeves), ducts/pipes still get 90° offset.

2. **`ClusterPlacementService.cs` lines 1187-1220**: ✅ **FIXED** - Simplified `ApplyRotation` method. Cable trays on floors now always skip 90° offset (matches individual sleeve behavior), with clearer logging.

3. **Logic Change**: 
   - **Before**: `skipOffset = isCableTrayCategory && isRotatedAxis` (conditional)
   - **After**: `skipOffset = isCableTrayCategory` (always true for cable trays)

## Testing

After fix, verify:
1. ✅ Cluster 1086312 rotation should match individual sleeves (135° average, applied as 135° with no 90° offset)
2. ✅ Cluster bounding box size should match union of individual sleeve bounding boxes
3. ✅ Cluster placement point should align with individual sleeve placement points
4. ✅ All cable tray clusters (straight and rotated axis) should use rotation angle directly

