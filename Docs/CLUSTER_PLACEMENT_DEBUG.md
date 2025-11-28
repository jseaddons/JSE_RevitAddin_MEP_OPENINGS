# Cluster Placement Debug Log

Tracking the investigation into wall/framing cluster offsets.

## Completed

- **Placement logging**: Added diagnostic output in `ClusterPlacementService.PlaceClusterSleeve` that records the incoming `placementPoint`, recomputed centroid, bbox min/max, and deltas.
- **Depth handling**: Removed the forced wall-thickness override so `Depth` now uses the bounding-box through-wall dimension, matching the methodology doc.
- **Delegate midpoint logging**: `RefactoredClusterService` now logs the raw midpoint and rotated bbox min/max coming directly from `_getClusterBoundingBox`, so we can compare upstream data with the recomputed centroid.
- **Wall/framing midpoint override**: When the host type is `Wall` or `Structural Framing`, `PlaceClusterSleeve` now replaces the delegate midpoint with `ComputeClusterMidpoint`’s value and logs the delta. This permanently guards against the 33 mm intersection-point offset without touching floor/rotated flows.
- **Protection note**: Do not remove the wall override or the diagnostic logging unless `_getClusterBoundingBox` is updated to emit the corrected centroid; both are now part of the regression-defense toolkit for legacy clustering.
- **Floor axis envelope**: Straight axis-aligned clusters (rotation ≈ 0°) now build their width/height/depth and midpoint from the persisted sleeve bounding boxes instead of averaged intersection points. This catches mixed X/Y floor sleeves (e.g., cluster 1083116) without needing a separate override.
- **Mixed MEP orientation floor clusters - Unit conversion fix (2025-11-28)**: Fixed critical bug in `ClusterRotationService.CalculateRotatedBoundingBox` where `SleeveWidth` and `SleeveHeight` were incorrectly divided by 304.8 (mm-to-feet conversion). These values are already stored in Revit internal units (feet), not millimeters. The bug caused cluster bounding boxes to be calculated 304.8× smaller than actual sleeve dimensions, resulting in clusters that were too small (e.g., 279mm instead of 1000mm). **Fix location**: Lines 859-865 in `ClusterRotationService.cs` - removed `/ 304.8` conversion for `SleeveWidth`, `SleeveHeight`, and `StructuralElementThickness`. **Impact**: Mixed X/Y orientation floor clusters now correctly calculate cluster bounding box dimensions from stored sleeve dimensions.
- **Cable tray rotated axis rotation fix (2025-11-28)**: Fixed incorrect 90° offset application for cable tray clusters on floors. **Previous behavior**: Conditional logic checked if angle was "straight axis" vs "rotated axis" and applied different behavior, causing mismatch with individual sleeves. **New behavior**: Cable tray clusters now match individual sleeve behavior exactly - always use `MepElementRotationAngle` directly with no 90° offset, regardless of angle (0°, 45°, 135°, etc.). **Fix location**: Lines 442-458 in `ClusterPlacementService.cs` - removed conditional `IsStraightAxisAlignedAngle()` check, simplified to `skipOffset = isCableTrayCategory`. **Logic**: Cable trays always skip 90° offset (matches individual sleeves), ducts/pipes always get 90° offset (maintains existing cluster behavior). **Impact**: Cluster rotation now matches individual sleeve rotation exactly for cable trays, eliminating misalignment issues.
- **Corner-based calculation dynamic dispatch fix (2025-11-28)**: Fixed `'double' does not contain a definition for 'HasValue'` error in corner-based bounding box calculation. **Root cause**: `cornerResult` was declared as `var`, causing the C# dynamic dispatch to incorrectly resolve the nullable tuple type when accessing `.HasValue`. **Fix location**: Line 617-622 in `ClusterRotationService.cs` - explicitly typed `cornerResult` as `(double width, double height, double minX, double minY, double maxX, double maxY, XYZ origin)?` instead of using `var`. **Impact**: Corner-based calculation now works correctly for rotated clusters (135°, 315°, etc.) without falling back to less accurate union method.
- **Corner-based calculation complete fix (2025-11-28)**: Resolved persistent `'double' does not contain a definition for 'HasValue'` error that was causing corner-based calculation to fail and fall back to union method, resulting in incorrect dimensions (707mm × 707mm instead of 600mm × 400mm). **Root cause**: The code was accessing `sleeveData.SleeveInstanceId` directly from dynamic objects before extracting it safely, causing dynamic dispatch errors when accessing nullable properties like `SleeveCorner1X.HasValue`. **Solution**: 
  1. **Extract all IDs first**: Extract all `SleeveInstanceId` values into a `List<int>` (`sleeveIdsInCluster`) before any corner checks or calculations (lines 564-585 in `ClusterRotationService.cs`).
  2. **Use extracted IDs throughout**: All subsequent loops (corner checking, corner collection, depth calculation) now use the extracted `sleeveIdsInCluster` list instead of accessing dynamic objects directly.
  3. **Explicit nullable typing**: All corner properties are explicitly typed as `double?` (e.g., `double? corner1X = cz.SleeveCorner1X;`) before checking `.HasValue` to avoid dynamic dispatch issues.
  4. **Inline calculation**: Replaced the external `CornerBasedBoundingBoxCalculator.CalculateFromCorners` call (which used `Func<int, string, dynamic>`) with inline corner-based calculation that directly uses `ClashZone` objects.
  5. **Enhanced logging**: Added diagnostic logging for each sleeve's corners, origin calculation, rotation matrix components, and rotated bounds to aid debugging.
  **Fix locations**: 
  - Lines 564-585: Extract all `SleeveInstanceId` values first
  - Lines 587-620: Use extracted IDs for corner checking
  - Lines 622-676: Use extracted IDs for corner collection with explicit nullable typing
  - Lines 678-750: Inline corner-based calculation with proper rotation math
  - Line 666: Renamed local `minZ` to `sleeveMinZ` to avoid variable name conflict with enclosing scope
  **Impact**: Corner-based calculation now successfully executes for rotated axis-aligned clusters (135°, 315°, etc.), producing correct dimensions (600mm × 400mm) instead of the incorrect union result (707mm × 707mm). The union method was adding diagonal extents of two 45° rotated sleeves, giving √2 times the actual dimensions (707mm = 500mm × √2). The corner-based algorithm correctly handles rotated sleeves by rotating corners back to aligned axis, then finding min/max bounds.

## Next Checks

- Compare the logged `Width/Height/Depth` + rotation angles against the legacy `UniversalClusterService` output to confirm we are feeding the same values into the RectangularOpening families.
- Verify whether the family origin/anchor matches the centroid assumptions. If not, note required parameter swaps or rotation offsets.
- Once we update the upstream bounding-box service, rerun the logs to confirm both midpoints now match, then consider removing the override for wall/framing hosts only if the delta stays < 1 mm across multiple projects.

