# Cluster Placement Debug Log

Tracking the investigation into wall/framing cluster offsets.

## Completed

- **Placement logging**: Added diagnostic output in `ClusterPlacementService.PlaceClusterSleeve` that records the incoming `placementPoint`, recomputed centroid, bbox min/max, and deltas.
- **Depth handling**: Removed the forced wall-thickness override so `Depth` now uses the bounding-box through-wall dimension, matching the methodology doc.
- **Delegate midpoint logging**: `RefactoredClusterService` now logs the raw midpoint and rotated bbox min/max coming directly from `_getClusterBoundingBox`, so we can compare upstream data with the recomputed centroid.
- **Wall/framing midpoint override**: When the host type is `Wall` or `Structural Framing`, `PlaceClusterSleeve` now replaces the delegate midpoint with `ComputeClusterMidpoint`’s value and logs the delta. This permanently guards against the 33 mm intersection-point offset without touching floor/rotated flows.
- **Protection note**: Do not remove the wall override or the diagnostic logging unless `_getClusterBoundingBox` is updated to emit the corrected centroid; both are now part of the regression-defense toolkit for legacy clustering.
- **Floor axis envelope**: Straight axis-aligned clusters (rotation ≈ 0°) now build their width/height/depth and midpoint from the persisted sleeve bounding boxes instead of averaged intersection points. This catches mixed X/Y floor sleeves (e.g., cluster 1083116) without needing a separate override.

## Next Checks

- Compare the logged `Width/Height/Depth` + rotation angles against the legacy `UniversalClusterService` output to confirm we are feeding the same values into the RectangularOpening families.
- Verify whether the family origin/anchor matches the centroid assumptions. If not, note required parameter swaps or rotation offsets.
- Once we update the upstream bounding-box service, rerun the logs to confirm both midpoints now match, then consider removing the override for wall/framing hosts only if the delta stays < 1 mm across multiple projects.

