# Revit geometry persistence rules

**Do not calculate. Only persist from Revit.**

These fields must **only** be set from actual Revit geometry (placed elements). Never derive them from dimensions, placement point + width/height, or formulas.

## Protected fields (ClashZone)

| Data | Rule | Allowed source |
|------|------|----------------|
| **SleeveCorner1X/Y/Z … SleeveCorner4X/Y/Z** | Only from Revit | `BatchSleeveCornerExtractor` (uses `SleeveCornerCalculationService.CalculateCornersFromInstance` on the placed FamilyInstance). Cluster placement may write corners from the cluster instance geometry. |
| **BoundingBoxMin/Max**, **RotatedBoundingBoxMin/Max** | Only from Revit | `element.get_BoundingBox(null)` and, for rotated bbox, that bbox transformed to element-local space. No half-width/half-height math. |

## Where persistence is allowed

- **Individual sleeves:** `BulkPlacementService.UpdateZoneGeometry` — uses Revit bbox only; does **not** set sleeve corners (those come from `BatchSleeveCornerExtractor` after placement).
- **Corner extraction:** `Services/Persistence/batch_corner_extractor.cs` — extracts corners from Revit via `CalculateCornersFromInstance` and persists via `ClashZoneRepository.BatchUpdateSleeveCorners`.
- **Clusters:** Cluster placement may set corners/bbox from the placed cluster FamilyInstance geometry only.

## What not to do

- Do **not** set `SleeveCorner*` or `RotatedBoundingBox*` from placement point + width/height + rotation.
- Do **not** add new code paths that “calculate” corners or bbox from dimensions; always read from Revit element geometry.

## No fallback

- **SleevePersistenceService.SaveSleeveCorners** persists corners only when geometry extraction succeeds (no math fallback). Same rule everywhere: only corner extraction from Revit.

## Enforcement

- **Cursor rule:** `.cursor/rules/revit-geometry.mdc` guides AI and developers when editing `.cs` files.
- When editing placement, persistence, or clustering, check this file and the XML on `ClashZone` so these rules stay enforced.
