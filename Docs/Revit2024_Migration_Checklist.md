# Revit 2024 Migration Checklist

## 1. Build / Packaging
- Create a dedicated build configuration that references Revit 2024 API assemblies while reusing the existing source tree.
- Define `REVIT_2024` (or similar) compilation symbol if compile-time branching becomes necessary.
- Rebuild add-in SHIMs (addin manifest, installer) to point at the correct DLL for each Revit major version.

## 2. Unit Handling
- `Services/RevitUnitConversionService.cs` now switches between Forge UnitTypeIds (Revit 2021+) and manual feet/mm conversions.
  - ✅ Already updated with runtime detection and fallbacks.
- `Services/MepIntersectionService.cs` wraps the shared converter and provides a fail-safe fallback to avoid type-initializer crashes.

### Services using `UnitUtils` / UnitTypeId (audit targets)
- `Services/UniversalClusterService.cs` — proximity tolerances, bounding boxes, overlap logging.
- `Services/UniversalSleevePlacerService.cs` — clearance, placement sizes, rounding logic.
- `Services/SleeveCoordinateService.cs` — coordinate normalization (search for `UnitUtils` when updating).
- `Services/Placement/` helpers (e.g. `SleevePlacementDimensionService`, `SleevePlacementValidationService`).
- `Services/FilterManagementService.cs` — when converting between saved XML values and Revit internal units.
- `refresh refactor/` services (`parameter_capture_service.cs`, etc.) where `UnitUtils` appears.

**Action:** Ensure each location either uses `RevitUnitConversionService.Instance` or catches exceptions when calling `UnitUtils.Convert*(..., UnitTypeId.*)` so it can fall back under Revit <2021 or when Forge IDs change.

## 3. Element & Category IDs
- Revit 2024 widens `BuiltInParameter` and `BuiltInCategory` enums (they map to `ElementId`).
  - Audit any casts to `int` or arithmetic comparisons. Replace with `ElementId` comparisons or calls to `new ElementId((int)...)` rebuilt against 2024.
- Verify repository classes (`Data/Repositories/…`) that persist `BuiltInCategory` or parameter IDs.

## 4. Transform / Coordinate Logic
- Review services that cache or manipulate `Transform` objects:
  - `Services/MepIntersectionService` (transform cache).
  - `Services/UniversalSleevePlacerService` (host transforms for placement).
  - `refresh refactor/intersection_processor.cs` (linked document transforms).
- No API changes reported, but run regression tests (linked model scenarios) after rebuilding against 2024 assemblies.

## 5. Extensible Storage & Schema Upgrades
- Revit 2024 improves schema conflict resolution. Confirm `GuidManager`, `GlobalIndexService`, and any custom storage code respond gracefully to file upgrade events.

## 6. Testing Plan
- Smoke tests in both Revit 2023 and 2024:
  - Sleeve placement end-to-end using typical MEP/structural models.
  - Refresh pipeline, parameter snapshot, and clustering flows.
  - Ensure logs show expected unit conversions (mm / feet) without errors.
- Capture any new warnings introduced by the 2024 API (obsolete members, etc.).

## 7. Deferred / Optional Work
- Python/Dynamo shims are optional; current architecture already exposes a runtime strategy via `IRevitUnitConversionService`.
- Only introduce separate DLLs if Autodesk ships breaking API changes in future versions.

> Keep this checklist updated as additional migration blockers appear during 2024 testing.

- **Placement & clustering** now load clash zones through `ClashZoneDataService` (SQLite primary, XML fallback).
  - Updated `SleevePlacementExternalEvent` and `UniversalSleevePlacementCommand` to stop parsing XML directly.
  - `UniversalClusterService` cache reload uses the same service; legacy XML parsing remains as a fallback for missing DB rows.
