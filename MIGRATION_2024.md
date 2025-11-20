# Revit 2024 Migration Summary (Dual Build: 2023 & 2024)

## Scope
- Supported versions now: 2023 (legacy baseline) and 2024 (new)
- Removed: 2021, 2022, 2025, 2026 (can be reintroduced by restoring csproj blocks)

## Key Changes
1. Pruned `JSE_RevitAddin_MEP_OPENINGS.csproj` configurations to `Debug/Release R23` and `Debug/Release R24` only.
2. Added `Helpers/VersionInfo.cs` for compile-time version flags.
3. Added `Helpers/GeometryOptionsFactory.cs` to standardize geometry extraction (higher fidelity in 2024).
4. Replaced all direct `new Options()` instantiations in active services with factory method.
5. Removed `"2025"` from runtime dependency scan list in `Application.cs`.
6. Strengthened backup folder exclusion patterns to avoid duplicate type ambiguity.

## Geometry Extraction Behavior
- 2023 build: Medium detail, excludes non-visible objects, references not computed (preserves known good behavior & performance).
- 2024 build: Fine detail only (no non-visible objects, no reference computation) to keep performance similar and focus fix on unit/tolerance consistency.

## How To Add Revit 2025 Later
1. Re-add `PropertyGroup` blocks for `R25` (TargetFramework likely `net8.0-windows`).
2. Introduce compile constant `REVIT2025_OR_GREATER` if new branching required (otherwise reuse `REVIT2024_OR_GREATER`).
3. Update `Configurations` property to include `Debug R25;Release R25`.
4. Add `"2025"` to `revitVersions` array in `Application.cs`.
5. Verify toolkit/package versions (pin exact 2025 versions across Nice3point packages & Revit API DLL references).
6. Adjust `GeometryOptionsFactory` only if 2025 requires changed extraction semantics.

## Validation Checklist
- Build `Debug R24`: ensure no ambiguous type or missing reference errors.
- Build `Debug R23`: confirm no regressions (intersection logic unchanged for 2023 pathway).
- Spot test intersection detection in Revit 2024: verify previously missed intersections now detected.
- Confirm SQLite dependencies still copied (check logs: `[SQLite] ✅ Copied dependency`).

## Reverting (If Needed)
- Restore prior csproj from source control; remove factory usage lines if rollback necessary.
- Delete new helper files if fully reverting.

## Notes
- Backup and archive folders aggressively excluded to prevent duplicate compilation & ambiguous calls.
- Factory centralizes geometry settings to keep future version adjustments low-risk.

_Last updated: 2025-11-20_
