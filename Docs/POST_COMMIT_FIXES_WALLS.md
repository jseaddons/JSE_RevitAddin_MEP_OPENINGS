## Post-Commit Fixes Summary (Walls only) — Ducts, Cable Trays, Pipes

Purpose: capture the exact functional edits we made after the last commit so you can restore and reapply quickly. Scope is placement on Walls (not Floors/Framing) and the upstream intersection flow that feeds it.

### 1) Intersection detection (feeds all categories on Walls)

What we fixed
- Use UI-chosen Reference files for MEP collection and Host files for structural collection. Robust name matching (strip parentheses, extension, `_detached`, underscores/dashes, case-insensitive).
- Oriented section-box: transform the active view’s section box into each link’s coordinate space and filter elements by that oriented box.
- Discipline-to-category mapping so filters like “Ventilation/HVAC” and “Electrical” still collect the right Revit categories.

Key code (high-level)
```csharp
// Services/IntersectionDetectionService.cs
// 1) Robust matching of selected reference files (MEP) vs host files (structural)
var selectedReferenceFiles = FilterUiStateProvider.GetSelectedReferenceFiles?.Invoke() ?? new List<string>();
Func<string,string> norm = ... // lowers, strips (..), ext, _detached, condenses spaces
var linkTitleNorm = norm(linkDoc.Title);
bool refMatch = selectedReferenceFiles.Count == 0 || selectedReferenceFiles.Any(f => linkTitleNorm.Contains(norm(f)) || norm(f).Contains(linkTitleNorm));
bool hostMatch = selectedHostFiles == null || selectedHostFiles.Count == 0 || selectedHostFiles.Any(f => linkTitleNorm.Contains(norm(f)) || norm(f).Contains(linkTitleNorm));

// 2) Oriented box (per link)
Transform inv = link.GetTotalTransform().Inverse;
XYZ linkMin = inv.OfPoint(modelMin);
XYZ linkMax = inv.OfPoint(modelMax);
Outline linkOutline = new Outline(actualMin, actualMax);

// 3) Discipline mapping
if (selectedMepCategories.Contains("Ventilation") || selectedMepCategories.Contains("HVAC")) { add Ducts/Fittings/Accessory/Terminal; }
if (selectedMepCategories.Contains("Electrical")) { add CableTray/TrayFitting/Conduit/ConduitFitting; }
if (selectedMepCategories.Contains("Plumbing")) { add PipeCurves/PipeFitting/PipeAccessory; }
```

Logs now show how many MEP/host elements were collected per link and the oriented section-box bounds.

### 2) Clash creation — placement point policy (Walls)

What we fixed
- We reverted to using the actual intersection center as the authoritative sleeve placement point for all hosts to avoid “host-center stacking”.
- We added a post-placement safety move that nudges the instance back to the exact intended point if the family snaps elsewhere.

Key code (concept)
```csharp
// Services/ClashZoneService.cs (CreateClashZone)
// placementPoint = intersectionPoint;  // do not override with host bbox center
clashZone.SleevePlacementPoint = intersectionPoint;

// Services/UniversalSleevePlacerService.cs
XYZ placementPoint = clashZone.SleevePlacementPoint;
// Prefer bbox center if it differs significantly (diagnostic guard)
if (clashZone.ClashBoundingBox != null)
{
    var bb = clashZone.ClashBoundingBox; var bbCenter = new XYZ((bb.Min.X+bb.Max.X)/2.0, (bb.Min.Y+bb.Max.Y)/2.0, (bb.Min.Z+bb.Max.Z)/2.0);
    if (placementPoint.DistanceTo(bbCenter) > 0.001) placementPoint = bbCenter;
}
var adjustedPlacementPoint = placementPoint + placementOffset; // strategy offset (dampers, etc.)

// After creation, force location to the exact point (some families snap)
var loc = sleeveInstance.Location as LocationPoint;
if (loc != null && loc.Point.DistanceTo(adjustedPlacementPoint) > 0.0001)
    loc.Move(adjustedPlacementPoint - loc.Point);
```

### 3) Orientation rules on Walls (final state)

Common
- Orientation is decided using the stored `StructuralElementNormal` (wall normal) and the MEP direction from `ClashZone`.
- Y-oriented wall/framing: allowed to rotate. X-oriented wall: skip rotation (unless special pipe rules apply).

Cable Trays (Walls)
- If Y-oriented wall: rotate by 90°.
- If X-oriented wall: no rotation.

Pipes (Walls)
- If wall is Y-oriented:
  - Circular sleeve: no rotation needed.
  - Rectangular sleeve: rotate 90°.
- If wall is X-oriented: align to MEP direction (rotate to atan2 of MEP XY).

Ducts (Walls)
- Follow general rule based on wall orientation:
  - Y-oriented wall: rotate 90° for rectangular; round (if any) does not need rotation.
  - X-oriented wall: no rotation.

Key code (abbreviated)
```csharp
// Services/UniversalSleevePlacerService.cs → SetSleeveOrientation
bool isWallY = Math.Abs(structNormal.Y) > Math.Abs(structNormal.X);
if (isPipe && isXWall) rotate to atan2(MEP);
else if (isCableTray && isWallY) rotate 90°;
else if (isWallY && isRectangular) rotate 90°; else no rotation.
```

### 4) Family selection (Walls)

- Pipes now select by `PipeOpeningType` (Circular/Rectangular) from UI/global settings — not by duct shape.
- Cable trays and ducts use rectangular opening families on walls by default.

```csharp
// Services/UniversalSleevePlacerService.cs → SelectUniversalFamily
// Pipes: decide via clashZone.PipeOpeningType (Circular vs Rectangular)
```

### 5) Pipe sizing and depth (host-agnostic but used on Walls)

- Diameter: Outside Diameter + 2×Insulation + 2×UI clearance.
- Depth: explicitly written — for walls this corresponds to wall thickness logic upstream.

```csharp
// finalDiameter = OD + 2*ins + 2*clearance;
// Depth parameter is set alongside width/height/diameter.
```

### 6) UI gating and state

- OK button enabled only when unresolved clash zones > 0 (loaded from the same auto-save path used by Refresh).
- Placement processes only zones marked `IsEligibleByCurrentUi` (we persist the flag rather than deleting zones for unselected hosts; floors are preserved for later runs).

```csharp
// Models/ClashZone.cs
public bool IsEligibleByCurrentUi { get; set; } = true;

// Refresh: mark eligibility for new zones based on current host types
foreach (var cz in newClashZones)
    cz.IsEligibleByCurrentUi = allowedHostTypes.Contains(cz.StructuralElementType);

// Placement: process only eligible zones
var zonesToProcess = clashZones.Where(cz => cz.IsEligibleByCurrentUi).ToList();
```

### 7) Logging to verify quickly

- Intersection log prints: section box per link, collected counts, and per-MEP intersection results.
- Placement log prints: chosen placement point → adjusted point; “moved to exact” correction; orientation decisions.

Checklist to reapply on a clean restore
1. IntersectionDetectionService.cs
   - Reference vs Host link collection (normalization, oriented box)
   - Discipline-to-category mapping for Ventilation/HVAC and Electrical
2. ClashZoneService.cs
   - Use `intersectionPoint` for `SleevePlacementPoint`
3. UniversalSleevePlacerService.cs
   - Orientation rules for Walls (cable tray 90° on Y-walls; pipes special cases; ducts general rule)
   - Pipe family selection by `PipeOpeningType`
   - Pipe sizing formula and depth write
   - Post-placement location correction to exact clash point
4. Gating/State
   - OK enabled when unresolved > 0 using auto-loaded filter
   - `IsEligibleByCurrentUi` flag set at Refresh and respected at placement

This document is intentionally narrowly scoped to reapplying wall-specific behavior for Ducts, Cable Trays, and Pipes along with the upstream intersection and placement-point fixes needed to make them reliable.

### Preserve-on-Restore Checklist (must not lose these)

1) Section Box Filter (oriented, host/shared coords)
- Files: `Services/IntersectionDetectionService.cs`, `Services/MepIntersectionService.cs`
- Keep: transforming active `View3D` section box to each link via `link.GetTotalTransform().Inverse`, using oriented min/max, bounding-box filter for both MEP and hosts. Logs: "Section box in linked file …", element counts.

2) SetSleeveParameters (write Depth; pipe sizing)
- File: `Services/UniversalSleevePlacerService.cs`
- Keep: pipe diameter = OD + 2×ins + 2×UI clearance; always attempt to set Depth. For framing, the written Depth must be `clashZone.StructuralElementThickness` (from framing type ‘b’).

3) SetSleeveOrientation (final rules on walls)
- File: `Services/UniversalSleevePlacerService.cs`
- Keep: wall/framing guard by structural normal; cable tray 90° on Y-walls; pipes align to MEP on X-walls; Y-walls: circular pipes no rotation, rectangular 90°; ducts follow guard (rectangular 90° on Y-walls). Logs for each branch.

4) Framing Depth Extraction (‘b’ type param in linked)
- File: `Services/ClashZoneService.cs` (GetElementThickness)
- Keep: resolve `FamilyInstance` type via `GetTypeId()` when `Symbol` is null; read type parameter key ‘b’ (case-insensitive) as the only authoritative breadth; log `[FRAMING-THICKNESS]` and `[CLASH-THICKNESS-ASSIGN]`. This feeds `clashZone.StructuralElementThickness` used by placement.

5) Placement Point Policy (avoid host-center stacking)
- Files: `Services/ClashZoneService.cs`, `Services/UniversalSleevePlacerService.cs`
- Keep: `SleevePlacementPoint = intersectionPoint`; after placement, force `LocationPoint` to the intended adjusted point if snapping occurs. Optional bbox-center sanity check can remain as diagnostic.


