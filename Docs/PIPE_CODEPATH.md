PIPE SLEEVE CODE PATH — Overview and Responsibilities

This document captures the runtime code path for pipe sleeve placement in this repository, the main classes involved, data shapes, the coordinate-space contract, and the historical double-transform pitfalls with recommended enforcement.

Files of interest
- Commands/PipeSleeveCommand.cs — command-level orchestration and tuple collection.
- Services/MepIntersectionService.cs — computes intersections between MEP elements and structural solids.
- Services/PipeSleevePlacerService.cs — orchestrates per-pipe processing, projection, duplication checks and calls the placer.
- Services/PipeSleevePlacer.cs — low-level placer that creates FamilyInstance, sets parameters, rotates and recenters.
- Services/DuctSleevePlacerService.cs — similar flow for ducts (see parity comments below).
- Services/OpeningDuplicationChecker.cs — central duplication-suppression logic.
- Helpers/WallCenterlineHelper.cs, Helpers/HostLevelHelper.cs — supporting utilities.

High-level responsibilities
- PipeSleeveCommand:
  - Collects: List<(Pipe pipe, Transform? transform)> — transform is present when pipe originates from a linked doc and maps link->active.
  - Collects structural hosts: List<(Element host, Transform? transform)>.
  - Calls PipeSleevePlacerService.PlaceAllPipeSleeves() to process.

- MepIntersectionService:
  - Accepts a MEP element + structural hosts list. For hosts that are in a linked document it "transforms" the structural solid into active-doc coordinates (e.g. SolidUtils.CreateTransformed).
  - Computes intersections and returns List<(Element host, BoundingBoxXYZ hostBBox, XYZ intersectionPoint)>.
  - IMPORTANT: intersectionPoint is returned in active-document coordinates because the service transforms linked host solids before intersection.

- PipeSleevePlacerService:
  - Iterates pipe tuples and calls MepIntersectionService.
  - For each intersection: determines host type (Wall/Floor/Framing), projects to wall centerline when appropriate and computes placePoint (all in active-doc coords).
  - Performs duplication suppression using OpeningDuplicationChecker.
  - Selects family symbol and calls _placer.PlaceSleeve(pipe, placePtToUse, direction, symbol, host).
  - MUST NOT re-apply link transforms to intersection-derived points. Doing so produces double-transform artifacts (large deltas).

- PipeSleevePlacer:
  - Creates the FamilyInstance at the exact XYZ passed in (expects active-doc coords).
  - Sets parameters (Diameter, Depth), rotates the instance and recenters it so the family geometry is centered through the host.
  - Logs instance id and final coordinates.

Data shapes and contract (enforce)
- Tuple semantics
  - (Pipe, Transform?) — transform is link->active mapping; exists when pipe originates in a linked document.
  - (Element host, Transform?) — hostTransform is link->active mapping for linked hosts.

- SINGLE SOURCE OF TRUTH: MepIntersectionService returns intersection XYZ in active-document coordinates.
  - All callers must assume that and must NOT call transform.OfPoint(intersectionPoint) on intersection results.
  - Transform.OfPoint should only be used when you compute a point from a source-element-local coordinate (for example: raw center from a linked damper instance). In that case transform.OfPoint converts that source-local point into active-doc coords exactly once.

Historical double-transform pattern (root cause)
- The intersection service transforms host solids to active-doc coords and computes intersectionPoint (active-doc coords).
- A caller mistakenly treated intersectionPoint as if it were in the link's local coords and applied transform.OfPoint(intersectionPoint) again. That yields a different coordinate (double-transform) and a large delta.
- This appeared in both `PipeSleevePlacerService` and `DuctSleevePlacerService` in code paths that tried to "help" missing transforms by finding host transforms and applying them to intersection/ projection points.

Fixes and enforcement applied (recommended)
- Add XML documentation to `MepIntersectionService.FindIntersections` clarifying the return coordinate space (active-doc).
- In placer services we added guard logs and removed transform.OfPoint(...) calls for intersection-derived points to prevent double transforms.
- Centralize transform-finding for diagnostic lookups (do not apply transforms to intersection coords).

Where to look for remaining transform calls
- Grep for: OfPoint( intersectionPoint ) or hostTransform.OfPoint(...)
- Important files previously changed: `Services/PipeSleevePlacerService.cs`, `Services/DuctSleevePlacerService.cs`, `Services/FireDamperSleevePlacerService.cs`.

Duplication of responsibilities
- The split between `*PlacerService` (orchestrator) and `*Placer` (low-level) is appropriate (SRP). Avoid moving projection or coordinate-space conversion into the low-level placer.
- Keep duplication checks centralized in `OpeningDuplicationChecker` and call it from the service layer only.

Quick checklist for reviewers and maintainers
- Ensure `MepIntersectionService` docs say: "returns active-document coordinates".
- When computing points from elements that come from links, transform the source-local point once at the service boundary into active-doc with the tuple transform, then use that active-doc point for all further checks.
- Do not apply hostTransform.OfPoint to points returned by `MepIntersectionService`.
- Add instance-id logging inside each placer after creating FamilyInstance so created elements are easy to find in Revit.

Next steps I can perform (pick one):
- A: Add inline XML comments to `MepIntersectionService.FindIntersections` and add guard-logs to any remaining placer files.
- B: Create a short `Docs/CONTRACTS.md` that lists all coordinate-space contracts across services.
- C: Add instance-id logging to each placer class (pipe/duct/damper) so created sleeves are traceable.

If you want me to proceed with one of the items above, reply with A/B/C and I will implement, compile, and report results.
