# Pipe Sleeve Add-in — README

Purpose
-------
This document explains the workflow, objectives, debugging history, code changes, run instructions, and lessons learned for the Pipe Sleeve placement add-in in this repository. It is intended for developers and BIM/MEP engineers who will run, maintain, and extend the add-in.

Objectives
----------
- Collect MEP elements (pipes, ducts, trays, conduits) from the host model and from visible linked models, restricted to the active 3D view's section box.
- Collect structural host + linked elements (walls, structural framing, floors) inside the same section box, using the same solid-based filter to avoid coordinate-space mismatches.
- Perform robust intersection tests between MEP runs and structural solids and place sleeve families where intersections occur.
- Keep section-box filtering active at all times (to avoid scanning full models and crashing Revit on large projects).
- Provide diagnostic logging to prove correct behavior and ease debugging.

Shared coordinates — importance and how we use them
-----------------------------------------------
- Why shared coordinates matter:
  - Linked Revit models (MEP, architectural, structural) can each have their own internal coordinate system. When you try to compare geometry (bounding boxes, points, lines) across documents you must be confident both sides are expressed in the same coordinate space. If they are not, simple AABB comparisons or point tests become invalid and will silently reject valid candidates.
  - Many false negatives in spatial queries come from comparing untransformed AABBs or points that live in link-local space against host-space geometry.

- How this project handles coordinates:
  - Collection uses a single spatial contract: build the active 3D view's section-box solid in host coordinates and, for each link, transform that section solid into the link's local coordinates and apply `ElementIntersectsSolidFilter` inside the link's document. This guarantees the filter logic runs entirely inside each document's native space and avoids cross-space AABB errors.
  - When we need to perform intersection checks in host space (for unified placement and caching), we transform MEP geometry (pipe endpoints, bounding boxes) from link-space into host-space and run host-space intersection overloads. The transform steps are logged with `[TransformDebug]` so you can confirm numeric consistency in the logs.
  - We avoid comparing transformed AABB boxes across spaces. Where possible we prefer transforming the smaller item (line/point) into the larger coordinate space and perform the numeric test there.

Performance & efficiency — how objectives are met
-----------------------------------------------
- Section-box-first filtering:
  - The active 3D view section box is used as the primary spatial filter for all collectors. Implementation uses `ElementIntersectsSolidFilter` against a `Solid` built from the section box. For links the same solid is transformed into link-local coordinates and the same filter is applied. This is orders of magnitude faster than scanning full models.

- Category limiting and minimal geometry extraction:
  - We only collect the categories we need (pipes, ducts, trays, conduits for MEP; walls, floors, structural framing for structure) to keep collector sets small.
  - Heavy geometry extraction (solids, face tests) is done after the cheap spatial filters pass. We use bounding-box pre-filters only when both the subject and the candidate are known to be expressed in the same coordinate space.

- Caching and re-use:
  - Extracted solids and tessellations are cached keyed by element id + link identity so subsequent intersection checks reuse geometry rather than re-extract it from the API.

- Transform minimization:
  - We minimize costly transform operations by transforming the section solid into each link once, and by transforming individual MEP lines into host space only when required for host-space intersection. Re-using a small number of transforms reduces numeric and CPU overhead.

- Diagnostic gating and light-weight logging:
  - Diagnostics (`[SectionBoxDiag]`, `[TransformDebug]`, `[Intersect]`) are informative but can be noisy. Add a runtime or compile-time gate (recommended next step) to enable verbose logs only during debugging runs.

- Safety-first pre-filters:
  - Where safe (same coordinate space) we perform a quick AABB overlap test to avoid full solid extraction. This remains a useful micro-optimization but must not be applied across coordinate spaces.

High-level workflow
-------------------

High-level workflow
-------------------
1. Initialize diagnostics logger early so collector diagnostics are captured.
2. Collect MEP elements via `Helpers.MepElementCollectorHelper.CollectMepElementsVisibleOnly(Document)`:
   - Uses `SectionBoxHelper` to build a section-box solid from the active 3D view and apply `ElementIntersectsSolidFilter`.
   - Handles linked models by transforming the section solid into each link's local coordinates.
   - Returns `List<(Element element, Transform? transform)>` where `transform==null` means host element.
3. Collect structural elements via `Services.MepIntersectionService.CollectStructuralElementsForDirectIntersectionVisibleOnly(Document, log)`:
   - Uses the same `MepElementCollectorHelper.CollectElementsVisibleOnly(doc, categories)` internally to filter structural categories with the solid-based section-box.
   - Returns `List<(Element element, Transform? transform)>` of structural elements (host+linked).
4. For each MEP element (pipe):
   - If the pipe originates in a linked doc, transform its location (line endpoints) and bounding box into host coordinates and call the host-line intersection overload.
   - Otherwise, call the normal intersection routine.
   - Spatial pre-filter by bounding box (host-space) is applied before extracting solids.
   - Geometry is cached and heavy face tests are performed only inside the section box.
5. Place sleeve families at the computed intersection centers; log placements and failures.

Files changed (concise)
-----------------------
- `Helpers/MepElementCollectorHelper.cs`
  - Central MEP collection; now logs raw vs filtered counts and samples (diagnostics) and continues to use `SectionBoxHelper` for solid-based filtering.
- `Helpers/SectionBoxHelper.cs`
  - Section-box solid creation and `FilterElementsBySectionBox` logic; added per-link diagnostics (transformed solid volume, passing counts, sample ids).
- `Services/MepIntersectionService.cs`
  - Intersection engine and structural collection.
  - Structural collection now re-uses `MepElementCollectorHelper.CollectElementsVisibleOnly(...)` to ensure the same section-box filtering is used for structural elements.
  - Added pre-filter diagnostics and preserved the host-line overload for linked MEP elements.
- `Commands/PipeSleeveCommand.cs`
  - Moved logger initialization earlier to capture collector diagnostics and wired `DebugLogger` as the logging backend.

How to build and run
--------------------
1. Build in PowerShell (use the configured configuration for your Revit target):

```powershell
dotnet build JSE_RevitAddin_MEP_OPENINGS.csproj -c "Debug R24" /p:Platform="Any CPU"
```

2. Deploy the generated add-in assembly (DLL) into your Revit add-ins folder or use your usual installer.
3. In Revit, open the 3D view you want to work in and enable the section box.
4. Run the `Pipe Sleeve` command from the add-in.
5. Check logs in the repository `Log` folder (files named `PipeSleeve_*.log`).

Key log markers and how to read them
-----------------------------------
- `[SectionBoxDiag] Raw elements - host=X, linked=Y` — raw counts before section box clipping.
- `[SectionBoxDiag] Raw sample: Id=..., Category=..., IsLinked=True` — raw sample list (useful to confirm MEP presence).
- `[SectionBoxDiag] Processing link='...' TransformOrigin=(x,y,z) TransformedSolidVol=V` — section solid mapped into link coordinates; `V` should be > 0.
- `[SectionBoxDiag] link='...' passingCount=N` — number of elements inside the transformed solid inside that link.
- `[SectionBoxDiag] link sample passing id=XXXX` — sample passing ids; these can be opened inside the linked model for verification.
- `[SectionBoxDiag] Filtered elements - host=A, linked=B` — counts returned to caller after section-box filtering.
- `[TransformDebug] Pipe <ID> transformed to host line Start=..., End=...` — a pipe from a link transformed to host coordinates before intersection.
- `Spatial filtering: processed N, skipped M elements via bounding box check` — per-pipe pre-filtering statistics.
- `[Intersect] Solid face count = F` and `[Intersect] Found K intersection point(s). First: ...` — face-level intersection diagnostics.

Common scenarios and interpretation
-----------------------------------
- Pre-filter count > 0 and passingCount = 0
  - The transformed section-box AABB did not overlap elements. Check `TransformedSolidVol` and sample bboxes in the log.
- passingCount > 0 and no intersections
  - Geometry face intersection may have failed (face types, Revit API behavior). Use `[Intersect]` diagnostics; if needed, fall back to tessellation/closest-distance checks.
- All counts zero
  - Confirm the active 3D view's section box and the visibility of linked instances (worksets, phases, link visibility settings).

Debug history (short timeline)
------------------------------
1. Symptom: add-in found pipes but reported zero intersections; placement count = 0.
2. Initial diagnosis: spatial pre-filter (AABB) was filtering out structural candidates due to coordinate-space mismatch between host and linked bounding boxes.
3. Instrumented collectors with `[SectionBoxDiag]` logs to observe raw vs filtered counts and per-link transformed solids.
4. Observed that raw MEP elements existed in links, and `SectionBoxHelper` identified passing MEP elements, but structural collection using AABB filters returned zero.
5. Fix: structural collection changed to reuse the solid-based `MepElementCollectorHelper` flow (the same logic used for MEP elements) so both MEP and structural collections use identical section-box logic.
6. Added diagnostics and transform-aware host-line overload so linked pipes are transformed before intersection.
7. Result: structural elements are collected correctly and intersections are detected.

8. Recent runtime transaction issue and fix (2025-08-19):
  - Symptom: Pipe sleeve placement attempted to create FamilyInstances but failed with "Starting a new transaction is not permitted." The `PipeSleeve` run logged placement exceptions and no instances were created.
  - Cause: `Services/PipeSleevePlacer.PlaceSleeve()` started its own Transaction for each sleeve while the calling command also created an active Transaction, producing a nested-transaction conflict.
  - Fix applied: the per-sleeve Transaction was removed from `Services/PipeSleevePlacer.PlaceSleeve()` so the placer now assumes the caller manages the Transaction lifecycle. This prevents nested transaction exceptions and allows placements to complete under the caller's transaction.
  - Verification: add-in rebuilt and tested; subsequent runs produced successful placements (check `Log/PipeSleeve_*.log` for PLACED/SUCCESS entries and created element ids).

Lessons learned
-------------
- Always use the same spatial-filtering contract for both sides of an intersection test (MEP vs structural). Mixing AABB tests across coordinate transforms leads to hard-to-find false negatives.
- Solid-based section-box filtering (transforming the section solid into link-local space and using ElementIntersectsSolidFilter) is more robust than manually transforming bounding boxes and comparing AABBs across transforms.
- Initialize loggers early so collector diagnostics are captured.
- For large models, keep section-box filtering active to avoid scanning entire models and crashing Revit.

Next steps & recommendations
----------------------------
- Add a small configuration flag to enable/disable verbose diagnostics (`[SectionBoxDiag]`, `[Intersect]` lines).
- Add support for additional structural categories (e.g., curtain walls, generic models) if your projects use non-standard categories for structural geometry.
- If face.Intersect returns inconsistent results across Revit versions, add a fallback intersection test using tessellated geometry or a line-to-solid closest-distance test.
- Consider a small unit/integration test harness that loads a test RVT and validates collector outputs (may be limited by Revit API licensing in CI).

How to reduce noise / disable diagnostics
----------------------------------------
Diagnostics are emitted via `Services.DebugLogger`. To silence them:
- Edit `Services.DebugLogger` (if present) to add a log level filter and set the default level to `Error`.
- Or remove / comment the `DebugLogger.Info(...)` calls in the helper/service files.

Contact / owning team
---------------------
- Repo: `JSE_RevitAddin_MEP_OPENINGS`
- Main files to review for future edits:
  - `Helpers/MepElementCollectorHelper.cs`
  - `Helpers/SectionBoxHelper.cs`
  - `Services/MepIntersectionService.cs`
  - `Commands/PipeSleeveCommand.cs`

Appendix: quick checklist for testing
-----------------------------------
- Open Revit and the intended RVT set.
- Open the 3D view you will use and enable the section box, align it to the area of interest.
- Run the add-in.
- Inspect the `Log/PipeSleeve_*.log` file and verify:
  - `[SectionBoxDiag] Raw elements` shows non-zero linked MEP counts.
  - `[SectionBoxDiag] link '...' passingCount` shows expected structural and MEP passes.
  - `[Intersect] Found` lines appear for pipes that actually cross structural geometry.


---
Generated: 2025-08-18

