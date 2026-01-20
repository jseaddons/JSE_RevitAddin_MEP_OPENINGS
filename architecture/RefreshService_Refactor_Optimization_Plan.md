# Refresh Service Refactor & Optimization Plan

_Last updated: 2025-11-12_

## 1. Context & Goals

The legacy `RefreshService` monolith (≈5,700 LOC) couples UI orchestration, clash detection, XML persistence, flag management, logging, and diagnostics. It is difficult to reason about, expensive to run on large models, and hard to evolve.

**Objectives**

1. Split responsibilities into focused helper classes with clear interfaces.
2. Reduce refresh runtime from minutes to seconds on unchanged models.
3. Cap memory footprint to <100 MB at 10 k clash zones.
4. Eliminate duplicate persistence branches (e.g., `Plumbing` vs `Plumbing_pipes`).
5. Preserve deterministic GUID + flag semantics; never double-place sleeves.
6. Maintain current behaviour until the new pipeline can be toggled in.

The “refresh refactor” scaffolding already exists (`RefreshContext`, `XmlCacheManager`, `ParameterCaptureService`, `ValidationService`, `PerformanceMonitor`). `IntersectionProcessor` is still pending. This document maps the remaining work.

---

## 2. Target Architecture

```
RefreshService (thin orchestrator, ≤200 LOC)
└─ SleevePlacementCoordinator (already exists)
└─ RefreshContext               (shared state/session cache)
└─ XmlCacheManager              (load once, reuse)
└─ ParameterCaptureService      (whitelisted metadata)
└─ ValidationService            (hash/timestamp skip)
└─ IntersectionProcessor        (lazy detection & geometry)
└─ ClashZonePersistenceService  (existing, fed normalized names)
└─ PerformanceMonitor           (phase timing & memory snapshots)
```

Each helper is responsible for one thing only; the orchestrator composes them based on mode (replace / replay / detect).

---

## 3. Refactor Phases

### Phase 1 – Structural Decomposition (Day 1–2)
1. **RefreshContext**  
   - Add members for: string pool, cached document path/hash, cached `DateTime` (file timestamp), XML cache handles, and mode flags (`Replay`, `Replace`, `FullDetection`).  
   - Provide helpers: `Intern(string)`, `HasDocumentChanged()`, `ResetSession()`.

2. **XmlCacheManager**  
   - Implement `LoadFilter(string baseName)`, `LoadGlobal(string category)`, caching results in `RefreshContext`.  
   - Ensure `Save` invalidates internal cache or updates entries in-place.

3. **Refactor `RefreshService.ExecuteRefreshInternal`**  
   - Instantiate `RefreshContext` at method start.  
   - Replace direct member fields (`existingClashZones`, `allFileCombosProcessed`, etc.) with context properties.  
   - Move early exit for replace mode into a helper (`ReplaceModeWorkflow`) returning `Result`.

4. **Unit Tests (minimal)**  
   - New helper classes should be testable without Revit (use stub interfaces where needed).

### Phase 2 – Performance Optimisations (Day 3–4)
1. **Parameter Snapshot Diet**  
   - Hardcode minimal whitelist (MEP: size/diameter/outside diam/system/level; Host: level/thickness/type/family) inside `ParameterCaptureService`.  
   - Use `RefreshContext.Intern` for every key/value before storing.

2. **ValidationService**  
   - Add `GetElementHash(ClashZone)` (HashCode.Combine of MEP & Host).  
   - Skip 3-point validation when stored hash matches current hash.  
   - Respect `HasDocumentChanged()` to bypass entire validation pass when model timestamp unchanged.

3. **Geometry Cache**  
   - Remove unconditional `MepIntersectionService.ClearGeometryCache()` calls from `RefreshService`.  
   - Only clear when `RefreshContext.HasDocumentChanged()` returns true.

4. **Replay / Detection decisions**  
   - Implement `IntersectionDecision` struct: `bool ShouldRunDetection`, `string Reason`.  
   - Use existing diagnostics to log the reason just once.

### Phase 3 – Intersection Processor (Day 5–6)
1. **Create `IntersectionProcessor`** (new file in `Services` or `refresh refactor`).  
   - `PrepareExistingZones()` → loads snapshot via `XmlCacheManager`, reconstructs vectors, returns `IntersectionDecision`.  
   - `RunDetectionIfNeeded()` → only instantiate `IntersectionDetectionService` when decision says so; uses `ValidationService` & `ParameterCaptureService`.  
   - `PostProcess()` → updates context caches; rebuilds placement snapshots for replay path.

2. **🔴 CRITICAL: Collector-Level Multi-Filter Optimization**  
   - **Problem**: Currently 5 filters (Section Box, Reference File, MEP Categories, Host File, Host Categories) are applied AFTER element collection, loading unnecessary elements into memory.  
   - **Solution**: Apply ALL 5 filters at `FilteredElementCollector` level using compound filters BEFORE loading elements.  
   - **Implementation**:
     ```csharp
     // In IntersectionProcessor.RunDetectionIfNeeded()
     // Build compound filter combining all 5 filters:
     var filters = new List<ElementFilter>();
     
     // Filter 1: Section Box
     if (sectionBoxOutline != null)
         filters.Add(new BoundingBoxIntersectsFilter(sectionBoxOutline));
     
     // Filter 2: MEP Categories
     filters.Add(new ElementCategoryFilter(mepCategories));
     
     // Filter 3: Reference Files (via ElementWorksetFilter or custom filter)
     // Filter 4: Host Files (via ElementWorksetFilter or custom filter)
     // Filter 5: Host Categories
     filters.Add(new ElementCategoryFilter(hostCategories));
     
     var compoundFilter = new LogicalAndFilter(filters);
     
     // Apply at collector level - only elements passing ALL filters are loaded
     var mepCollector = new FilteredElementCollector(doc)
         .WherePasses(compoundFilter)
         .WhereElementIsNotElementType();
     ```
   - **Expected Impact**: 
     - Reduce memory footprint by 60-80% (only load needed elements)
     - Reduce intersection checking time by 60-80% (fewer elements to check)
     - Faster collection phase (Revit API filters are optimized)
   - **Files to Modify**: 
     - `Services/IntersectionDetectionService.cs` (or new `IntersectionProcessor.cs`)
     - Refactor `CollectElements()` method to use compound filters
   - **Reference**: See `CONSOLIDATED_PENDING_OPTIMIZATIONS.md` task #7

3. **Integrate into `RefreshService` orchestrator**  
   - Replace inline detection code with `IntersectionProcessor` calls.  
   - Ensure `Replace` path returns before persistence; `Replay` path loads existing zones and returns without modifying XML.  
   - Only `FullDetection` path proceeds to persistence.

### Phase 4 – Persistence & Normalisation (Day 7)
1. **Base-name normalisation**  
   - Before calling `ClashZonePersistenceService.SaveClashZones`, resolve base name via `FilterNameHelper.GetBaseFilterName(_filterName, filter?.Name, category)`.  
   - Log both raw & normalized names (`[PERSIST-NAME] Raw='Plumbing_pipes', Base='Plumbing'`).

2. **Single Branch Guarantee**  
   - Update `ClashZonePersistenceService.BuildFilterFileName`/`BuildFilterGroupName` to accept normalized base names when provided (low risk).  
   - Add safeguard: if `_baseName` already ends with `_pipes`, also compute alternate base name and merge entries (diagnostic warning).

3. **Verify**  
   - Run refresh with “Adopt” both off and on; confirm Global XML keeps one `<Filter Name="Plumbing">`.  
   - Confirm category snapshots still written to `Plumbing_pipes.xml`, but main `Plumbing.xml` is synced.

### Phase 5 – Diagnostics & Polish (Day 8+)
1. **PerformanceMonitor**  
   - Wrap major phases (`LoadXml`, `Decision`, `Detection`, `Capture`, `Persistence`) with `Measure`.  
   - Record memory snapshots via `MemoryProfiler.TakeSnapshot` (no `Dispose`/`Reset` calls needed; null out when done).

2. **Removal of Legacy Blocks**  
   - Delete chunks in `RefreshService.cs` that are superseded by helpers.  
   - Keep docstrings and fallback logging until confident.

3. **Config Flag for new pipeline**  
   - Introduce toggle `OptimizationFlags.UseRefactoredRefresh` to let us switch between old and new flows during rollout.

---

## 4. Name Normalisation Details

The duplicate branch in Global XML (`Plumbing` vs `Plumbing_pipes`) appears whenever the base filter name includes the category suffix. Fix:

```csharp
var effectiveBaseName = FilterNameHelper.GetBaseFilterName(
    _filterName,                // e.g., "Plumbing_pipes"
    targetFilter?.Name,         // e.g., "Plumbing"
    category);                  // e.g., "Pipes"

var persistenceService = new ClashZonePersistenceService(...);
persistenceService.SaveClashZones(allClashZones, effectiveBaseName, targetFilter, allowStructuralUpdates: true);
```

`FilterNameHelper` should:  
1. Trim whitespace; strip extension;  
2. Remove trailing `_{category}` (case-insensitive);  
3. Return fallback to target filter name if `_filterName` missing or makes no sense.

Add logging in persistence to confirm:  
`[PERSIST-NAME] Raw='Plumbing_pipes', Normalized='Plumbing', Category='Pipes'`.

---

## 5. Sequencing & Ownership (Sample)

| Sequence | Task | Owner | Notes |
| --- | --- | --- | --- |
| 1 | RefreshContext enhancements | Dev A | 0.5 day |
| 2 | XmlCacheManager single-load | Dev A | 0.5 day |
| 3 | ParameterCaptureService whitelist | Dev B | 0.5 day |
| 4 | ValidationService hashes/timestamps | Dev B | 0.5 day |
| 5 | IntersectionProcessor implementation | Dev C (with Dev B) | 1.5 days |
| 6 | Orchestrator integration | Dev A + Dev B | 1 day |
| 7 | Persistence normalisation | Dev A | 0.5 day |
| 8 | Diagnostics & toggle | Dev C | 0.5 day |

Total approx. 5–6 engineering days with parallelisation.

---

## 6. Risk Mitigation

1. **Behaviour drift**: keep the legacy pipeline behind a flag. Run regression refresh on sample projects (small, medium, large) after each phase.
2. **XML schema changes**: ensure new helpers re-use existing `ClashZone` models; no schema change expected.
3. **Performance regressions**: track `perf.log` and `memory_profiling` snapshots before/after each phase.
4. **Logging noise**: maintain existing `[CLASH_DEBUG]` and `[FLAG-MANAGER]` breadcrumbs to confirm pipeline steps; once stable, can reduce verbosity.
5. **Incomplete normalisation**: add unit tests covering variations (`Plumbing`, `Plumbing.xml`, `Plumbing_pipes`, `plumbing.PIPES`). Confirm they all normalise to `Plumbing`.

---

## 7. Verification Checklist

*Benign cases*
- Replay (Adopt OFF) → no XML writes; `[CLASH_DEBUG] Replay path detected` logged.
- Replace mode → only flags updated; `Pipes_global.xml` merges into single branch.
- Adopt ON → detection runs once; branch remains normalized; placement uses new data.

*Performance metrics*
- Time to refresh with unchanged model < 20 s at 10 k zones.
- Memory usage < 100 MB; string pool reduces duplicates.
- Parameter count per zone ≤ 15.

*Data integrity*
- Global XML has only one `<Filter Name="Plumbing">`.
- `Plumbing.xml` remains master snapshot; `Plumbing_pipes.xml` (category) only updated when detection runs.
- Flags remain consistent (deleted sleeves set `IsResolved=false`, `SleeveInstanceId=-1`).

---

## 8. Migration Notes

1. Check in helper classes under `Services/Refresh/` (or continue using `refresh refactor/` until ready).
2. Update `.csproj` to include new classes once they’re ready; keep existing folder excluded until replacement pipeline is complete.
3. Document toggles in `OptimizationFlags` so QA can switch pipelines without rebuild.
4. Update architectural docs (`SLEEVE_PERSISTENCE_PLAN.md`, `SLEEVE_PLACEMENT_METHODOLOGY.md`) with new flow diagrams once refactor is merged.

---

## 9. Open Questions

- Do we need a migration script to prune legacy clusters/filters created by the old duplication? (Optional; low priority.)
- Should the new helpers live in a dedicated namespace (`JSE_RevitAddin_MEP_OPENINGS.Services.Refresh`) to simplify unit testing? (Recommended.)
- Can we switch to incremental detection (Phase 3 flags) after baseline is stable? (Future enhancement.)

---

## 10. Next Steps

1. Review this plan together and adjust sequencing if needed.
2. Start with Phase 1 tasks—low risk and high payoff.
3. Re-enable `refresh refactor/` folder in the project only after helper classes are finalised.
4. Keep logs from each phase for comparison (`perf.log`, `refresh_*.log`, `memory_profiling_*.log`).

Once these steps are complete we’ll have a modular, high-performance refresh pipeline that keeps flag management and placement data consistent across replay, replace, and detection modes.


