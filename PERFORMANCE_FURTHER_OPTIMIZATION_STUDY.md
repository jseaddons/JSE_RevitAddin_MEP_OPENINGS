# Further Optimization & Implementation Study (2025-11-24)

## 1. Current Performance Snapshot (Logs Provided)

| Phase | Metric | Observed | Target | Gap |
|-------|--------|---------|--------|-----|
| Individual Sleeve Placement | Total 3 sleeves / 1087ms | 3 / 1.087s ≈ 2.76 sleeves/s | 50+ sleeves/s | 94% below |
| First Sleeve Instance Creation | 66ms create + 36ms params = 102ms core | Goal < 25ms | High cold-start penalty |
| Subsequent Sleeve Instance Creation | 16–17ms create + 6–8ms params ≈ 24ms | Goal < 20ms | Slightly above |
| Cluster Sleeve Placement | 1 cluster / 399ms (371ms placing) | 10+ clusters/s | 90% below |
| Refresh Intersection Processing | 983ms for 6 zones (≈164ms/zone) | <2ms/zone (500+/s) | 98% below |
| Flag Reset (Refresh) | 444ms | Goal < 50ms | Large batch DB overhead |
| Parameter Setting (Width/Height/Depth) | ~0ms (ticks < 2500) | <5ms | Achieved |
| Memory Delta (Refresh) | -21.4MB | Controlled | OK |

### Key Bottlenecks
1. Intersection processing (≈164ms/zone) – dominating refresh time.
2. Cluster sleeve creation (371ms for a single cluster sleeve).
3. First individual sleeve cold-start (≈100ms vs ~24ms warmed).
4. Flag reset & DB updates (444ms) – likely inefficient bulk updates or excessive per-row operations.
5. Disabled curve-in-bounding-box fast rejection path (commented out) increasing solid intersection workload.
6. Section box filtering in `SectionBoxHelper` still using solid intersection filter (`ElementIntersectsSolidFilter`) instead of cheaper outline intersection (`BoundingBoxIntersectsFilter`).

## 2. Achievements Confirmed
- Deferred metadata write optimization: Parameter timing shows near-zero time for geometry-driven params.
- Batch placement path reduces regenerations (single regeneration logged).
- Geometry & transform caching present (LRU + transform cache).
- Spatial grid and category whitelists active.
- Non-critical metadata deferred and batched.

## 3. Priority Optimization Opportunities
| # | Area | Action | Expected Gain | Effort | Risk | Files Impacted |
|---|------|--------|---------------|--------|------|----------------|
| 1 | Section Box Filtering | Replace solid filter with bounding box intersects (host + linked) | 3–5× fewer candidates | Low | Low | `Helpers/SectionBoxHelper.cs` |
| 2 | Fast Curve-In-BBox Stage | Re-enable `TestCurveInBoundingBox` after reliability validation | 4–8× fewer solid ops | Low | Low-Med (false negatives risk) | `Services/MepIntersectionService.cs` |
| 3 | Intersection Staging LOD | Implement progressive LOD pipeline (LOD0 outline → LOD1 curve → LOD2 solid) | User-perceived speed (sub-second preview) | Med | Low (behind flag) | `MepIntersectionService.cs`, `OptimizationFlags.cs` |
| 4 | Geometry Cache Structure | Store `List<Solid>` per element (compound walls) instead of single solid | Avoid repeated geometry extraction | Med | Low | `MepIntersectionService.cs` |
| 5 | Intersection Metrics Instrumentation | Log candidate counts per stage (grid → bbox → curve → solid) | Identifies poor filter selectivity | Low | Low | `MepIntersectionService.cs` |
| 6 | First Sleeve Cold Start | Pre-warm: load & activate all required symbols + force lightweight regeneration | Reduce first-instance >50% | Low | Low | `UniversalSleevePlacerService.cs`, `FamilyLoadingService.cs` |
| 7 | Cluster Sleeve Creation Algorithm | Optimize rotated bounding box & angle calc (vector math, remove redundant transforms) | 30–50% faster cluster placement | Med | Low-Med | `OpeningCommandOrchestrator.cs`, `UniversalClusterService.*` |
| 8 | Flag Reset Bulk Efficiency | Replace current loop with single SQL UPDATE using IN clause or temp table | 5–10× speed | Low | Low | `Data/Repositories/ClashZoneRepository.cs` |
| 9 | Parallel Cheap Intersection Pass | Parallelize bbox + curve filtering only (read-only math) | 2–3× faster refresh intersection | Med | Med (thread gating) | `MepIntersectionService.cs` |
|10 | Incremental Detection | Enable & harden geometry fingerprint reuse (skip unchanged pairs) | 10×+ refresh speed on repeated runs | High | High | `IncrementalClashTracker (new)`, `MepIntersectionService.cs` |

## 4. Detailed Implementation Guidance

### 4.1 Replace Solid Section Box Filtering
**Current Issue:** `SectionBoxHelper` uses `ElementIntersectsSolidFilter` (higher cost). 
**Change:** Compute oriented box → axis-aligned world outline → apply `BoundingBoxIntersectsFilter` for host and transformed outline for links.
**Patch Sketch:**
```csharp
var sectionBox = view3D.GetSectionBox();
var t = sectionBox.Transform;
var worldMin = t.OfPoint(sectionBox.Min);
var worldMax = t.OfPoint(sectionBox.Max);
var outline = new Outline(worldMin, worldMax);
var bbFilter = new BoundingBoxIntersectsFilter(outline);
var passingHostIds = new FilteredElementCollector(uiDoc.Document, hostIds)
    .WherePasses(bbFilter)
    .ToElementIds();
// For link: transform outline by inverse link transform
```
**Flags:** None required. Safe default.

### 4.2 Re-Enable Curve-In-Bounding-Box Filter
**Current:** Commented out block around lines ~740–820. 
**Action:** Restore and instrument: log rejection stats: `curveRejected`, `solidTestsPerformed`.
**Fallback:** If rejection rate <30% over N runs, auto-disable via adaptive flag.
**Add Flag:** `OptimizationFlags.UseCurveInBoundingBoxPreFilter` (default true).

### 4.3 Progressive LOD Pipeline
**LOD0:** Cheap outline + grid; deliver preliminary clash *count* & list of candidate IDs.
**LOD1:** Curve-in-bbox filter; compute approximate intersection centers.
**LOD2:** Full solid intersection (existing path).
Add method `DetectClashesProgressive(mode)` where `mode` requested from UI. 
**Flags:** `UseProgressiveLOD`. 
**API Surface:** Keep current `FindIntersectionsBatch` for LOD2; add wrappers for LOD0/LOD1.

### 4.4 Geometry Cache – Multi-Solid Storage
Change cache value from `Solid?` to `List<Solid>`; rename helpers:
```csharp
private static Lazy<Dictionary<string, List<Solid>>> _geometryCache;
private static bool TryGetFromGeometryCache(string key, out List<Solid> solids);
```
Migration: Accessors wrap legacy until all references adjusted. Benefit: compound walls avoid repeat geometry traversal per layer.

### 4.5 Intersection Metrics Instrumentation
Add struct:
```csharp
record IntersectionStageMetrics(int GridCandidates, int BBoxCandidates, int CurveCandidates, int SolidTests, int Hits);
```
Log to `intersection_metrics.log`. Use rolling window averages for adaptively tuning tolerance.

### 4.6 Sleeve Creation Cold-Start Optimization
Steps:
1. Preload all required families earlier in workflow (already partially done) and call `Activate()` once.
2. Dummy micro-instance create & delete inside transaction to force symbol caching.
3. Convert first-sleeve path to reuse pre-resolved level & symbol references.
Add instrumentation for `FirstSleeveWarmup`. Expect reduction from 66ms → <30ms.

### 4.7 Cluster Placement Optimizations
- Replace rotated bounding box recomputation with direct basis vector projection.
- Cache angle result per host/system key.
- Defer cluster sleeve parameter writes (if any non-critical) into batch.
- Avoid sleep/regenerate sequences unless necessary – measure bounding box availability pre-regenerate.

### 4.8 Bulk Flag Reset Optimization
Current: Iterative flag updates. New:
```sql
UPDATE ClashZones SET ReadyForPlacementFlag = 0 WHERE ZoneGuid IN (...);
```
Use parameterized command with chunking (e.g., 500 GUIDs). Add single transaction scope.

### 4.9 Parallel Cheap Pass
Pattern:
```csharp
Parallel.ForEach(mepElements, opts, mep => {
  // Only bounding box + grid + curve test, collect candidates
});
// Sequential: solid intersections (Revit API geometry)
```
Ensure no Revit API geometry extraction inside parallel region—only cached bboxes & math. Gate by `OptimizationFlags.UseParallelCheapPass`.

### 4.10 Incremental Detection
Add `ClashFingerprint` (existing plan) storing geometry hash + candidate structural IDs.
On refresh:
- Hash unchanged → skip full path, trust prior results unless any dependent structural hash changed.
- Track invalidation triggers: element changed, link reloaded, section box modified.
Persist fingerprints in lightweight local DB table or memory if session-only.

## 5. Suggested New Feature Flags
```csharp
public static bool UseCurveInBoundingBoxPreFilter { get; set; } = true;
public static bool UseProgressiveLOD { get; set; } = false; // experimental
public static bool UseParallelCheapPass { get; set; } = false; // guarded
```

## 6. Measurement & Validation Plan
| Metric | Tool/Log | Success Threshold |
|--------|----------|-------------------|
| Curve pre-filter rejection % | intersection_metrics.log | >60% rejection |
| Solid intersection tests per clash | metrics log | <1.4 average |
| First sleeve creation time | placement_performance.log | <30ms create, <10ms params |
| Cluster sleeve place time | cluster_performance.log | <120ms |
| Intersection processing per zone | refresh performance log | <8ms |
| Flag reset duration | refresh performance log | <50ms |
| Memory per zone | refresh performance log | <5KB (post optimization) |

## 7. Rollout Strategy
1. Implement SectionBoxHelper change (safe). 
2. Introduce curve-in-bbox prefilter + metrics (flag ON). 
3. Add metrics logging scaffold & geometry cache upgrade. 
4. Warmup sleeve creation improvements. 
5. Bulk flag reset SQL optimization. 
6. Enable adaptive tuning based on metrics (auto-switch strategies). 
7. Progressive LOD & incremental detection behind flags (off until validated). 

## 8. Risk Mitigation
| Risk | Mitigation |
|------|------------|
| False negatives with curve pre-filter | Add fallback solid test sampling; log misses; disable automatically if mismatch detected. |
| Incorrect transformed section box bounds | Unit tests for outline transform vs element bounding box intersection. |
| Multi-solid cache memory growth | Track cache size; LRU eviction unchanged. |
| Parallel race conditions | Restrict parallel block to pure math (no `get_Geometry`). |
| Incremental detection stale results | Force invalidation on global events (DocumentChanged, link reload, section box change). |

## 9. Immediate Action Patches (Summary)
1. `SectionBoxHelper.cs`: Replace solid filter with outline filter logic. 
2. `OptimizationFlags.cs`: Add three new flags. 
3. `MepIntersectionService.cs`: Re-enable curve pre-filter + metrics recording scaffold. 
4. `ClashZoneRepository.cs`: Add bulk reset method using single SQL command. 
5. `UniversalSleevePlacerService.cs`: Add pre-warm routine before first placement.

## 10. Decomposition for Next Sprint
| Day | Tasks |
|-----|-------|
| 1 | SectionBoxHelper patch + logging & curve pre-filter re-enable |
| 2 | Metrics scaffolding + geometry cache List<Solid> migration |
| 3 | Sleeve warmup & flag reset SQL batch |
| 4 | Progressive LOD initial API + UI flag wiring |
| 5 | Incremental detection prototype + validation harness |

## 11. Open Questions
- Do we need UI controls for LOD selection or auto-advance? 
- Should intersection metrics persist across sessions for adaptive tuning? 
- Acceptable cluster sleeve timing target relative to individual sleeves? 

## 12. Appendix – Parameter Timing Validation
Sample param timing ticks (<2500) confirm write cost negligible post deferral. Continue sampling but can disable instrumentation in production (`EnableParameterTimingInstrumentation=false`).

---
**Conclusion:** Primary speed blockers are intersection stage selectivity (curve filter off), cluster placement algorithm inefficiencies, and per-refresh intersection complexity. Implementing outlined low-risk changes (bbox filter fix + curve pre-filter + bulk flag reset) should yield 5–10× immediate improvement in refresh and cluster operations, paving way for advanced LOD and incremental detection to meet aggressive targets.
