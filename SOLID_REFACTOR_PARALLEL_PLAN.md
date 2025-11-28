# SOLID Refactor Parallel Execution Plan

Date: 2025-11-25
Branch: `restore-today`
Scope: Architectural refactor to improve SOLID compliance without loss of existing features (fail-safe execution, DB integration, optimization layers: batching, R-Tree DB index, spatial grid, parallel clearance).

## Invariants To Preserve (MUST NOT REGRESS)
- Fail-safe: No unhandled exceptions abort global placement loop; continue-on-error semantics remain.
- Data integrity: DB (ClashZoneRepository / Sleeve data) updates remain atomic per sleeve.
- Optimization features: Deferred parameter batching, R-Tree DB index usage, spatial grid filtering, parallel clearance calculation.
- Logging contract: Existing log file names stay (placement_performance.log, parameter_batching_performance.log, placement_errors.log, service_instantiation.log).
- Flags behavior: Current OptimizationFlags values remain source of truth until replaced by injected config objects.
- Deployment mode: Disables performance logging & verbose diagnostics automatically.

## High-Level Target Architecture
```
PlacementCommand -> ISleevePlacementOrchestrator
  ISleevePlacementOrchestrator executes ordered IPlacementStage components with an immutable PlacementContext + builder-style updates
    PlacementContext (immutable inputs, correlation id) carries ClashZones, CalculatedDimensions, FamilySymbols, InstancesPlaced, Errors, Metrics via With* methods
  Stages:
    1 FilterStage (zone filtering)
    2 SmartReplayStage (attempt reuse of saved data)
    3 DimensionStage (IClearanceStrategy + geometry adapters)
    4 FamilySelectionStage (ISleeveFamilySelector)
    5 InstancePlacementStage (ISleeveInstancePlacer)
    6 ParameterAssignStage (uses IParameterBatchingService)
    7 FlushParametersStage (regenerate + flush)
    8 PersistenceStage (ISleevePersistenceService -> DB/XML)
    9 FlagUpdateStage (IFlagUpdateService)
    10 SummaryStage (IPerformanceSink / ILogger)
```
Each stage adheres to SRP and is replaceable (OCP). The orchestrator depends only on abstractions (DIP).

## Team Division (Parallel, Non-Overlapping)

### Team A – Orchestrator & Pipeline Core (You / Current Agent)
- Deliverables:
  - `ISleevePlacementOrchestrator` interface
  - `IPlacementStage` interface
  - `PlacementContext` (immutable inputs, builder-style With* methods, correlation id)
  - Basic sequential orchestration + error boundary (central try/catch wrapping each stage)
  - Result pattern for stages and orchestrator to model partial failures
- Constraints:
  - Do NOT refactor existing services yet; only wrap current flows.
  - Use existing `NewSleevePlacerService` logic as source to carve stage boundaries (no functional changes).
- Acceptance Criteria:
  - Orchestrator can run with stub stages delegating to existing methods.
  - Maintains placement output identical to pre-refactor for test model.
  - Publishes migration milestones to remove wrapped legacy calls (avoid permanent strangler)

### Team B – Configuration & Environment Abstraction
- Deliverables:
  - `IOptimizationConfig` (versioned, grouped immutable records: PlacementOptions, DetectionOptions, PerformanceOptions)
  - Adapter reading current `OptimizationFlags` into config at startup.
  - `IAppEnvironment` (DeploymentMode, Paths, LoggingPolicy).
  - Replace direct static flag reads in non-Team-A files via injected `IOptimizationConfig` (ONLY after Team A finalizes interfaces).
- Constraints:
  - Must not delete or alter `OptimizationFlags` yet; provide coexistence adapter.
  - No business logic moves.
- Acceptance Criteria:
  - All new services resolve via DI without changing existing call sites besides constructor injection extensions.
  - Config `Version` increments tracked and asserted in orchestrator startup logs

### Team C – Optimization Layer Isolation
- Deliverables:
  - `IClearanceStrategy` (already partly implied; ensure interface boundary formalized).
  - `IRTreeIndexProvider` abstraction wrapping current R-Tree DB index usage (build, query API).
  - `IZonePreFilterService` (spatial grid + R-tree composition) returning filtered zones.
  - `IParallelPlanningService` (encapsulates parallel clearance calculation & threshold logic).
  - Move algorithm decisions (e.g. count > 10 triggers parallel) into strategy objects.
- Constraints:
  - No changes to output ordering or filtering semantics.
  - Provide feature toggles through injected `IOptimizationConfig` only (not direct flags).
- Acceptance Criteria:
  - Unit-testable services with deterministic inputs/outputs.
  - Performance delta negligible (<2% variance) compared to legacy flow on benchmark dataset.

### Team D – Persistence & Error/Logging Framework
- Deliverables:
  - `ISleevePersistenceService` (encapsulates DB updates + XML save path). Async-first: `Task PersistPlacementAsync(PlacementContext)`, `Task UpdateInstanceAsync(ClashZone, FamilyInstance)`.
  - Transaction boundaries and compensation strategy (rollback/cleanup if persistence fails after flush).
  - `IErrorPolicy` (decides log level, continuation, aggregation). Implementation: `FailSafeErrorPolicy`.
  - `ILogger` abstraction (wraps DebugLogger + SafeFileLogger) with structured logging capability and correlation id.
  - Replace direct repository construction with persistence service (coexisting adapter initially).
- Constraints:
  - Must preserve existing log file names & message prefixes.
  - DB schemas untouched.
- Acceptance Criteria:
  - All persistence operations route through one abstraction.
  - Error handling no longer scattered; orchestrator receives standardized result objects.
  - DB operations transactional per batch; on failure, compensation handlers invoked

## Integration Sequencing (Avoid Conflicts)
| Phase | Team A | Team B | Team C | Team D |
|-------|--------|--------|--------|--------|
| 1     | Define core interfaces | Draft config models | Draft strategy interfaces | Draft error/logging interfaces |
| 2     | Implement orchestrator skeleton | Implement adapter from flags | Implement concrete optimization services | Implement persistence + logging adapters |
| 3     | Replace direct calls with stages | Begin replacing flag reads | Wire optimization services into stages | Swap DB calls with persistence service |
| 4     | Remove legacy conditional blocks gradually | Deprecate portions of `OptimizationFlags` | Performance regression test | Centralize error handling & finalize |

No team edits the same file in Phase 1–2 except orchestrator introduction. Legacy `NewSleevePlacerService` remains read-only for Teams B–D until Team A publishes stage boundaries.

## Proposed New Interfaces (Draft Signatures)
```csharp
public interface ISleevePlacementOrchestrator {
  OrchestratorResult Execute(IReadOnlyList<ClashZone> zones, IPerformanceMonitor perf);
}

public interface IPlacementStage {
  string Name { get; }
  StageResult Execute(PlacementContext context, IPerformanceMonitor perf);
}

public sealed class PlacementContext {
  public string CorrelationId { get; } = Guid.NewGuid().ToString();
  public Document Doc { get; }
  public IReadOnlyList<ClashZone> InputZones { get; }
  public IOptimizationConfig Config { get; }
  public IReadOnlyList<ClashZone> FilteredZones { get; }
  public IReadOnlyDictionary<int, (double width, double height, double diameter, bool circular)> Dimensions { get; }
  public IReadOnlyList<FamilyInstance> PlacedInstances { get; }
  public IReadOnlyList<string> Errors { get; }
  public IReadOnlyDictionary<string, object> Metrics { get; }
  public PlacementContext(Document doc, IReadOnlyList<ClashZone> zones, IOptimizationConfig config,
    IReadOnlyList<ClashZone> filtered = null,
    IReadOnlyDictionary<int,(double,double,double,bool)> dims = null,
    IReadOnlyList<FamilyInstance> placed = null,
    IReadOnlyList<string> errors = null,
    IReadOnlyDictionary<string,object> metrics = null) {
    Doc = doc; InputZones = zones; Config = config;
    FilteredZones = filtered ?? Array.Empty<ClashZone>();
    Dimensions = dims ?? new Dictionary<int,(double,double,double,bool)>();
    PlacedInstances = placed ?? Array.Empty<FamilyInstance>();
    Errors = errors ?? Array.Empty<string>();
    Metrics = metrics ?? new Dictionary<string, object>();
  }
  public PlacementContext WithFilteredZones(IReadOnlyList<ClashZone> filtered) =>
    new(Doc, InputZones, Config, filtered, Dimensions, PlacedInstances, Errors, Metrics);
  public PlacementContext WithDimensions(int zoneId, (double w,double h,double d,bool c) dim) {
    var dict = new Dictionary<int,(double,double,double,bool)>(Dimensions) { [zoneId] = dim };
    return new(Doc, InputZones, Config, FilteredZones, dict, PlacedInstances, Errors, Metrics);
  }
  public PlacementContext WithPlaced(FamilyInstance inst) {
    var list = PlacedInstances.ToList(); list.Add(inst);
    return new(Doc, InputZones, Config, FilteredZones, Dimensions, list, Errors, Metrics);
  }
  public PlacementContext WithError(string error) {
    var list = Errors.ToList(); list.Add(error);
    return new(Doc, InputZones, Config, FilteredZones, Dimensions, PlacedInstances, list, Metrics);
  }
  public PlacementContext WithMetric(string k, object v) {
    var dict = new Dictionary<string,object>(Metrics) { [k] = v };
    return new(Doc, InputZones, Config, FilteredZones, Dimensions, PlacedInstances, Errors, dict);
  }
}

public interface IOptimizationConfig {
  int Version { get; }
  PlacementOptions Placement { get; }
  DetectionOptions Detection { get; }
  PerformanceOptions Performance { get; }
}

public record PlacementOptions(bool UseParallelClearance, bool UseParameterBatching, bool UseSmartReplay);
public record DetectionOptions(bool UseRTreeIndex, bool UseSpatialGrid);
public record PerformanceOptions(bool EnablePerfLogging);

public sealed record StageResult(bool Success, PlacementContext Context, IReadOnlyList<string> Errors);
public sealed record OrchestratorResult(bool Success, int PlacedCount, int ErrorCount, string CorrelationId);

public interface IRTreeIndexProvider { void Build(IReadOnlyList<ClashZone> zones); IReadOnlyList<ClashZone> Query(BoundingBoxXYZ box); }
public interface IZonePreFilterService { IReadOnlyList<ClashZone> Filter(IReadOnlyList<ClashZone> zones); }
public interface IParallelPlanningService { void PlanDimensions(IReadOnlyList<ClashZone> zones, PlacementContext ctx); }
public interface ISleeveFamilySelector { FamilySymbol Select(ClashZone zone, bool circular); }
public interface ISleeveInstancePlacer {
  FamilyInstance Place(
    FamilySymbol symbol,
    XYZ placementPoint,
    Level level,
    double rotation,
    PlacementContext ctx);
}
public interface ISleevePersistenceService {
  Task PersistPlacementAsync(PlacementContext ctx);
  Task UpdateInstanceAsync(ClashZone zone, FamilyInstance instance);
}
public interface IErrorPolicy { void Handle(Exception ex, string scope, PlacementContext ctx, bool critical = false); }
public interface ILogger { void Info(string msg); void Warning(string msg); void Error(string msg); }
```

## Acceptance & Testing Strategy
### Unit Tests
- Each stage tested in isolation with synthetic inputs, verifying idempotence and immutability (no in-place mutation of context)
- Error policy tests: thrown exceptions translated into StageResult errors without propagation
- Config adapter tests for versioning

### Integration Tests
- Stage combinations pipelines (Filter->Dimension->Placement) produce identical outputs vs baseline
- Persistence mocked to validate transaction boundaries and compensation calls

### Performance Benchmarks
- Benchmark suites for batching on/off, parallel clearance thresholds, R-Tree on/off

### Acceptance Test Matrix (Key Scenarios)
| Scenario | Expected Outcome | Non-Regression Checks |
|----------|------------------|------------------------|
| Normal placement 500 zones | Same count placed | Time +/-5%, DB entries intact |
| Missing family | Logged warning, continue | No crash, error in placement_errors.log |
| Batching disabled | Direct parameter writes | Flush stage skipped safely |
| Deployment mode | No perf logs | Core errors still logged |
| R-Tree disabled | Spatial grid fallback works | Placement count equal |
| Parallel clearance toggle | Dim calc time changes | Correct dimensions identical |
| Smart Replay active | Reduced dimension recalcs | Context marks reused zones |
| Persistence failure after flush | Compensation executed, continue or abort per policy | No DB corruption; logs contain correlation id |
| Large dataset 1000+ zones | Batching/streaming used to limit memory | No OOM; throughput within target |

## Risk Mitigation
- Gate merges with diff-based placement validation (counts & dimensions vs baseline JSON snapshot).
- Preserve legacy service until orchestrator passes full regression matrix.
- Introduce feature flags for each new abstraction to allow rollback.

## Logging Standardization (Future Team D)
Format: `[UTC ISO][Stage][Level] Message | zoneId= | sleeveId=`
All stages use injected `ILogger`; file routing handled by implementation.

## Performance Hooks
All stages receive `IPerformanceMonitor` and report per-stage metrics (Start/Stop). Optional `IPerformanceSink` can export aggregated orchestrator results. Keep existing file outputs until all teams aligned.

## Deliverable Timeline (Suggested)
- Week 1: Interfaces + skeleton (Teams A–D Phase 1–2)
- Week 2: Integration & partial replacement (Team A leads coordination)
- Week 3–4: Optimization isolation + persistence abstraction + unit/integration tests
- Week 5: Full regression, performance tuning, bug fixes; deprecate legacy direct calls

## Coordination Protocol
- Daily sync: 10 min status (blocked interfaces, integration points).
- Interface change proposals: PR + RFC comment block required before alteration.
- Tag commits with `REF-A`, `REF-B`, `REF-C`, `REF-D` for traceability.

## Do Not Change (Until Final Phase)
- File names for existing log outputs.
- Public signatures of already consumed command entry points.
- Database schema and existing XML formats.

## Next Immediate Step (Team A)
Create interface files (empty bodies) in `Services/Interfaces/Refactor/` and an orchestrator stub in `Services/Placement/` (deferred until explicit approval). Include immutable `PlacementContext`, `StageResult`, and correlation id support.

---
Prepared by: Division A (Agent)
Status: Ready for parallel adoption
