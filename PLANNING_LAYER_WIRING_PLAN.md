# Parallel Planning Layer - Wiring Plan

## Overview
Integration plan for `ParallelSleevePlacementPlanner` into `UniversalSleevePlacerService` to enable:
- **Early skip detection** (before Revit API calls)
- **Reordering** (high-risk zones processed last)
- **Performance optimization** (parallel pre-computation)
- **Better diagnostics** (risk classification)

## Architecture Principles

### ✅ Safety First
- **No breaking changes** - Existing code continues to work
- **Feature flag** - Can be enabled/disabled via configuration
- **Graceful fallback** - If planner fails, fall back to existing logic
- **Backward compatible** - Works with existing ClashZone processing

### ✅ OOP Design
- **Dependency Injection** - Planner injected via constructor (optional)
- **Interface-based** - Uses `ISleevePlacementPlanner` abstraction
- **Single Responsibility** - Planner handles pure math, service handles Revit API
- **Open/Closed** - Extensible without modifying existing code

## Integration Points

### 1. Constructor Injection (Optional Dependency)

**File:** `Services/UniversalSleevePlacerService.cs`

**Location:** Constructor signature

```csharp
public UniversalSleevePlacerService(
    Document doc,
    OpeningConditions conditions,
    ISleevePlacementStrategy strategy,
    Dictionary<string, double> clearanceSettings,
    string filterName,
    bool isReplayPath,
    ISleevePlacementPlanner? planner = null)  // ✅ NEW: Optional planner injection
{
    // ... existing constructor code ...
    
    // ✅ NEW: Store planner (or create default if not provided)
    _planner = planner ?? new ParallelSleevePlacementPlanner();
}
```

**Benefits:**
- Testable (can inject mock planner)
- Optional (backward compatible)
- Configurable (can use different planner implementations)

---

### 2. Planning Phase (Early in Placement)

**File:** `Services/UniversalSleevePlacerService.cs`

**Location:** `PlaceAllSleevesInTransaction` method, **BEFORE** main placement loop

**Insert after:** Line ~520 (after `placementLoopTracker` initialization, before `foreach` loop)

```csharp
// ✅ PERFORMANCE: Track main placement loop
using (var placementLoopTracker = performanceMonitor.TrackOperation("Place Individual Sleeves Loop"))
{
    // ✅ NEW: PLANNING PHASE - Pre-compute all sleeve data in parallel
    List<ClashZone> zonesToProcess;
    Dictionary<Guid, SleevePlacementPlanningDto> planningMap = null;
    SleevePlacementPlanningResult planningResult = null;
    
    using (var planningTracker = performanceMonitor?.TrackOperation("Sleeve Placement Planning"))
    {
        try
        {
            // Run parallel planning
            planningResult = _planner.Plan(clashZones);
            planningTracker?.SetItemCount(planningResult.TotalCount);
            
            // Create lookup map: ClashZoneId -> PlanningDto
            planningMap = planningResult.Items
                .ToDictionary(dto => dto.ClashZoneId, dto => dto);
            
            // ✅ EARLY SKIP: Filter out zones marked for skip
            var skipGuids = planningResult.Items
                .Where(dto => dto.ShouldSkip)
                .Select(dto => dto.ClashZoneId)
                .ToHashSet();
            
            zonesToProcess = sortedClashZones
                .Where(cz => !skipGuids.Contains(cz.Id))
                .ToList();
            
            // ✅ REORDERING: Sort by risk (low risk first, high risk last)
            // This ensures problematic zones are processed last (less likely to block good zones)
            zonesToProcess = zonesToProcess
                .OrderBy(cz => planningMap.TryGetValue(cz.Id, out var dto) 
                    ? (int)dto.ClearanceRisk 
                    : int.MaxValue)
                .ThenByDescending(cz => planningMap.TryGetValue(cz.Id, out var dto) 
                    ? dto.RawMepSizeFt 
                    : 0.0)
                .ToList();
            
            // ✅ DIAGNOSTIC: Log planning metrics
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[PLANNING] ✅ Planning completed: " +
                    $"Total={planningResult.TotalCount}, " +
                    $"Skipped={planningResult.SkippedCount}, " +
                    $"HighRisk={planningResult.HighRiskCount}, " +
                    $"CriticalRisk={planningResult.CriticalRiskCount}, " +
                    $"Duration={planningResult.PlanningDurationMs:F1}ms");
                
                SafeFileLogger.SafeAppendText("planning_debug.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] === PLANNING PHASE ===\n" +
                    $"Total Zones: {planningResult.TotalCount}\n" +
                    $"Skipped: {planningResult.SkippedCount}\n" +
                    $"High Risk: {planningResult.HighRiskCount}\n" +
                    $"Critical Risk: {planningResult.CriticalRiskCount}\n" +
                    $"Planning Duration: {planningResult.PlanningDurationMs:F1}ms\n" +
                    $"Zones to Process: {zonesToProcess.Count}\n\n");
            }
        }
        catch (Exception planningEx)
        {
            // ✅ GRACEFUL FALLBACK: If planning fails, use original sorted list
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Warning($"[PLANNING] ⚠️ Planning phase failed, falling back to original logic: {planningEx.Message}");
                SafeFileLogger.SafeAppendText("planning_errors.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] Planning error: {planningEx.Message}\n{planningEx.StackTrace}\n");
            }
            
            zonesToProcess = sortedClashZones; // Fallback to original list
            planningMap = null; // No planning data available
        }
    } // End planning phase tracking
    
    // ✅ MAIN PLACEMENT LOOP: Use reordered and filtered zones
    foreach (var clashZone in zonesToProcess)  // ✅ CHANGED: Use zonesToProcess instead of sortedClashZones
    {
        // ... existing placement loop code ...
    }
}
```

**Benefits:**
- **Early skip detection** - Zones marked for skip are filtered before Revit API calls
- **Reordering** - Low-risk zones processed first
- **Performance** - Parallel pre-computation saves time
- **Diagnostics** - Risk classification available for logging

---

### 3. DTO Data Usage in Placement Loop

**File:** `Services/UniversalSleevePlacerService.cs`

**Location:** Inside placement loop, where clearance/sizing is calculated

**Insert after:** Line ~600 (after validation, before clearance calculation)

```csharp
// ✅ NEW: Use planning DTO data if available (pre-computed values)
SleevePlacementPlanningDto? dto = null;
if (planningMap != null && planningMap.TryGetValue(clashZone.Id, out var planningDto))
{
    dto = planningDto;
    
    // ✅ DIAGNOSTIC: Log that we're using pre-computed values
    if (!DeploymentConfiguration.DeploymentMode)
    {
        DebugLogger.Info($"[PLANNING] Using pre-computed values for Zone={clashZone.Id}: " +
            $"TargetWidth={dto.TargetWidthFt:F3}ft, " +
            $"TargetHeight={dto.TargetHeightFt:F3}ft, " +
            $"Clearance={dto.ClearanceFt:F3}ft, " +
            $"Risk={dto.ClearanceRisk}");
    }
}

// ✅ CLEARANCE CALCULATION: Use DTO if available, otherwise calculate
double clearance;
if (dto != null)
{
    // Use pre-computed clearance from planning phase
    clearance = RevitUnitConversionService.Instance.ToInternalUnits(dto.ClearanceFt, UnitTypeId.Feet);
}
else
{
    // ✅ FALLBACK: Calculate clearance using existing logic
    clearance = CalculateClearance(clashZone, mepSize);
}
```

**Benefits:**
- **Performance** - Reuse pre-computed values
- **Consistency** - Same values used for planning and placement
- **Fallback** - Works even if planning data unavailable

---

### 4. Sleeve Size Selection (Use DTO Target Dimensions)

**File:** `Services/UniversalSleevePlacerService.cs`

**Location:** Where sleeve dimensions are determined (around family selection)

**Insert after:** Family selection logic

```csharp
// ✅ NEW: Use DTO target dimensions if available
double targetWidth = 0.0;
double targetHeight = 0.0;
if (dto != null)
{
    targetWidth = RevitUnitConversionService.Instance.ToInternalUnits(dto.TargetWidthFt, UnitTypeId.Feet);
    targetHeight = RevitUnitConversionService.Instance.ToInternalUnits(dto.TargetHeightFt, UnitTypeId.Feet);
}

// ✅ FALLBACK: Calculate dimensions if DTO not available
if (targetWidth <= 0 || targetHeight <= 0)
{
    // Use existing dimension calculation logic
    targetWidth = CalculateSleeveWidth(clashZone, mepSize);
    targetHeight = CalculateSleeveHeight(clashZone, mepSize);
}
```

---

### 5. Rotation Angle (Use DTO Rotation)

**File:** `Services/UniversalSleevePlacerService.cs`

**Location:** Where rotation angle is determined (for floor-hosted sleeves)

**Insert after:** Rotation angle calculation

```csharp
// ✅ NEW: Use DTO rotation angle if available
double rotationAngle = 0.0;
if (dto != null && dto.RotationAngleDeg != 0.0)
{
    rotationAngle = dto.RotationAngleDeg * Math.PI / 180.0; // Convert degrees to radians
}
else
{
    // ✅ FALLBACK: Calculate rotation using existing logic
    rotationAngle = CalculateRotationAngle(clashZone);
}
```

---

### 6. Batch Logging (Use DTO LogSummary)

**File:** `Services/UniversalSleevePlacerService.cs`

**Location:** After placement loop completes

**Insert after:** Line ~2568 (before return statement)

```csharp
// ✅ NEW: Batch write planning logs (if planning was used)
if (planningResult != null && !DeploymentConfiguration.DeploymentMode)
{
    try
    {
        var planningLogPath = SafeFileLogger.GetLogFilePath("planning_debug.log");
        var batchLog = new System.Text.StringBuilder();
        batchLog.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] === PLANNING DETAILS ===");
        
        foreach (var dto in planningResult.Items)
        {
            batchLog.AppendLine(dto.LogSummary);
        }
        
        batchLog.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] === END PLANNING DETAILS ===\n");
        File.AppendAllText(planningLogPath, batchLog.ToString());
    }
    catch (Exception logEx)
    {
        // Don't fail placement if logging fails
        if (!DeploymentConfiguration.DeploymentMode)
        {
            DebugLogger.Warning($"[PLANNING] Failed to write planning logs: {logEx.Message}");
        }
    }
}
```

---

### 7. Performance Metrics Integration

**File:** `Services/UniversalSleevePlacerService.cs`

**Location:** Performance report generation

**Update:** Add planning metrics to performance report

```csharp
// ✅ NEW: Include planning metrics in performance report
if (planningResult != null)
{
    performanceMonitor?.RecordPlanningMetrics(
        planningResult.TotalCount,
        planningResult.SkippedCount,
        planningResult.HighRiskCount,
        planningResult.CriticalRiskCount,
        planningResult.PlanningDurationMs);
}
```

---

## Configuration (Feature Flag)

**File:** `DeploymentConfiguration.cs` (or create new config)

```csharp
/// <summary>
/// Enable parallel planning layer for sleeve placement.
/// When enabled, pre-computes sleeve dimensions, clearance, and risk classification in parallel.
/// </summary>
public static bool EnableParallelPlanning { get; set; } = true; // Default: enabled
```

**Usage in UniversalSleevePlacerService:**

```csharp
// ✅ FEATURE FLAG: Only use planner if enabled
if (DeploymentConfiguration.EnableParallelPlanning && _planner != null)
{
    // Run planning phase
}
else
{
    // Use original logic (no planning)
    zonesToProcess = sortedClashZones;
    planningMap = null;
}
```

---

## Testing Strategy

### Unit Tests
1. **Planner Tests** - Test `ParallelSleevePlacementPlanner` in isolation
2. **Integration Tests** - Test planner + service integration
3. **Fallback Tests** - Test graceful degradation when planner fails

### Integration Tests
1. **Small batch** (< 12 zones) - Should use sequential processing
2. **Large batch** (> 12 zones) - Should use parallel processing
3. **Mixed risk** - Verify reordering works correctly
4. **Skip detection** - Verify zones are skipped early

### Performance Tests
1. **Before/After** - Compare placement times with/without planning
2. **Scalability** - Test with 100, 500, 1000+ zones
3. **Memory** - Verify no memory leaks in parallel processing

---

## Migration Path

### Phase 1: Integration (Current)
- ✅ Add planner injection to constructor
- ✅ Add planning phase before placement loop
- ✅ Use DTO data where available
- ✅ Maintain fallback to existing logic
- ✅ Add feature flag

### Phase 2: Optimization (Future)
- ✅ Replace all clearance calculations with DTO values
- ✅ Replace all dimension calculations with DTO values
- ✅ Remove redundant calculations in placement loop

### Phase 3: Advanced Features (Future)
- ✅ Risk-based retry logic (retry high-risk zones with different parameters)
- ✅ Predictive skip (skip zones likely to fail based on risk)
- ✅ Adaptive planning (adjust parameters based on success rate)

---

## Error Handling

### Planner Failure
- **Catch exception** in planning phase
- **Log error** to `planning_errors.log`
- **Fall back** to original `sortedClashZones` list
- **Continue** with existing placement logic

### DTO Missing
- **Check** if `planningMap` contains ClashZoneId
- **Fall back** to existing calculation logic
- **Log warning** if DTO expected but missing

### Data Mismatch
- **Validate** DTO data before use (check for NaN, Infinity, negative values)
- **Fall back** to existing calculation if invalid
- **Log warning** for data quality issues

---

## Performance Considerations

### Memory
- **DTO objects** are lightweight (immutable, no references)
- **Planning map** is `Dictionary<Guid, DTO>` - O(1) lookup
- **ConcurrentBag** used in parallel processing (thread-safe)

### CPU
- **Parallel processing** only for batches >= 12 zones
- **Sequential processing** for small batches (avoid overhead)
- **Configurable parallelism** via `maxDegreeOfParallelism`

### I/O
- **Batch logging** - Write all planning logs at once (not per-zone)
- **Async logging** - Consider async file writes for large batches

---

## Code Changes Summary

### Files to Modify
1. **`Services/UniversalSleevePlacerService.cs`**
   - Add `_planner` field
   - Update constructor signature
   - Add planning phase before placement loop
   - Use DTO data in placement loop
   - Add batch logging

### Files NOT Modified
- ✅ All existing placement logic remains unchanged
- ✅ All existing calculation methods remain unchanged
- ✅ All existing error handling remains unchanged

---

## Rollback Plan

If issues arise:
1. **Disable feature flag** - `DeploymentConfiguration.EnableParallelPlanning = false`
2. **Remove planner injection** - Set `_planner = null` in constructor
3. **Revert to original loop** - Use `sortedClashZones` directly

---

## Success Metrics

### Performance
- **Planning time** < 5% of total placement time
- **Placement time** reduced by 10-20% (due to early skips and reordering)
- **Memory usage** < 5MB additional overhead

### Quality
- **Skip accuracy** - 95%+ of skipped zones correctly identified
- **Risk classification** - 90%+ accuracy in risk assessment
- **Zero regressions** - All existing tests pass

---

## Next Steps

1. **Review this plan** with team
2. **Create feature branch** - `feature/parallel-planning-integration`
3. **Implement Phase 1** - Basic integration with feature flag
4. **Test thoroughly** - Unit tests, integration tests, performance tests
5. **Deploy to staging** - Test with real projects
6. **Monitor metrics** - Track performance improvements
7. **Iterate** - Refine based on feedback

---

## Questions to Resolve

1. **Feature flag location** - Where should `EnableParallelPlanning` be defined?
2. **Logging strategy** - Should planning logs be separate or merged with placement logs?
3. **Performance monitoring** - Should planning metrics be in same report or separate?
4. **Default behavior** - Should planning be enabled by default or opt-in?

---

## Appendix: Code Snippets

### Complete Planning Phase Integration

See full code in `PLANNING_INTEGRATION_SNIPPETS.md` (to be created)

### Unit Test Examples

See `PLANNING_UNIT_TESTS.md` (to be created)

### Performance Benchmarks

See `PLANNING_PERFORMANCE_BENCHMARKS.md` (to be created)

