# Planning Layer Integration - Step-by-Step Implementation

## Prerequisites
- ✅ Planning components already created (confirmed)
- ✅ No existing files modified (safety maintained)
- ✅ Ready for integration

---

## Step 1: Add Planner Field and Constructor Parameter

**File:** `Services/UniversalSleevePlacerService.cs`

### 1.1 Add Field Declaration

**Location:** After line ~40 (with other private fields)

```csharp
// ✅ NEW: Optional planner for parallel pre-computation
private readonly ISleevePlacementPlanner? _planner;
```

### 1.2 Update Constructor

**Location:** Line ~49 (constructor signature)

**Change from:**
```csharp
public UniversalSleevePlacerService(
    Document doc,
    OpeningConditions conditions,
    ISleevePlacementStrategy strategy,
    Dictionary<string, double> clearanceSettings = null,
    string filterName = null,
    FlagManager flagManager = null,
    bool isReplayPath = false)
```

**Change to:**
```csharp
public UniversalSleevePlacerService(
    Document doc,
    OpeningConditions conditions,
    ISleevePlacementStrategy strategy,
    Dictionary<string, double> clearanceSettings = null,
    string filterName = null,
    FlagManager flagManager = null,
    bool isReplayPath = false,
    ISleevePlacementPlanner? planner = null)  // ✅ NEW: Optional planner injection
```

### 1.3 Initialize Planner in Constructor Body

**Location:** After line ~66 (after `_flagManager` initialization)

**Add:**
```csharp
// ✅ NEW: Initialize planner (create default if not provided)
_planner = planner ?? new ParallelSleevePlacementPlanner();
```

---

## Step 2: Add Planning Phase Before Placement Loop

**File:** `Services/UniversalSleevePlacerService.cs`

**Location:** Line ~520 (after `placementLoopTracker` initialization, before `foreach` loop)

### 2.1 Insert Planning Phase Code

**Replace:**
```csharp
// ✅ PERFORMANCE: Track main placement loop
using (var placementLoopTracker = performanceMonitor.TrackOperation("Place Individual Sleeves Loop"))
{
    foreach (var clashZone in sortedClashZones)
    {
```

**With:**
```csharp
// ✅ PERFORMANCE: Track main placement loop
using (var placementLoopTracker = performanceMonitor.TrackOperation("Place Individual Sleeves Loop"))
{
    // ✅ NEW: PLANNING PHASE - Pre-compute all sleeve data in parallel
    List<ClashZone> zonesToProcess;
    Dictionary<Guid, SleevePlacementPlanningDto>? planningMap = null;
    SleevePlacementPlanningResult? planningResult = null;
    
    // ✅ FEATURE FLAG: Only use planner if enabled
    bool usePlanning = DeploymentConfiguration.EnableParallelPlanning && _planner != null;
    
    if (usePlanning)
    {
        using (var planningTracker = performanceMonitor?.TrackOperation("Sleeve Placement Planning"))
        {
            try
            {
                // Run parallel planning
                planningResult = _planner!.Plan(clashZones);
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
    }
    else
    {
        // ✅ FALLBACK: Use original sorted list if planning disabled
        zonesToProcess = sortedClashZones;
        planningMap = null;
    }
    
    // ✅ MAIN PLACEMENT LOOP: Use reordered and filtered zones
    foreach (var clashZone in zonesToProcess)  // ✅ CHANGED: Use zonesToProcess instead of sortedClashZones
    {
```

---

## Step 3: Use DTO Data in Placement Loop

**File:** `Services/UniversalSleevePlacerService.cs`

**Location:** Inside placement loop, after validation (around line ~600)

### 3.1 Get Planning DTO for Current Zone

**Insert after:** Validation code (around line ~600)

```csharp
// ✅ NEW: Get planning DTO for current zone (if available)
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
```

### 3.2 Use DTO for Clearance Calculation

**Location:** Where clearance is calculated (find `CalculateClearance` call)

**Replace:**
```csharp
double clearance = CalculateClearance(clashZone, mepSize);
```

**With:**
```csharp
// ✅ CLEARANCE CALCULATION: Use DTO if available, otherwise calculate
double clearance;
if (dto != null)
{
    // Use pre-computed clearance from planning phase
    clearance = RevitUnitConversionService.Instance.ToInternalUnits(dto.ClearanceFt, UnitTypeId.Feet);
    
    if (!DeploymentConfiguration.DeploymentMode)
    {
        DebugLogger.Info($"[PLANNING] Using pre-computed clearance: {dto.ClearanceFt:F3}ft = {clearance:F6} internal units");
    }
}
else
{
    // ✅ FALLBACK: Calculate clearance using existing logic
    clearance = CalculateClearance(clashZone, mepSize);
}
```

### 3.3 Use DTO for Sleeve Dimensions

**Location:** Where sleeve width/height are calculated (find dimension calculation)

**Insert:**
```csharp
// ✅ NEW: Use DTO target dimensions if available
double targetWidth = 0.0;
double targetHeight = 0.0;
if (dto != null)
{
    targetWidth = RevitUnitConversionService.Instance.ToInternalUnits(dto.TargetWidthFt, UnitTypeId.Feet);
    targetHeight = RevitUnitConversionService.Instance.ToInternalUnits(dto.TargetHeightFt, UnitTypeId.Feet);
    
    if (!DeploymentConfiguration.DeploymentMode)
    {
        DebugLogger.Info($"[PLANNING] Using pre-computed dimensions: " +
            $"Width={dto.TargetWidthFt:F3}ft = {targetWidth:F6} internal, " +
            $"Height={dto.TargetHeightFt:F3}ft = {targetHeight:F6} internal");
    }
}

// ✅ FALLBACK: Calculate dimensions if DTO not available
if (targetWidth <= 0 || targetHeight <= 0)
{
    // Use existing dimension calculation logic
    targetWidth = CalculateSleeveWidth(clashZone, mepSize);
    targetHeight = CalculateSleeveHeight(clashZone, mepSize);
}
```

### 3.4 Use DTO for Rotation Angle

**Location:** Where rotation angle is calculated (for floor-hosted sleeves)

**Insert:**
```csharp
// ✅ NEW: Use DTO rotation angle if available
double rotationAngle = 0.0;
if (dto != null && Math.Abs(dto.RotationAngleDeg) > 0.001)
{
    rotationAngle = dto.RotationAngleDeg * Math.PI / 180.0; // Convert degrees to radians
    
    if (!DeploymentConfiguration.DeploymentMode)
    {
        DebugLogger.Info($"[PLANNING] Using pre-computed rotation: {dto.RotationAngleDeg:F1}° = {rotationAngle:F6} radians");
    }
}
else
{
    // ✅ FALLBACK: Calculate rotation using existing logic
    rotationAngle = CalculateRotationAngle(clashZone);
}
```

---

## Step 4: Add Feature Flag Configuration

**File:** `DeploymentConfiguration.cs` (or wherever configuration is stored)

**Add:**
```csharp
/// <summary>
/// Enable parallel planning layer for sleeve placement.
/// When enabled, pre-computes sleeve dimensions, clearance, and risk classification in parallel.
/// Default: true (enabled)
/// </summary>
public static bool EnableParallelPlanning { get; set; } = true;
```

---

## Step 5: Add Batch Logging After Placement

**File:** `Services/UniversalSleevePlacerService.cs`

**Location:** After placement loop completes (around line ~2568, before return statement)

**Insert:**
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
        
        DebugLogger.Info($"[PLANNING] ✅ Batch logged {planningResult.Items.Count} planning entries to {planningLogPath}");
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

## Step 6: Add Required Using Statements

**File:** `Services/UniversalSleevePlacerService.cs`

**Location:** Top of file (with other using statements)

**Add:**
```csharp
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;  // ✅ NEW: For ISleevePlacementPlanner and DTOs
```

---

## Step 7: Update Call Sites (Optional)

If you want to inject a custom planner, update call sites:

**File:** `SleevePlacementCoordinator.cs` (or wherever service is instantiated)

**Example:**
```csharp
// ✅ NEW: Create planner instance (can be shared/reused)
var planner = new ParallelSleevePlacementPlanner(
    minParallelCount: 12,
    maxDegree: Environment.ProcessorCount);

// ✅ NEW: Pass planner to service
var service = new UniversalSleevePlacerService(
    doc,
    conditions,
    strategy,
    clearanceSettings,
    filterName,
    flagManager,
    isReplayPath,
    planner);  // ✅ NEW: Inject planner
```

---

## Testing Checklist

### ✅ Compilation
- [ ] Code compiles without errors
- [ ] No warnings about unused variables
- [ ] All using statements present

### ✅ Functionality
- [ ] Planning phase runs successfully
- [ ] Zones are filtered correctly (skipped zones removed)
- [ ] Zones are reordered correctly (low risk first)
- [ ] DTO data is used when available
- [ ] Fallback works when DTO unavailable
- [ ] Existing placement logic still works

### ✅ Performance
- [ ] Planning time < 5% of total placement time
- [ ] No significant memory increase
- [ ] Parallel processing works for large batches
- [ ] Sequential processing works for small batches

### ✅ Error Handling
- [ ] Planning failure doesn't crash placement
- [ ] Missing DTO doesn't crash placement
- [ ] Invalid DTO data doesn't crash placement
- [ ] Logging failures don't crash placement

### ✅ Logging
- [ ] Planning metrics logged correctly
- [ ] Batch logs written correctly
- [ ] Error logs written on failure
- [ ] Diagnostic logs show DTO usage

---

## Rollback Instructions

If issues arise, disable planning:

**Option 1: Feature Flag**
```csharp
DeploymentConfiguration.EnableParallelPlanning = false;
```

**Option 2: Remove Planner**
```csharp
_planner = null;  // In constructor
```

**Option 3: Revert Code**
- Remove planning phase code
- Change `zonesToProcess` back to `sortedClashZones`
- Remove DTO usage code

---

## Performance Expectations

### Small Batch (< 12 zones)
- **Planning time:** ~1-5ms
- **Overhead:** Negligible
- **Benefit:** Early skip detection

### Medium Batch (12-100 zones)
- **Planning time:** ~10-50ms
- **Overhead:** < 2% of total time
- **Benefit:** Early skips + reordering

### Large Batch (100+ zones)
- **Planning time:** ~50-200ms
- **Overhead:** < 5% of total time
- **Benefit:** Significant time savings from early skips

---

## Next Steps After Integration

1. **Test with real projects** - Verify behavior matches expectations
2. **Monitor performance** - Track planning metrics
3. **Refine heuristics** - Adjust risk classification if needed
4. **Optimize further** - Replace more calculations with DTO values
5. **Add unit tests** - Test planner in isolation
6. **Document usage** - Update user documentation

---

## Support

If you encounter issues:
1. Check `planning_errors.log` for errors
2. Check `planning_debug.log` for diagnostics
3. Verify feature flag is enabled
4. Test with small batch first
5. Disable planning if needed (feature flag)

