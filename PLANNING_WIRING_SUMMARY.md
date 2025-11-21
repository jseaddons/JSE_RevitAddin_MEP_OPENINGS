# Planning Layer Wiring - Quick Reference

## 📋 Overview
Complete wiring plan to integrate `ParallelSleevePlacementPlanner` into `UniversalSleevePlacerService`.

## ✅ Files to Modify

### 1. `Services/DeploymentConfiguration.cs`
**Add feature flag:**
```csharp
/// <summary>
/// Enable parallel planning layer for sleeve placement.
/// When enabled, pre-computes sleeve dimensions, clearance, and risk classification in parallel.
/// Default: true (enabled)
/// </summary>
public static bool EnableParallelPlanning { get; set; } = true;
```

### 2. `Services/UniversalSleevePlacerService.cs`

#### 2.1 Add Using Statement (Top of file)
```csharp
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;
```

#### 2.2 Add Field (After line ~40)
```csharp
// ✅ NEW: Optional planner for parallel pre-computation
private readonly ISleevePlacementPlanner? _planner;
```

#### 2.3 Update Constructor (Line ~49)
**Add parameter:**
```csharp
ISleevePlacementPlanner? planner = null
```

**Add initialization (after line ~66):**
```csharp
// ✅ NEW: Initialize planner (create default if not provided)
_planner = planner ?? new ParallelSleevePlacementPlanner();
```

#### 2.4 Add Planning Phase (Line ~520, before foreach loop)
**See:** `PLANNING_INTEGRATION_IMPLEMENTATION.md` Step 2 for complete code

#### 2.5 Use DTO in Placement Loop (Line ~600, after validation)
**See:** `PLANNING_INTEGRATION_IMPLEMENTATION.md` Step 3 for complete code

#### 2.6 Add Batch Logging (Line ~2568, before return)
**See:** `PLANNING_INTEGRATION_IMPLEMENTATION.md` Step 5 for complete code

## 🔧 Key Integration Points

### Planning Phase Location
- **File:** `UniversalSleevePlacerService.cs`
- **Method:** `PlaceAllSleevesInTransaction`
- **Line:** ~520 (after `placementLoopTracker`, before `foreach`)

### DTO Usage Location
- **File:** `UniversalSleevePlacerService.cs`
- **Method:** `PlaceAllSleevesInTransaction`
- **Line:** ~600 (inside placement loop, after validation)

## 📊 Expected Benefits

1. **Early Skip Detection** - Zones marked for skip filtered before Revit API calls
2. **Reordering** - Low-risk zones processed first (better success rate)
3. **Performance** - Parallel pre-computation saves 10-20% total time
4. **Diagnostics** - Risk classification for better visibility

## 🛡️ Safety Features

- ✅ **Feature Flag** - Can be disabled instantly
- ✅ **Graceful Fallback** - Falls back to original logic on error
- ✅ **Backward Compatible** - Works with existing code
- ✅ **No Breaking Changes** - All existing logic preserved

## 📝 Testing Checklist

- [ ] Code compiles without errors
- [ ] Planning phase runs successfully
- [ ] Zones filtered correctly (skipped removed)
- [ ] Zones reordered correctly (low risk first)
- [ ] DTO data used when available
- [ ] Fallback works when DTO unavailable
- [ ] Feature flag works (enable/disable)
- [ ] Error handling works (planning failure doesn't crash)
- [ ] Logging works (planning metrics logged)

## 🚀 Quick Start

1. **Add feature flag** to `DeploymentConfiguration.cs`
2. **Add planner field** to `UniversalSleevePlacerService`
3. **Update constructor** to accept planner
4. **Add planning phase** before placement loop
5. **Use DTO data** in placement loop
6. **Test with small batch** first
7. **Monitor performance** metrics

## 📚 Detailed Documentation

- **Architecture Plan:** `PLANNING_LAYER_WIRING_PLAN.md`
- **Step-by-Step Guide:** `PLANNING_INTEGRATION_IMPLEMENTATION.md`
- **This Summary:** `PLANNING_WIRING_SUMMARY.md`

## ⚠️ Important Notes

1. **Feature Flag Default:** `true` (enabled by default)
2. **Planner Optional:** Can be `null` (backward compatible)
3. **Fallback Always Works:** Original logic preserved
4. **No Breaking Changes:** All existing code unchanged

## 🔄 Rollback

If issues arise:
```csharp
DeploymentConfiguration.EnableParallelPlanning = false;
```

This instantly disables planning and uses original logic.

