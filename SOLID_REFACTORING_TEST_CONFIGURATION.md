# SOLID Refactoring Test Configuration Guide

**Date:** December 2025  
**Status:** Active Testing Configuration

---

## Overview

This document provides configuration instructions for testing the SOLID-refactored services in the MEP Openings application. All refactored services can be enabled/disabled via feature flags in `OptimizationFlags.cs`.

---

## Quick Reference

### Feature Flags Location
**File:** `Services/OptimizationFlags.cs`

### Main Configuration Flags

| Flag | Default | Purpose | Status |
|------|---------|---------|--------|
| `UseNewSleevePlacerService` | `true` | Enable SOLID-refactored sleeve placement service | ✅ **Active** |
| `UseRefactoredClashZoneFlagServices` | `false` | Enable SOLID-refactored ClashZoneService & FlagManager | ⏳ **In Progress** |

---

## 1. Sleeve Placement Service Testing

### 1.1 Configuration

**Flag:** `OptimizationFlags.UseNewSleevePlacerService`

**Current Default:** `true` (enabled)

**Location in Code:**
```csharp
// Services/OptimizationFlags.cs (line ~282)
public static bool UseNewSleevePlacerService { get; set; } = true;
```

### 1.2 Enable/Disable

**Enable (Test Refactored Service):**
```csharp
// In your test code or initialization
OptimizationFlags.UseNewSleevePlacerService = true;
```

**Disable (Use Legacy - Currently Not Available):**
```csharp
// Note: Legacy UniversalSleevePlacerService is currently disabled
OptimizationFlags.UseNewSleevePlacerService = false;
// Will throw InvalidOperationException - legacy service moved to backup
```

### 1.3 What Gets Tested

When `UseNewSleevePlacerService = true`:

✅ **SOLID Principles:**
- **SRP:** Each service has single responsibility
- **OCP:** Extensible through interfaces (ISleevePlacementStrategy, IClearanceStrategy)
- **LSP:** All implementations substitutable
- **ISP:** Small, focused interfaces
- **DIP:** Depends on abstractions (interfaces), not concrete classes

✅ **Features:**
- Dependency Injection (IFlagManager, IParameterBatchingService, IPerformanceMonitor)
- Deferred parameter batching (4-6× performance boost)
- Parallel clearance calculation support
- Comprehensive crash-safe error handling
- Performance monitoring and logging
- Smart Replay logic (reuses saved data when conditions unchanged)

✅ **Optimizations Preserved:**
- Batch parameter updates
- Family symbol caching
- Spatial grid filtering
- Parallel planning (when enabled)

### 1.4 Testing Checklist

- [ ] **Placement Counts:** Verify same number of sleeves placed as legacy
- [ ] **Performance:** Check `placement_performance.log` for timing metrics
- [ ] **Errors:** Check `placement_errors.log` for any failures
- [ ] **Parameters:** Verify all sleeve parameters set correctly
- [ ] **Flags:** Verify `IsResolved` flags updated correctly
- [ ] **Database:** Verify sleeve data saved to database
- [ ] **Batching:** Verify parameter batching working (check `parameter_batching_performance.log`)

### 1.5 Log Files to Monitor

**Performance Metrics:**
- `placement_performance.log` - Operation timing, item counts, success rates
- `parameter_batching_performance.log` - Batching metrics, flush times

**Errors:**
- `placement_errors.log` - Placement failures, error details

**Service Instantiation:**
- `service_instantiation.log` - Service creation, dependency injection

**Example Log Entry:**
```
[12:34:56] PlaceAllSleeves: 1234ms (100 zones, 12.34ms/zone)
[12:34:56] ParameterBatching: 45ms (500 parameters batched, 0.09ms/param)
[12:34:56] TotalPlaced: 95, TotalSkipped: 5, TotalErrors: 0, SuccessRate: 100%
```

---

## 2. ClashZoneService & FlagManager Testing

### 2.1 Configuration

**Flag:** `OptimizationFlags.UseRefactoredClashZoneFlagServices`

**Current Default:** `false` (disabled - refactoring in progress)

**Location in Code:**
```csharp
// Services/OptimizationFlags.cs (line ~310)
public static bool UseRefactoredClashZoneFlagServices { get; set; } = false;
```

### 2.2 Enable/Disable

**Enable (Test Refactored Services):**
```csharp
// In your test code or initialization
OptimizationFlags.UseRefactoredClashZoneFlagServices = true;
```

**Disable (Use Legacy Services):**
```csharp
OptimizationFlags.UseRefactoredClashZoneFlagServices = false;
// Uses legacy ClashZoneService and FlagManager (concrete classes)
```

### 2.3 What Gets Tested

When `UseRefactoredClashZoneFlagServices = true`:

✅ **SOLID Principles:**
- **SRP:** Split responsibilities (Cleanup, Filter, Validation, Flag Management)
- **OCP:** Extensible through interfaces
- **LSP:** All implementations substitutable
- **ISP:** Small, focused interfaces
- **DIP:** Depends on abstractions only

✅ **Interfaces Created:**
- `IClashZoneService` - Main orchestrator
- `IClashZoneCleanupService` - Cleanup operations
- `IClashZoneFilterService` - Filtering operations
- `IClashZoneValidationService` - Validation operations
- `IFlagManager` - Flag management
- `IInstanceIdManager` - Instance ID management
- `ISessionTracker` - Session tracking (removes static state)

✅ **Optimizations Preserved:**
- Batch flag updates (`BatchUpdateFlags()` - 4-6× faster)
- Batch collection (collect all sleeves once)
- HashSet lookups (O(1) performance)
- Pre-loaded entries (calculate once, use many times)
- SectionBoxHelper reuse

### 2.4 Testing Checklist

- [ ] **Flag Reset:** Verify flags reset correctly for deleted sleeves
- [ ] **Batch Updates:** Verify `BatchUpdateFlags()` still called (check logs)
- [ ] **SectionBox:** Verify section box filtering works (uses SectionBoxHelper)
- [ ] **Performance:** No performance regression (within 2% of baseline)
- [ ] **Data Integrity:** Flag hierarchy preserved (cluster flags first)
- [ ] **Revit Verification:** Revit API verification still works (authoritative source)
- [ ] **Session Tracking:** Recently placed cluster sleeves protected from deletion

### 2.5 Implementation Status

**Current Status:** ⏳ **In Progress**

**See:** `CLASHZONE_FLAGMANAGER_3TEAM_IMPLEMENTATION_PLAN.md` for detailed implementation plan

**Teams:**
- **Team F:** FlagManager Refactoring
- **Team G:** ClashZoneService Refactoring
- **Team H:** Integration & Wiring

**Timeline:** 4 weeks (Interface Definition → Implementation → Wiring → Testing)

---

## 3. Complete Test Configuration

### 3.1 Full SOLID Refactoring Test

**Enable All Refactored Services:**
```csharp
// Enable SOLID-refactored sleeve placement
OptimizationFlags.UseNewSleevePlacerService = true;

// Enable SOLID-refactored ClashZoneService & FlagManager (when ready)
OptimizationFlags.UseRefactoredClashZoneFlagServices = true;

// Enable related optimizations
OptimizationFlags.UseBatchedParameterWrites = true;
OptimizationFlags.UseParallelClearanceCalculation = true;
OptimizationFlags.EnablePerformanceLogging = true;
```

### 3.2 Legacy Fallback Test

**Disable All Refactored Services:**
```csharp
// Disable SOLID-refactored services
OptimizationFlags.UseNewSleevePlacerService = false;
OptimizationFlags.UseRefactoredClashZoneFlagServices = false;

// Note: Legacy UniversalSleevePlacerService is currently disabled
// This will throw InvalidOperationException
```

### 3.3 Hybrid Test (Recommended for Gradual Migration)

**Enable Refactored Sleeve Placement, Keep Legacy ClashZone/Flag:**
```csharp
// Use refactored sleeve placement
OptimizationFlags.UseNewSleevePlacerService = true;

// Keep legacy ClashZoneService & FlagManager
OptimizationFlags.UseRefactoredClashZoneFlagServices = false;
```

---

## 4. Performance Testing

### 4.1 Baseline Metrics

**Before Enabling Refactored Services:**
1. Run placement on test model
2. Record:
   - Total placement time
   - Sleeves placed count
   - Errors count
   - Memory usage

### 4.2 Refactored Metrics

**After Enabling Refactored Services:**
1. Run placement on same test model
2. Compare:
   - Total placement time (should be within 2% of baseline)
   - Sleeves placed count (should match exactly)
   - Errors count (should be same or lower)
   - Memory usage (should be similar or better)

### 4.3 Performance Logs

**Check These Log Files:**
- `placement_performance.log` - Overall timing
- `parameter_batching_performance.log` - Batching metrics
- `service_instantiation.log` - Service creation timing

**Key Metrics to Compare:**
- `PlaceAllSleeves` - Total placement time
- `ParameterBatching` - Batching time (should be 4-6× faster)
- `BatchUpdateFlags` - Flag update time (should be 4-6× faster)
- `TotalPlaced` - Should match baseline
- `SuccessRate` - Should be 100% or higher

---

## 5. Error Testing

### 5.1 Fail-Safe Mechanisms

**Test Scenarios:**
1. **Missing Family:** Should log error, continue with next zone
2. **Invalid Dimensions:** Should use fallbacks, continue
3. **Database Failure:** Should rollback transaction, log error
4. **Parameter Write Failure:** Should log error, continue

**Expected Behavior:**
- ✅ No unhandled exceptions
- ✅ Errors logged to `placement_errors.log`
- ✅ Processing continues for other zones
- ✅ Summary shows error count

### 5.2 Error Logs

**Check:** `placement_errors.log`

**Example Entry:**
```
[12:34:56] [ERROR] Failed to place sleeve for ClashZone {guid}: Family not found
[12:34:56] [ERROR] Failed to update parameter for sleeve {id}: Parameter not found
```

---

## 6. Integration Testing

### 6.1 End-to-End Test

**Test Flow:**
1. Refresh → Detect intersections → Save to database
2. Place Sleeves → Use refactored service → Update flags
3. Cluster Sleeves → Use refactored cluster service → Update cluster flags
4. Verify → Check all sleeves placed, flags updated, database consistent

### 6.2 Database Verification

**Check Database:**
```sql
-- Verify sleeves placed
SELECT COUNT(*) FROM ClashZones WHERE SleeveInstanceId > 0;

-- Verify flags updated
SELECT COUNT(*) FROM ClashZones WHERE IsResolvedFlag = 1;

-- Verify cluster sleeves placed
SELECT COUNT(*) FROM ClashZones WHERE IsClusterResolvedFlag = 1;
```

---

## 7. Rollback Plan

### 7.1 Quick Disable

**If Issues Found:**
```csharp
// Disable refactored services immediately
OptimizationFlags.UseNewSleevePlacerService = false;
OptimizationFlags.UseRefactoredClashZoneFlagServices = false;
```

**Note:** Legacy `UniversalSleevePlacerService` is currently disabled, so this will throw an exception. Re-enable legacy service first if needed.

### 7.2 Gradual Rollback

**Step 1:** Disable only problematic service
```csharp
// Keep refactored sleeve placement, disable ClashZone/Flag refactoring
OptimizationFlags.UseNewSleevePlacerService = true;
OptimizationFlags.UseRefactoredClashZoneFlagServices = false;
```

**Step 2:** Monitor logs for issues
**Step 3:** Re-enable when fixed

---

## 8. Configuration Best Practices

### 8.1 Development Mode

**Recommended Settings:**
```csharp
OptimizationFlags.UseNewSleevePlacerService = true;
OptimizationFlags.EnablePerformanceLogging = true;
OptimizationFlags.UseRefactoredClashZoneFlagServices = false; // Until ready
```

### 8.2 Production Mode

**Recommended Settings:**
```csharp
OptimizationFlags.UseNewSleevePlacerService = true; // Tested and stable
OptimizationFlags.EnablePerformanceLogging = false; // Disabled in deployment
OptimizationFlags.UseRefactoredClashZoneFlagServices = false; // Until fully tested
```

### 8.3 Testing Mode

**Recommended Settings:**
```csharp
OptimizationFlags.UseNewSleevePlacerService = true;
OptimizationFlags.EnablePerformanceLogging = true;
OptimizationFlags.UseRefactoredClashZoneFlagServices = true; // When ready
DeploymentConfiguration.DeploymentMode = false; // Enable verbose logging
```

---

## 9. Troubleshooting

### 9.1 Common Issues

**Issue:** "Legacy placement service is disabled"
**Solution:** Set `OptimizationFlags.UseNewSleevePlacerService = true`

**Issue:** Performance regression
**Solution:** Check `placement_performance.log`, verify optimizations enabled

**Issue:** Missing interfaces
**Solution:** Ensure all Team F/G interfaces created (see implementation plan)

**Issue:** Dependency injection failures
**Solution:** Check factory classes created correctly (Team H wiring)

### 9.2 Debug Mode

**Enable Debug Logging:**
```csharp
DeploymentConfiguration.DeploymentMode = false;
OptimizationFlags.EnablePerformanceLogging = true;
```

**Check Log Files:**
- `placement_debug.log` - Detailed placement operations
- `placement_performance.log` - Performance metrics
- `service_instantiation.log` - Service creation details

---

## 10. Success Criteria

### 10.1 Functional Requirements

- ✅ All sleeves placed correctly
- ✅ All parameters set correctly
- ✅ All flags updated correctly
- ✅ Database consistent
- ✅ No data loss

### 10.2 Performance Requirements

- ✅ Placement time within 2% of baseline
- ✅ Parameter batching 4-6× faster
- ✅ Flag updates 4-6× faster
- ✅ Memory usage similar or better

### 10.3 Quality Requirements

- ✅ No unhandled exceptions
- ✅ All errors logged
- ✅ Fail-safe mechanisms working
- ✅ SOLID principles followed

---

**Document Status:** ✅ Complete  
**Last Updated:** December 2025  
**Related Documents:**
- `CLASHZONE_FLAGMANAGER_3TEAM_IMPLEMENTATION_PLAN.md`
- `CLASHZONE_FLAGMANAGER_SOLID_ANALYSIS.md`
- `SLEEVE_PLACEMENT_METHODOLOGY_REFACTORED.md`

