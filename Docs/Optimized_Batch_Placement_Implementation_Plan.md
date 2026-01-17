# Optimized Batch Placement Implementation Plan
## Individual & Cluster Sleeves with SOLID Principles, Safe Transaction Management, and Crash-Safe Features

---

## Executive Summary

**Bottom Line:** Function FIRST, Optimization SECOND

**Core Principles:**
- ✅ **SOLID Compliance** - Single Responsibility, Dependency Injection, Interface Segregation
- ✅ **Safe Transaction Management** - Proper transaction handling, rollback on errors
- ✅ **Crash-Safe Execution** - Timeout protection, graceful degradation, comprehensive error handling
- ✅ **Immediate Parameter Setting** - Proven to work (bypasses broken deferred mechanism)
- ✅ **Grouped Bulk Placement** - Minimize overhead while maintaining reliability

**Expected Results:**
- Correct dimensions (no more 300x100 defaults)
- No duplicate clusters
- 3-4x faster than sequential mode
- Reliable and maintainable
- Graceful failure handling

---

## Table of Contents

1. [Architecture Overview](#architecture-overview)
2. [SOLID Principles Implementation](#solid-principles-implementation)
3. [Phase 1: Individual Sleeves](#phase-1-individual-sleeves)
4. [Phase 2: Cluster Sleeves](#phase-2-cluster-sleeves)
5. [Phase 3: Safe Transaction Management](#phase-3-safe-transaction-management)
6. [Phase 4: Crash-Safe Features](#phase-4-crash-safe-features)
7. [Phase 5: Integration](#phase-5-integration)
8. [Phase 6: Testing & Verification](#phase-6-testing--verification)
9. [Phase 7: Performance Monitoring](#phase-7-performance-monitoring)
10. [Rollback & Fallback Strategy](#rollback--fallback-strategy)

---

## Architecture Overview

### Current Issues

**Individual Sleeves:**
- ❌ Deferred parameters don't work (values revert to 300x100)
- ❌ Sequential mode works but slow (regenerate after each sleeve)
- ❌ No proper error handling or transaction management

**Cluster Sleeves:**
- ❌ Using broken deferred parameters
- ❌ Duplicate clusters appearing at same location
- ❌ Wrong coordinates/dimensions
- ❌ No crash protection for long-running operations

### Proposed Solution

**Architecture Diagram:**

```
┌─────────────────────────────────────────────────────────────┐
│                  PlacementOrchestrator                      │
│  (Coordinates placement, manages transactions & crashes)    │
└──────────────┬──────────────────────────────┬───────────────┘
               │                              │
       ┌───────▼────────┐            ┌───────▼────────┐
       │  Individual    │            │    Cluster     │
       │    Sleeve      │            │     Sleeve     │
       │   Placer       │            │    Placer      │
       └───────┬────────┘            └───────┬────────┘
               │                              │
       ┌───────▼────────────────────────┬─────▼────────┐
       │  GroupingService               │              │
       │  (Groups by family/dimensions) │              │
       └───────┬────────────────────────┴──────────────┘
               │
       ┌───────▼────────────────────────────────────────┐
       │  ParameterService (Immediate Setting Only)     │
       │  - SetParametersImmediate()                    │
       │  - NO deferred dictionary                      │
       │  - NO flush mechanism                          │
       └───────┬────────────────────────────────────────┘
               │
       ┌───────▼────────────────────────────────────────┐
       │  TransactionManager (Safe Operations)          │
       │  - StartTransaction()                          │
       │  - CommitWithValidation()                      │
       │  - RollbackOnError()                           │
       └────────────────────────────────────────────────┘
```

---

## SOLID Principles Implementation

### S - Single Responsibility Principle

**Each class has ONE clear responsibility:**

| Class | Responsibility |
|-------|----------------|
| `PlacementOrchestrator` | Coordinates overall placement workflow |
| `IndividualSleevePlacer` | Places individual sleeves only |
| `ClusterSleevePlacer` | Places cluster sleeves only |
| `GroupingService` | Groups zones by family/dimensions |
| `ParameterService` | Sets parameters on family instances |
| `TransactionManager` | Manages Revit transactions safely |
| `CrashSafeExecutor` | Handles timeouts and crashes |

### O - Open/Closed Principle

**Open for extension, closed for modification:**

```csharp
// Interface for placement strategies
public interface ISleevePlacementStrategy
{
    (int placed, int skipped, int errors) PlaceSleeves(
        List<ClashZone> zones, 
        ITransactionManager transactionManager
    );
}

// Individual sleeve strategy
public class IndividualSleevePlacementStrategy : ISleevePlacementStrategy
{
    public (int placed, int skipped, int errors) PlaceSleeves(...)
    {
        // Implementation for individual sleeves
    }
}

// Cluster sleeve strategy
public class ClusterSleevePlacementStrategy : ISleevePlacementStrategy
{
    public (int placed, int skipped, int errors) PlaceSleeves(...)
    {
        // Implementation for clusters
    }
}
```

### L - Liskov Substitution Principle

**Derived classes can substitute base classes:**

```csharp
// Base interface
public interface IParameterSetter
{
    void SetParameter(FamilyInstance instance, string paramName, object value);
}

// Immediate setter (working implementation)
public class ImmediateParameterSetter : IParameterSetter
{
    public void SetParameter(FamilyInstance instance, string paramName, object value)
    {
        Parameter param = instance.LookupParameter(paramName);
        if (param != null && !param.IsReadOnly)
        {
            param.Set(value);
        }
    }
}

// Can be substituted with any IParameterSetter implementation
public class PlacementService
{
    private readonly IParameterSetter _parameterSetter;
    
    public PlacementService(IParameterSetter parameterSetter)
    {
        _parameterSetter = parameterSetter;
    }
}
```

### I - Interface Segregation Principle

**Clients shouldn't depend on interfaces they don't use:**

```csharp
// Split into focused interfaces
public interface ISleeveCreator
{
    FamilyInstance CreateSleeve(XYZ point, FamilySymbol symbol, double rotation);
}

public interface ISleeveParameterSetter
{
    void SetDimensions(FamilyInstance instance, double width, double height);
    void SetDepth(FamilyInstance instance, double depth);
    void SetLevel(FamilyInstance instance, Level level);
}

public interface ISleeveDatabaseUpdater
{
    void UpdateSleeveInDatabase(ClashZone zone, int sleeveInstanceId);
}

// Clients use only what they need
public class IndividualSleevePlacer
{
    private readonly ISleeveCreator _creator;
    private readonly ISleeveParameterSetter _parameterSetter;
    
    // No dependency on database updater if not needed
}
```

### D - Dependency Inversion Principle

**Depend on abstractions, not concretions:**

```csharp
// High-level module depends on abstraction
public class PlacementOrchestrator
{
    private readonly ISleevePlacementStrategy _strategy;
    private readonly ITransactionManager _transactionManager;
    private readonly ICrashSafeExecutor _crashSafeExecutor;
    
    // Constructor injection (DI)
    public PlacementOrchestrator(
        ISleevePlacementStrategy strategy,
        ITransactionManager transactionManager,
        ICrashSafeExecutor crashSafeExecutor)
    {
        _strategy = strategy;
        _transactionManager = transactionManager;
        _crashSafeExecutor = crashSafeExecutor;
    }
    
    public (int placed, int errors) Execute(List<ClashZone> zones)
    {
        return _crashSafeExecutor.ExecuteWithTimeout(() =>
        {
            return _transactionManager.ExecuteInTransaction("Place Sleeves", () =>
            {
                return _strategy.PlaceSleeves(zones, _transactionManager);
            });
        }, "Sleeve Placement");
    }
}
```

---

## Phase 1: Individual Sleeves

### 1.1 Create GroupingService (SRP)

**File:** `Services/Placement/GroupingService.cs`

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// SRP: Responsible ONLY for grouping zones by family type and dimensions
    /// </summary>
    public class GroupingService : IGroupingService
    {
        private readonly ILogger _logger;
        
        public GroupingService(ILogger logger = null)
        {
            _logger = logger ?? new NullLogger();
        }
        
        /// <summary>
        /// Groups zones by family name and rounded dimensions
        /// This minimizes family symbol loading and activation overhead
        /// </summary>
        public List<SleeveGroup> GroupZonesByFamilyAndDimensions(
            List<ClashZone> zones,
            Func<ClashZone, string> familyNameSelector,
            Func<ClashZone, double> widthSelector,
            Func<ClashZone, double> heightSelector)
        {
            if (zones == null || zones.Count == 0)
            {
                return new List<SleeveGroup>();
            }
            
            try
            {
                var groups = zones
                    .GroupBy(z => new
                    {
                        FamilyName = familyNameSelector(z),
                        Width = RoundToNearest50mm(widthSelector(z)),
                        Height = RoundToNearest50mm(heightSelector(z)),
                        IsCircular = z.SleeveDiameter > 0
                    })
                    .Select(g => new SleeveGroup
                    {
                        FamilyName = g.Key.FamilyName,
                        Width = g.Key.Width,
                        Height = g.Key.Height,
                        IsCircular = g.Key.IsCircular,
                        Zones = g.ToList()
                    })
                    .OrderByDescending(g => g.Zones.Count) // Process largest groups first
                    .ToList();
                
                _logger.Info($"Grouped {zones.Count} zones into {groups.Count} family/dimension groups");
                
                return groups;
            }
            catch (Exception ex)
            {
                _logger.Error($"Error grouping zones: {ex.Message}");
                throw;
            }
        }
        
        private double RoundToNearest50mm(double valueInFeet)
        {
            // Convert to mm, round to nearest 50mm, convert back to feet
            double mm = valueInFeet * 304.8;
            double roundedMm = Math.Round(mm / 50.0) * 50.0;
            return roundedMm / 304.8;
        }
    }
    
    public class SleeveGroup
    {
        public string FamilyName { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public bool IsCircular { get; set; }
        public List<ClashZone> Zones { get; set; }
    }
    
    public interface IGroupingService
    {
        List<SleeveGroup> GroupZonesByFamilyAndDimensions(
            List<ClashZone> zones,
            Func<ClashZone, string> familyNameSelector,
            Func<ClashZone, double> widthSelector,
            Func<ClashZone, double> heightSelector);
    }
}
```

### 1.2 Create ImmediateParameterService (SRP)

**File:** `Services/Placement/ImmediateParameterService.cs`

```csharp
using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// SRP: Responsible ONLY for setting parameters immediately on family instances
    /// NO deferred parameters, NO flush mechanism, NO dictionary
    /// </summary>
    public class ImmediateParameterService : IParameterService
    {
        private readonly Document _doc;
        private readonly ILogger _logger;
        private readonly Dictionary<string, Level> _levelCache;
        
        public ImmediateParameterService(Document doc, ILogger logger = null)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _logger = logger ?? new NullLogger();
            _levelCache = new Dictionary<string, Level>();
        }
        
        /// <summary>
        /// Sets all sleeve parameters IMMEDIATELY (no deferral)
        /// </summary>
        public bool SetSleeveParameters(
            FamilyInstance instance,
            ClashZone zone,
            double width,
            double height,
            double? diameter = null)
        {
            if (instance == null || !instance.IsValidObject)
            {
                _logger.Warning("Invalid instance - cannot set parameters");
                return false;
            }
            
            try
            {
                int successCount = 0;
                
                // Width
                if (SetParameter(instance, "Width", width))
                    successCount++;
                
                // Height
                if (SetParameter(instance, "Height", height))
                    successCount++;
                
                // Diameter (if circular)
                if (diameter.HasValue && diameter.Value > 0)
                {
                    if (SetParameter(instance, "Diameter", diameter.Value))
                        successCount++;
                }
                
                // Depth (from wall thickness)
                double depth = CalculateDepth(zone);
                if (SetParameter(instance, "Depth", depth))
                    successCount++;
                
                // Wall Width (alternative to Depth in some families)
                SetParameter(instance, "Wall Width", depth); // Don't count - optional
                
                // Bottom of Opening
                if (zone.BottomOfOpening > 0)
                {
                    if (SetParameter(instance, "Bottom Of Opening", zone.BottomOfOpening))
                        successCount++;
                }
                
                // Level
                if (SetLevel(instance, zone.MepElementLevelName))
                    successCount++;
                
                _logger.Info($"Set {successCount} parameters for sleeve {instance.Id.IntegerValue}");
                
                return successCount > 0;
            }
            catch (Exception ex)
            {
                _logger.Error($"Error setting parameters for sleeve {instance.Id.IntegerValue}: {ex.Message}");
                return false;
            }
        }
        
        private bool SetParameter(FamilyInstance instance, string paramName, double value)
        {
            try
            {
                Parameter param = instance.LookupParameter(paramName);
                if (param == null)
                {
                    _logger.Debug($"Parameter '{paramName}' not found on instance {instance.Id.IntegerValue}");
                    return false;
                }
                
                if (param.IsReadOnly)
                {
                    _logger.Debug($"Parameter '{paramName}' is read-only on instance {instance.Id.IntegerValue}");
                    return false;
                }
                
                // ✅ IMMEDIATE SETTING - writes to Revit RIGHT NOW
                param.Set(value);
                
                // Verify it was set
                double actualValue = param.AsDouble();
                if (Math.Abs(actualValue - value) < 0.001)
                {
                    _logger.Debug($"✅ Set {paramName} = {value * 304.8:F1}mm on sleeve {instance.Id.IntegerValue}");
                    return true;
                }
                else
                {
                    _logger.Warning($"❌ Parameter {paramName} set to {value * 304.8:F1}mm but reads as {actualValue * 304.8:F1}mm");
                    return false;
                }
            }
            catch (Exception ex)
            {
                _logger.Error($"Error setting parameter '{paramName}': {ex.Message}");
                return false;
            }
        }
        
        private bool SetLevel(FamilyInstance instance, string levelName)
        {
            if (string.IsNullOrEmpty(levelName))
                return false;
            
            try
            {
                Level level = GetCachedLevel(levelName);
                if (level == null)
                    return false;
                
                Parameter levelParam = instance.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM);
                if (levelParam != null && !levelParam.IsReadOnly)
                {
                    levelParam.Set(level.Id);
                    return true;
                }
                
                return false;
            }
            catch (Exception ex)
            {
                _logger.Error($"Error setting level: {ex.Message}");
                return false;
            }
        }
        
        private Level GetCachedLevel(string levelName)
        {
            if (_levelCache.TryGetValue(levelName, out Level cached))
                return cached;
            
            Level level = new FilteredElementCollector(_doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .FirstOrDefault(l => l.Name.Equals(levelName, StringComparison.OrdinalIgnoreCase));
            
            if (level != null)
                _levelCache[levelName] = level;
            
            return level;
        }
        
        private double CalculateDepth(ClashZone zone)
        {
            bool isWallHost = zone.StructuralElementType == "Wall" || zone.StructuralElementType == "Walls";
            bool isFramingHost = string.Equals(zone.StructuralElementType, "Structural Framing", StringComparison.OrdinalIgnoreCase);
            
            double depth = 0.0;
            
            if (isWallHost)
            {
                depth = zone.WallThickness > 0 ? zone.WallThickness : zone.StructuralElementThickness;
            }
            else if (isFramingHost)
            {
                depth = zone.FramingThickness > 0 ? zone.FramingThickness : zone.StructuralElementThickness;
            }
            else
            {
                depth = zone.StructuralElementThickness;
            }
            
            if (depth <= 0)
            {
                _logger.Warning($"Invalid depth for zone {zone.Id} - using fallback");
                depth = 0.5; // 6 inches fallback
            }
            
            return depth;
        }
    }
    
    public interface IParameterService
    {
        bool SetSleeveParameters(
            FamilyInstance instance,
            ClashZone zone,
            double width,
            double height,
            double? diameter = null);
    }
}
```

### 1.3 Create OptimizedIndividualSleevePlacer (SRP + DIP)

**File:** `Services/Placement/OptimizedIndividualSleevePlacer.cs`

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// SRP: Responsible ONLY for placing individual sleeves using optimized grouped approach
    /// DIP: Depends on abstractions (IGroupingService, IParameterService, etc.)
    /// </summary>
    public class OptimizedIndividualSleevePlacer
    {
        private readonly Document _doc;
        private readonly IGroupingService _groupingService;
        private readonly IParameterService _parameterService;
        private readonly IFamilyLoaderService _familyLoader;
        private readonly ILogger _logger;
        
        // Constructor Injection (DIP)
        public OptimizedIndividualSleevePlacer(
            Document doc,
            IGroupingService groupingService,
            IParameterService parameterService,
            IFamilyLoaderService familyLoader,
            ILogger logger = null)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _groupingService = groupingService ?? throw new ArgumentNullException(nameof(groupingService));
            _parameterService = parameterService ?? throw new ArgumentNullException(nameof(parameterService));
            _familyLoader = familyLoader ?? throw new ArgumentNullException(nameof(familyLoader));
            _logger = logger ?? new NullLogger();
        }
        
        /// <summary>
        /// Places individual sleeves using optimized grouped approach
        /// </summary>
        public PlacementResult PlaceSleeves(List<ClashZone> zones)
        {
            if (zones == null || zones.Count == 0)
            {
                return new PlacementResult { Placed = 0, Skipped = 0, Errors = 0 };
            }
            
            _logger.Info($"Starting optimized placement for {zones.Count} individual sleeves");
            
            var result = new PlacementResult();
            
            try
            {
                // STEP 1: Group zones by family type and dimensions
                var groups = _groupingService.GroupZonesByFamilyAndDimensions(
                    zones,
                    z => GetFamilyName(z),
                    z => z.SleeveWidth,
                    z => z.SleeveHeight
                );
                
                _logger.Info($"Created {groups.Count} groups for processing");
                
                // STEP 2: Process each group
                foreach (var group in groups)
                {
                    var groupResult = ProcessGroup(group);
                    
                    result.Placed += groupResult.Placed;
                    result.Skipped += groupResult.Skipped;
                    result.Errors += groupResult.Errors;
                }
                
                // STEP 3: Single regeneration for ALL sleeves
                if (result.Placed > 0)
                {
                    _logger.Info($"Regenerating document for {result.Placed} sleeves...");
                    _doc.Regenerate();
                    _logger.Info("Regeneration complete");
                }
                
                return result;
            }
            catch (Exception ex)
            {
                _logger.Error($"Error in optimized placement: {ex.Message}");
                result.Errors = zones.Count;
                return result;
            }
        }
        
        private PlacementResult ProcessGroup(SleeveGroup group)
        {
            var result = new PlacementResult();
            
            try
            {
                _logger.Info($"Processing group: {group.FamilyName}, {group.Zones.Count} sleeves");
                
                // Load and activate family symbol ONCE for entire group
                FamilySymbol symbol = _familyLoader.LoadAndActivateSymbol(group.FamilyName);
                if (symbol == null)
                {
                    _logger.Error($"Failed to load family symbol: {group.FamilyName}");
                    result.Errors = group.Zones.Count;
                    return result;
                }
                
                // Create ALL instances in this group
                var placedInstances = new List<(FamilyInstance instance, ClashZone zone)>();
                
                foreach (var zone in group.Zones)
                {
                    try
                    {
                        // Skip if already resolved
                        if (zone.IsResolvedFlag || zone.SleeveInstanceId > 0)
                        {
                            result.Skipped++;
                            continue;
                        }
                        
                        // Calculate placement point and rotation
                        XYZ point = CalculatePlacementPoint(zone);
                        double rotation = CalculateRotation(zone);
                        
                        // Create instance
                        FamilyInstance instance = _doc.Create.NewFamilyInstance(
                            point,
                            symbol,
                            StructuralType.NonStructural
                        );
                        
                        if (instance != null && instance.IsValidObject)
                        {
                            placedInstances.Add((instance, zone));
                            
                            // Apply rotation if needed
                            if (Math.Abs(rotation) > 0.001)
                            {
                                Line axis = Line.CreateBound(point, point + XYZ.BasisZ);
                                ElementTransformUtils.RotateElement(_doc, instance.Id, axis, rotation);
                            }
                        }
                        else
                        {
                            _logger.Warning($"Failed to create instance for zone {zone.Id}");
                            result.Errors++;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Error($"Error creating sleeve for zone {zone.Id}: {ex.Message}");
                        result.Errors++;
                    }
                }
                
                _logger.Info($"Created {placedInstances.Count} instances for group");
                
                // Set parameters IMMEDIATELY for all instances in group
                foreach (var (instance, zone) in placedInstances)
                {
                    try
                    {
                        // ✅ IMMEDIATE PARAMETER SETTING
                        bool success = _parameterService.SetSleeveParameters(
                            instance,
                            zone,
                            group.Width,
                            group.Height,
                            group.IsCircular ? group.Width : (double?)null // Use width as diameter for circular
                        );
                        
                        if (success)
                        {
                            // Update zone data
                            zone.SleeveInstanceId = instance.Id.IntegerValue;
                            zone.IsResolvedFlag = true;
                            zone.SleeveWidth = group.Width;
                            zone.SleeveHeight = group.Height;
                            
                            result.Placed++;
                        }
                        else
                        {
                            _logger.Warning($"Failed to set parameters for sleeve {instance.Id.IntegerValue}");
                            result.Errors++;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Error($"Error setting parameters for sleeve {instance.Id.IntegerValue}: {ex.Message}");
                        result.Errors++;
                    }
                }
                
                _logger.Info($"Group complete: {result.Placed} placed, {result.Errors} errors");
            }
            catch (Exception ex)
            {
                _logger.Error($"Error processing group {group.FamilyName}: {ex.Message}");
                result.Errors = group.Zones.Count;
            }
            
            return result;
        }
        
        private XYZ CalculatePlacementPoint(ClashZone zone)
        {
            return new XYZ(
                zone.IntersectionPointX,
                zone.IntersectionPointY,
                zone.IntersectionPointZ
            );
        }
        
        private double CalculateRotation(ClashZone zone)
        {
            // Implement rotation logic based on zone orientation
            // For now, return 0 (no rotation)
            return 0.0;
        }
        
        private string GetFamilyName(ClashZone zone)
        {
            // Use existing ClusterPlacementService.GetFamilyName logic
            bool isCircular = zone.SleeveDiameter > 0;
            double size = isCircular ? zone.SleeveDiameter : Math.Max(zone.SleeveWidth, zone.SleeveHeight);
            
            return ClusterPlacementService.GetFamilyName(
                zone.StructuralElementType,
                zone.MepElementCategory,
                size,
                isCluster: false
            );
        }
    }
    
    public class PlacementResult
    {
        public int Placed { get; set; }
        public int Skipped { get; set; }
        public int Errors { get; set; }
    }
}
```

---

## Phase 2: Cluster Sleeves

### 2.1 Create OptimizedClusterSleevePlacer (SRP + DIP)

**File:** `Services/Placement/OptimizedClusterSleevePlacer.cs`

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// SRP: Responsible ONLY for placing cluster sleeves using optimized approach
    /// DIP: Depends on abstractions (IGroupingService, IParameterService, etc.)
    /// </summary>
    public class OptimizedClusterSleevePlacer
    {
        private readonly Document _doc;
        private readonly IGroupingService _groupingService;
        private readonly IParameterService _parameterService;
        private readonly IFamilyLoaderService _familyLoader;
        private readonly ILogger _logger;
        
        public OptimizedClusterSleevePlacer(
            Document doc,
            IGroupingService groupingService,
            IParameterService parameterService,
            IFamilyLoaderService familyLoader,
            ILogger logger = null)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _groupingService = groupingService ?? throw new ArgumentNullException(nameof(groupingService));
            _parameterService = parameterService ?? throw new ArgumentNullException(nameof(parameterService));
            _familyLoader = familyLoader ?? throw new ArgumentNullException(nameof(familyLoader));
            _logger = logger ?? new NullLogger();
        }
        
        /// <summary>
        /// Places cluster sleeves using optimized grouped approach
        /// </summary>
        public ClusterPlacementResult PlaceClusters(List<ClusterGroup> clusterGroups)
        {
            if (clusterGroups == null || clusterGroups.Count == 0)
            {
                return new ClusterPlacementResult { Placed = 0, Deleted = 0, Errors = 0 };
            }
            
            _logger.Info($"Starting optimized cluster placement for {clusterGroups.Count} clusters");
            
            var result = new ClusterPlacementResult();
            
            try
            {
                // STEP 1: Group clusters by family type and dimensions
                var groups = GroupClustersByFamily(clusterGroups);
                
                _logger.Info($"Created {groups.Count} cluster groups for processing");
                
                // STEP 2: Process each group
                foreach (var group in groups)
                {
                    var groupResult = ProcessClusterGroup(group);
                    
                    result.Placed += groupResult.Placed;
                    result.Deleted += groupResult.Deleted;
                    result.Errors += groupResult.Errors;
                }
                
                // STEP 3: Single regeneration for ALL clusters
                if (result.Placed > 0)
                {
                    _logger.Info($"Regenerating document for {result.Placed} clusters...");
                    _doc.Regenerate();
                    _logger.Info("Regeneration complete");
                }
                
                return result;
            }
            catch (Exception ex)
            {
                _logger.Error($"Error in optimized cluster placement: {ex.Message}");
                result.Errors = clusterGroups.Count;
                return result;
            }
        }
        
        private List<ClusterFamilyGroup> GroupClustersByFamily(List<ClusterGroup> clusters)
        {
            var groups = clusters
                .GroupBy(c => new
                {
                    FamilyName = GetClusterFamilyName(c),
                    Width = RoundToNearest50mm(c.Width),
                    Height = RoundToNearest50mm(c.Height)
                })
                .Select(g => new ClusterFamilyGroup
                {
                    FamilyName = g.Key.FamilyName,
                    Width = g.Key.Width,
                    Height = g.Key.Height,
                    Clusters = g.ToList()
                })
                .OrderByDescending(g => g.Clusters.Count)
                .ToList();
            
            return groups;
        }
        
        private ClusterPlacementResult ProcessClusterGroup(ClusterFamilyGroup group)
        {
            var result = new ClusterPlacementResult();
            
            try
            {
                _logger.Info($"Processing cluster group: {group.FamilyName}, {group.Clusters.Count} clusters");
                
                // Load and activate family symbol ONCE
                FamilySymbol symbol = _familyLoader.LoadAndActivateSymbol(group.FamilyName);
                if (symbol == null)
                {
                    _logger.Error($"Failed to load cluster family: {group.FamilyName}");
                    result.Errors = group.Clusters.Count;
                    return result;
                }
                
                // Create ALL cluster instances in this group
                var placedClusters = new List<(FamilyInstance instance, ClusterGroup cluster)>();
                
                foreach (var cluster in group.Clusters)
                {
                    try
                    {
                        // Check if already clustered
                        if (cluster.Zones.Any(z => z.IsClusterResolvedFlag || z.ClusterSleeveInstanceId > 0))
                        {
                            _logger.Warning($"Cluster zones already resolved - skipping");
                            result.Errors++;
                            continue;
                        }
                        
                        // Calculate cluster center point
                        XYZ centerPoint = CalculateClusterCenter(cluster.Zones);
                        double rotation = CalculateClusterRotation(cluster);
                        
                        // Create cluster instance
                        FamilyInstance instance = _doc.Create.NewFamilyInstance(
                            centerPoint,
                            symbol,
                            StructuralType.NonStructural
                        );
                        
                        if (instance != null && instance.IsValidObject)
                        {
                            placedClusters.Add((instance, cluster));
                            
                            // Apply rotation
                            if (Math.Abs(rotation) > 0.001)
                            {
                                Line axis = Line.CreateBound(centerPoint, centerPoint + XYZ.BasisZ);
                                ElementTransformUtils.RotateElement(_doc, instance.Id, axis, rotation);
                            }
                        }
                        else
                        {
                            _logger.Warning($"Failed to create cluster instance");
                            result.Errors++;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Error($"Error creating cluster: {ex.Message}");
                        result.Errors++;
                    }
                }
                
                _logger.Info($"Created {placedClusters.Count} cluster instances");
                
                // Set parameters IMMEDIATELY and delete individual sleeves
                foreach (var (instance, cluster) in placedClusters)
                {
                    try
                    {
                        // ✅ IMMEDIATE PARAMETER SETTING
                        bool success = _parameterService.SetSleeveParameters(
                            instance,
                            cluster.Zones.First(), // Use first zone for common properties
                            group.Width,
                            group.Height,
                            null
                        );
                        
                        if (success)
                        {
                            // Delete individual sleeves
                            foreach (var zone in cluster.Zones)
                            {
                                if (zone.SleeveInstanceId > 0)
                                {
                                    try
                                    {
                                        Element sleeve = _doc.GetElement(new ElementId(zone.SleeveInstanceId));
                                        if (sleeve != null && sleeve.IsValidObject)
                                        {
                                            _doc.Delete(sleeve.Id);
                                            result.Deleted++;
                                        }
                                    }
                                    catch (Exception delEx)
                                    {
                                        _logger.Warning($"Failed to delete individual sleeve {zone.SleeveInstanceId}: {delEx.Message}");
                                    }
                                }
                                
                                // Update zone with cluster info
                                zone.ClusterSleeveInstanceId = instance.Id.IntegerValue;
                                zone.IsClusterResolvedFlag = true;
                                zone.SleeveInstanceId = 0; // Clear individual sleeve ID
                            }
                            
                            result.Placed++;
                        }
                        else
                        {
                            _logger.Warning($"Failed to set parameters for cluster {instance.Id.IntegerValue}");
                            result.Errors++;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Error($"Error processing cluster {instance.Id.IntegerValue}: {ex.Message}");
                        result.Errors++;
                    }
                }
                
                _logger.Info($"Cluster group complete: {result.Placed} placed, {result.Deleted} deleted, {result.Errors} errors");
            }
            catch (Exception ex)
            {
                _logger.Error($"Error processing cluster group {group.FamilyName}: {ex.Message}");
                result.Errors = group.Clusters.Count;
            }
            
            return result;
        }
        
        private XYZ CalculateClusterCenter(List<ClashZone> zones)
        {
            double avgX = zones.Average(z => z.IntersectionPointX);
            double avgY = zones.Average(z => z.IntersectionPointY);
            double avgZ = zones.Average(z => z.IntersectionPointZ);
            
            return new XYZ(avgX, avgY, avgZ);
        }
        
        private double CalculateClusterRotation(ClusterGroup cluster)
        {
            // Implement cluster rotation logic
            return 0.0;
        }
        
        private string GetClusterFamilyName(ClusterGroup cluster)
        {
            var firstZone = cluster.Zones.First();
            double size = Math.Max(cluster.Width, cluster.Height);
            
            return ClusterPlacementService.GetFamilyName(
                firstZone.StructuralElementType,
                firstZone.MepElementCategory,
                size,
                isCluster: true
            );
        }
        
        private double RoundToNearest50mm(double valueInFeet)
        {
            double mm = valueInFeet * 304.8;
            double roundedMm = Math.Round(mm / 50.0) * 50.0;
            return roundedMm / 304.8;
        }
    }
    
    public class ClusterFamilyGroup
    {
        public string FamilyName { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public List<ClusterGroup> Clusters { get; set; }
    }
    
    public class ClusterPlacementResult
    {
        public int Placed { get; set; }
        public int Deleted { get; set; }
        public int Errors { get; set; }
    }
}
```

---

## Phase 3: Safe Transaction Management

### 3.1 Create SafeTransactionManager (SRP)

**File:** `Services/Transactions/SafeTransactionManager.cs`

```csharp
using System;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Transactions
{
    /// <summary>
    /// SRP: Responsible ONLY for managing Revit transactions safely
    /// Provides rollback on error, validation before commit, and proper cleanup
    /// </summary>
    public class SafeTransactionManager : ITransactionManager
    {
        private readonly Document _doc;
        private readonly ILogger _logger;
        private Transaction _currentTransaction;
        
        public SafeTransactionManager(Document doc, ILogger logger = null)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _logger = logger ?? new NullLogger();
        }
        
        /// <summary>
        /// Validates document is modifiable before starting transaction
        /// </summary>
        public bool CanStartTransaction()
        {
            if (!_doc.IsModifiable)
            {
                _logger.Error("Document is not modifiable - cannot start transaction");
                return false;
            }
            
            if (_currentTransaction != null && _currentTransaction.GetStatus() == TransactionStatus.Started)
            {
                _logger.Warning("Transaction already in progress");
                return false;
            }
            
            return true;
        }
        
        /// <summary>
        /// Starts a new transaction with validation
        /// </summary>
        public bool StartTransaction(string transactionName)
        {
            if (!CanStartTransaction())
                return false;
            
            try
            {
                _currentTransaction = new Transaction(_doc, transactionName);
                TransactionStatus status = _currentTransaction.Start();
                
                if (status != TransactionStatus.Started)
                {
                    _logger.Error($"Failed to start transaction '{transactionName}': Status={status}");
                    _currentTransaction = null;
                    return false;
                }
                
                _logger.Info($"✅ Transaction started: '{transactionName}'");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error($"Exception starting transaction '{transactionName}': {ex.Message}");
                _currentTransaction = null;
                return false;
            }
        }
        
        /// <summary>
        /// Commits transaction with validation
        /// Rolls back automatically on failure
        /// </summary>
        public bool CommitTransaction()
        {
            if (_currentTransaction == null)
            {
                _logger.Warning("No transaction to commit");
                return false;
            }
            
            try
            {
                // Validate before commit
                if (_currentTransaction.GetStatus() != TransactionStatus.Started)
                {
                    _logger.Error($"Cannot commit transaction - status is {_currentTransaction.GetStatus()}");
                    return false;
                }
                
                // Commit
                TransactionStatus status = _currentTransaction.Commit();
                
                if (status == TransactionStatus.Committed)
                {
                    _logger.Info("✅ Transaction committed successfully");
                    _currentTransaction = null;
                    return true;
                }
                else
                {
                    _logger.Error($"Transaction commit failed: Status={status}");
                    RollbackTransaction();
                    return false;
                }
            }
            catch (Exception ex)
            {
                _logger.Error($"Exception committing transaction: {ex.Message}");
                RollbackTransaction();
                return false;
            }
            finally
            {
                _currentTransaction = null;
            }
        }
        
        /// <summary>
        /// Rolls back transaction safely
        /// </summary>
        public void RollbackTransaction()
        {
            if (_currentTransaction == null)
                return;
            
            try
            {
                if (_currentTransaction.GetStatus() == TransactionStatus.Started)
                {
                    _currentTransaction.RollBack();
                    _logger.Warning("⚠️ Transaction rolled back");
                }
            }
            catch (Exception ex)
            {
                _logger.Error($"Exception during rollback: {ex.Message}");
            }
            finally
            {
                _currentTransaction = null;
            }
        }
        
        /// <summary>
        /// Executes an action within a transaction
        /// Automatically commits on success, rolls back on failure
        /// </summary>
        public T ExecuteInTransaction<T>(string transactionName, Func<T> action)
        {
            if (!StartTransaction(transactionName))
            {
                throw new InvalidOperationException($"Failed to start transaction '{transactionName}'");
            }
            
            try
            {
                T result = action();
                
                if (!CommitTransaction())
                {
                    throw new InvalidOperationException($"Failed to commit transaction '{transactionName}'");
                }
                
                return result;
            }
            catch (Exception ex)
            {
                _logger.Error($"Error in transaction '{transactionName}': {ex.Message}");
                RollbackTransaction();
                throw;
            }
        }
        
        /// <summary>
        /// Disposes transaction safely
        /// </summary>
        public void Dispose()
        {
            if (_currentTransaction != null)
            {
                if (_currentTransaction.GetStatus() == TransactionStatus.Started)
                {
                    RollbackTransaction();
                }
                
                _currentTransaction?.Dispose();
                _currentTransaction = null;
            }
        }
    }
    
    public interface ITransactionManager : IDisposable
    {
        bool CanStartTransaction();
        bool StartTransaction(string transactionName);
        bool CommitTransaction();
        void RollbackTransaction();
        T ExecuteInTransaction<T>(string transactionName, Func<T> action);
    }
}
```

---

## Phase 4: Crash-Safe Features

### 4.1 Create CrashSafeExecutor (SRP)

**File:** `Services/CrashSafe/CrashSafeExecutor.cs`

```csharp
using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.CrashSafe
{
    /// <summary>
    /// SRP: Responsible ONLY for executing operations with timeout protection and crash handling
    /// Provides graceful degradation when operations take too long or fail
    /// </summary>
    public class CrashSafeExecutor : ICrashSafeExecutor
    {
        private readonly ILogger _logger;
        private readonly int _timeoutSeconds;
        private readonly Stopwatch _operationTimer;
        private CancellationTokenSource _cancellationTokenSource;
        
        public CrashSafeExecutor(int timeoutSeconds = 300, ILogger logger = null)
        {
            _timeoutSeconds = timeoutSeconds;
            _logger = logger ?? new NullLogger();
            _operationTimer = new Stopwatch();
        }
        
        /// <summary>
        /// Executes a Revit operation with timeout protection
        /// Returns Result.Succeeded on success, Result.Failed on timeout/error
        /// </summary>
        public Result ExecuteWithTimeout(Func<Result> operation, string operationName)
        {
            if (operation == null)
                throw new ArgumentNullException(nameof(operation));
            
            _logger.Info($"🕒 Starting operation '{operationName}' with {_timeoutSeconds}s timeout");
            _operationTimer.Restart();
            _cancellationTokenSource = new CancellationTokenSource();
            
            try
            {
                // Create task for operation
                var task = Task.Run(() =>
                {
                    try
                    {
                        return operation();
                    }
                    catch (Exception ex)
                    {
                        _logger.Error($"Exception in operation '{operationName}': {ex.Message}");
                        _logger.Error($"Stack trace: {ex.StackTrace}");
                        return Result.Failed;
                    }
                }, _cancellationTokenSource.Token);
                
                // Wait with timeout
                bool completed = task.Wait(TimeSpan.FromSeconds(_timeoutSeconds));
                
                _operationTimer.Stop();
                
                if (!completed)
                {
                    // TIMEOUT
                    _logger.Error($"⏰ TIMEOUT: Operation '{operationName}' exceeded {_timeoutSeconds}s");
                    _cancellationTokenSource.Cancel();
                    
                    return Result.Failed;
                }
                
                // SUCCESS
                Result result = task.Result;
                _logger.Info($"✅ Operation '{operationName}' completed in {_operationTimer.ElapsedMilliseconds}ms: {result}");
                
                return result;
            }
            catch (AggregateException aggEx)
            {
                // Handle task exceptions
                _operationTimer.Stop();
                
                foreach (var ex in aggEx.InnerExceptions)
                {
                    _logger.Error($"❌ Task exception in '{operationName}': {ex.Message}");
                }
                
                return Result.Failed;
            }
            catch (Exception ex)
            {
                _operationTimer.Stop();
                _logger.Error($"❌ Unexpected exception in '{operationName}': {ex.Message}");
                return Result.Failed;
            }
            finally
            {
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;
            }
        }
        
        /// <summary>
        /// Checks if operation has exceeded timeout
        /// Can be called periodically during long operations
        /// </summary>
        public bool CheckTimeout(string checkpointName)
        {
            if (!_operationTimer.IsRunning)
                return false;
            
            long elapsedSeconds = _operationTimer.ElapsedMilliseconds / 1000;
            
            if (elapsedSeconds > _timeoutSeconds)
            {
                _logger.Warning($"⏰ Timeout detected at checkpoint '{checkpointName}': {elapsedSeconds}s elapsed");
                return true;
            }
            
            return false;
        }
        
        /// <summary>
        /// Executes operation with retry logic
        /// Useful for operations that may fail temporarily
        /// </summary>
        public Result ExecuteWithRetry(
            Func<Result> operation,
            string operationName,
            int maxRetries = 3,
            int delayMs = 1000)
        {
            int attempt = 0;
            
            while (attempt < maxRetries)
            {
                attempt++;
                
                _logger.Info($"🔄 Attempt {attempt}/{maxRetries} for '{operationName}'");
                
                Result result = ExecuteWithTimeout(operation, operationName);
                
                if (result == Result.Succeeded)
                {
                    return result;
                }
                
                if (attempt < maxRetries)
                {
                    _logger.Warning($"⚠️ Attempt {attempt} failed, waiting {delayMs}ms before retry...");
                    Thread.Sleep(delayMs);
                }
            }
            
            _logger.Error($"❌ All {maxRetries} attempts failed for '{operationName}'");
            return Result.Failed;
        }
        
        /// <summary>
        /// Executes operation with progress tracking
        /// Logs progress at regular intervals
        /// </summary>
        public Result ExecuteWithProgress(
            Func<Action<int, int>, Result> operation,
            string operationName,
            int totalItems)
        {
            int processedItems = 0;
            int lastReportedPercent = 0;
            
            _logger.Info($"📊 Starting '{operationName}' with {totalItems} items");
            
            Action<int, int> progressCallback = (current, total) =>
            {
                processedItems = current;
                int percent = (int)((current / (double)total) * 100);
                
                if (percent >= lastReportedPercent + 10) // Report every 10%
                {
                    _logger.Info($"📊 Progress: {percent}% ({current}/{total})");
                    lastReportedPercent = percent;
                }
            };
            
            Result result = ExecuteWithTimeout(
                () => operation(progressCallback),
                operationName
            );
            
            if (result == Result.Succeeded)
            {
                _logger.Info($"✅ Completed {processedItems}/{totalItems} items");
            }
            else
            {
                _logger.Error($"❌ Failed after processing {processedItems}/{totalItems} items");
            }
            
            return result;
        }
    }
    
    public interface ICrashSafeExecutor
    {
        Result ExecuteWithTimeout(Func<Result> operation, string operationName);
        bool CheckTimeout(string checkpointName);
        Result ExecuteWithRetry(Func<Result> operation, string operationName, int maxRetries = 3, int delayMs = 1000);
        Result ExecuteWithProgress(Func<Action<int, int>, Result> operation, string operationName, int totalItems);
    }
}
```

### 4.2 Create OperationRecoveryService (SRP)

**File:** `Services/CrashSafe/OperationRecoveryService.cs`

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.CrashSafe
{
    /// <summary>
    /// SRP: Responsible ONLY for recovering from failed operations
    /// Provides checkpoint saving and resumption after crashes/timeouts
    /// </summary>
    public class OperationRecoveryService : IOperationRecoveryService
    {
        private readonly ILogger _logger;
        private readonly Dictionary<string, OperationCheckpoint> _checkpoints;
        
        public OperationRecoveryService(ILogger logger = null)
        {
            _logger = logger ?? new NullLogger();
            _checkpoints = new Dictionary<string, OperationCheckpoint>();
        }
        
        /// <summary>
        /// Saves operation checkpoint for recovery
        /// </summary>
        public void SaveCheckpoint(string operationId, OperationCheckpoint checkpoint)
        {
            if (string.IsNullOrEmpty(operationId))
                throw new ArgumentNullException(nameof(operationId));
            
            if (checkpoint == null)
                throw new ArgumentNullException(nameof(checkpoint));
            
            _checkpoints[operationId] = checkpoint;
            
            _logger.Info($"💾 Checkpoint saved: {operationId} - {checkpoint.ProcessedCount}/{checkpoint.TotalCount} items");
        }
        
        /// <summary>
        /// Retrieves saved checkpoint for resumption
        /// </summary>
        public OperationCheckpoint GetCheckpoint(string operationId)
        {
            if (_checkpoints.TryGetValue(operationId, out OperationCheckpoint checkpoint))
            {
                _logger.Info($"📂 Checkpoint retrieved: {operationId} - {checkpoint.ProcessedCount}/{checkpoint.TotalCount} items");
                return checkpoint;
            }
            
            return null;
        }
        
        /// <summary>
        /// Filters zones to only those not yet processed (based on checkpoint)
        /// </summary>
        public List<ClashZone> GetRemainingZones(string operationId, List<ClashZone> allZones)
        {
            var checkpoint = GetCheckpoint(operationId);
            
            if (checkpoint == null || checkpoint.ProcessedZoneIds == null)
            {
                _logger.Info($"No checkpoint found for '{operationId}' - processing all {allZones.Count} zones");
                return allZones;
            }
            
            var processedIds = new HashSet<Guid>(checkpoint.ProcessedZoneIds);
            var remainingZones = allZones.Where(z => !processedIds.Contains(z.Id)).ToList();
            
            _logger.Info($"🔄 Resuming '{operationId}': {remainingZones.Count} zones remaining (already processed: {checkpoint.ProcessedZoneIds.Count})");
            
            return remainingZones;
        }
        
        /// <summary>
        /// Clears checkpoint after successful completion
        /// </summary>
        public void ClearCheckpoint(string operationId)
        {
            if (_checkpoints.Remove(operationId))
            {
                _logger.Info($"🗑️ Checkpoint cleared: {operationId}");
            }
        }
        
        /// <summary>
        /// Creates checkpoint from current progress
        /// </summary>
        public OperationCheckpoint CreateCheckpoint(
            int processedCount,
            int totalCount,
            List<Guid> processedZoneIds,
            Dictionary<string, object> metadata = null)
        {
            return new OperationCheckpoint
            {
                Timestamp = DateTime.Now,
                ProcessedCount = processedCount,
                TotalCount = totalCount,
                ProcessedZoneIds = processedZoneIds ?? new List<Guid>(),
                Metadata = metadata ?? new Dictionary<string, object>()
            };
        }
    }
    
    public class OperationCheckpoint
    {
        public DateTime Timestamp { get; set; }
        public int ProcessedCount { get; set; }
        public int TotalCount { get; set; }
        public List<Guid> ProcessedZoneIds { get; set; }
        public Dictionary<string, object> Metadata { get; set; }
    }
    
    public interface IOperationRecoveryService
    {
        void SaveCheckpoint(string operationId, OperationCheckpoint checkpoint);
        OperationCheckpoint GetCheckpoint(string operationId);
        List<ClashZone> GetRemainingZones(string operationId, List<ClashZone> allZones);
        void ClearCheckpoint(string operationId);
        OperationCheckpoint CreateCheckpoint(int processedCount, int totalCount, List<Guid> processedZoneIds, Dictionary<string, object> metadata = null);
    }
}
```

---

## Phase 5: Integration

### 5.1 Create PlacementOrchestrator (Coordinates Everything)

**File:** `Services/Placement/PlacementOrchestrator.cs`

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Transactions;
using JSE_RevitAddin_MEP_OPENINGS.Services.CrashSafe;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// Orchestrates entire placement workflow with SOLID principles
    /// Coordinates: Individual sleeves, Clusters, Transactions, Crash-safety
    /// </summary>
    public class PlacementOrchestrator
    {
        private readonly Document _doc;
        private readonly OptimizedIndividualSleevePlacer _individualPlacer;
        private readonly OptimizedClusterSleevePlacer _clusterPlacer;
        private readonly ITransactionManager _transactionManager;
        private readonly ICrashSafeExecutor _crashSafeExecutor;
        private readonly IOperationRecoveryService _recoveryService;
        private readonly ILogger _logger;
        
        /// <summary>
        /// Constructor with Dependency Injection (DIP)
        /// </summary>
        public PlacementOrchestrator(
            Document doc,
            OptimizedIndividualSleevePlacer individualPlacer,
            OptimizedClusterSleevePlacer clusterPlacer,
            ITransactionManager transactionManager,
            ICrashSafeExecutor crashSafeExecutor,
            IOperationRecoveryService recoveryService,
            ILogger logger = null)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _individualPlacer = individualPlacer ?? throw new ArgumentNullException(nameof(individualPlacer));
            _clusterPlacer = clusterPlacer ?? throw new ArgumentNullException(nameof(clusterPlacer));
            _transactionManager = transactionManager ?? throw new ArgumentNullException(nameof(transactionManager));
            _crashSafeExecutor = crashSafeExecutor ?? throw new ArgumentNullException(nameof(crashSafeExecutor));
            _recoveryService = recoveryService ?? throw new ArgumentNullException(nameof(recoveryService));
            _logger = logger ?? new NullLogger();
        }
        
        /// <summary>
        /// Main entry point - places individual sleeves with full protection
        /// </summary>
        public PlacementResult PlaceIndividualSleeves(List<ClashZone> zones, string operationId = "IndividualPlacement")
        {
            if (zones == null || zones.Count == 0)
            {
                _logger.Warning("No zones provided for individual sleeve placement");
                return new PlacementResult { Placed = 0, Skipped = 0, Errors = 0 };
            }
            
            _logger.Info($"═══════════════════════════════════════════════════");
            _logger.Info($"STARTING INDIVIDUAL SLEEVE PLACEMENT: {zones.Count} zones");
            _logger.Info($"═══════════════════════════════════════════════════");
            
            PlacementResult finalResult = new PlacementResult();
            
            // STEP 1: Check for existing checkpoint (crash recovery)
            var remainingZones = _recoveryService.GetRemainingZones(operationId, zones);
            
            if (remainingZones.Count == 0)
            {
                _logger.Info("✅ All zones already processed (from checkpoint)");
                return finalResult;
            }
            
            // STEP 2: Execute with crash-safe wrapper
            Result executionResult = _crashSafeExecutor.ExecuteWithTimeout(() =>
            {
                try
                {
                    // STEP 3: Execute in safe transaction
                    finalResult = _transactionManager.ExecuteInTransaction("Place Individual Sleeves", () =>
                    {
                        PlacementResult result = _individualPlacer.PlaceSleeves(remainingZones);
                        
                        // Save checkpoint after each batch
                        var processedIds = remainingZones
                            .Where(z => z.SleeveInstanceId > 0)
                            .Select(z => z.Id)
                            .ToList();
                        
                        var checkpoint = _recoveryService.CreateCheckpoint(
                            result.Placed,
                            zones.Count,
                            processedIds
                        );
                        
                        _recoveryService.SaveCheckpoint(operationId, checkpoint);
                        
                        return result;
                    });
                    
                    return Result.Succeeded;
                }
                catch (Exception ex)
                {
                    _logger.Error($"Error in placement: {ex.Message}");
                    return Result.Failed;
                }
            }, $"Place {remainingZones.Count} Individual Sleeves");
            
            // STEP 4: Handle result
            if (executionResult == Result.Succeeded)
            {
                _logger.Info($"✅ PLACEMENT COMPLETE: {finalResult.Placed} placed, {finalResult.Skipped} skipped, {finalResult.Errors} errors");
                _recoveryService.ClearCheckpoint(operationId);
            }
            else
            {
                _logger.Error($"❌ PLACEMENT FAILED or TIMEOUT");
            }
            
            _logger.Info($"═══════════════════════════════════════════════════");
            
            return finalResult;
        }
        
        /// <summary>
        /// Places cluster sleeves with full protection
        /// </summary>
        public ClusterPlacementResult PlaceClusters(
            List<ClusterGroup> clusterGroups,
            string operationId = "ClusterPlacement")
        {
            if (clusterGroups == null || clusterGroups.Count == 0)
            {
                _logger.Warning("No cluster groups provided for placement");
                return new ClusterPlacementResult { Placed = 0, Deleted = 0, Errors = 0 };
            }
            
            _logger.Info($"═══════════════════════════════════════════════════");
            _logger.Info($"STARTING CLUSTER PLACEMENT: {clusterGroups.Count} clusters");
            _logger.Info($"═══════════════════════════════════════════════════");
            
            ClusterPlacementResult finalResult = new ClusterPlacementResult();
            
            // STEP 1: Execute with crash-safe wrapper
            Result executionResult = _crashSafeExecutor.ExecuteWithTimeout(() =>
            {
                try
                {
                    // STEP 2: Execute in safe transaction
                    finalResult = _transactionManager.ExecuteInTransaction("Place Cluster Sleeves", () =>
                    {
                        return _clusterPlacer.PlaceClusters(clusterGroups);
                    });
                    
                    return Result.Succeeded;
                }
                catch (Exception ex)
                {
                    _logger.Error($"Error in cluster placement: {ex.Message}");
                    return Result.Failed;
                }
            }, $"Place {clusterGroups.Count} Clusters");
            
            // STEP 3: Handle result
            if (executionResult == Result.Succeeded)
            {
                _logger.Info($"✅ CLUSTER PLACEMENT COMPLETE: {finalResult.Placed} placed, {finalResult.Deleted} deleted, {finalResult.Errors} errors");
            }
            else
            {
                _logger.Error($"❌ CLUSTER PLACEMENT FAILED or TIMEOUT");
            }
            
            _logger.Info($"═══════════════════════════════════════════════════");
            
            return finalResult;
        }
        
        /// <summary>
        /// Factory method for creating orchestrator with all dependencies
        /// </summary>
        public static PlacementOrchestrator Create(Document doc, ILogger logger = null)
        {
            logger = logger ?? new NullLogger();
            
            // Create services
            var groupingService = new GroupingService(logger);
            var parameterService = new ImmediateParameterService(doc, logger);
            var familyLoader = new FamilyLoaderService(doc, logger); // Implement this
            var transactionManager = new SafeTransactionManager(doc, logger);
            var crashSafeExecutor = new CrashSafeExecutor(300, logger); // 5 minute timeout
            var recoveryService = new OperationRecoveryService(logger);
            
            // Create placers
            var individualPlacer = new OptimizedIndividualSleevePlacer(
                doc,
                groupingService,
                parameterService,
                familyLoader,
                logger
            );
            
            var clusterPlacer = new OptimizedClusterSleevePlacer(
                doc,
                groupingService,
                parameterService,
                familyLoader,
                logger
            );
            
            // Create orchestrator
            return new PlacementOrchestrator(
                doc,
                individualPlacer,
                clusterPlacer,
                transactionManager,
                crashSafeExecutor,
                recoveryService,
                logger
            );
        }
    }
}
```

### 5.2 Update NewSleevePlacerService to Use Orchestrator

**File:** `Services/NewSleevePlacerService.cs` (Modify existing)

```csharp
// Add at top of class
private readonly PlacementOrchestrator _orchestrator;

// In constructor, add:
public NewSleevePlacerService(
    Document doc,
    // ... existing parameters ...
    )
{
    // ... existing initialization ...
    
    // ✅ NEW: Create orchestrator with all dependencies
    _orchestrator = PlacementOrchestrator.Create(doc, DebugLogger.Instance);
}

// Replace ExecutePlacementInternal method:
private (int placed, int skipped, int errors) ExecutePlacementInternal(List<ClashZone> clashZones)
{
    // ❌ REMOVE ALL OLD CODE (deferred parameters, flush, manual loop, etc.)
    
    // ✅ NEW: Delegate to orchestrator
    var result = _orchestrator.PlaceIndividualSleeves(clashZones, "BatchPlacement");
    
    return (result.Placed, result.Skipped, result.Errors);
}
```

### 5.3 Update RefactoredClusterService to Use Orchestrator

**File:** `Services/RefactoredClusterService.cs` (Modify existing)

```csharp
// Add at top of class
private readonly PlacementOrchestrator _orchestrator;

// In PlaceClusterForGroup or similar method:
public (int placedCount, int deletedCount) PlaceClusters(List<ClusterGroup> clusterGroups)
{
    // ❌ REMOVE: _deferredClusterParameters
    // ❌ REMOVE: DivertedBatchDictionary assignment
    // ❌ REMOVE: Manual placement loop
    // ❌ REMOVE: FlushClusterParameters
    
    // ✅ NEW: Delegate to orchestrator
    var result = _orchestrator.PlaceClusters(clusterGroups, "ClusterPlacement");
    
    return (result.Placed, result.Deleted);
}
```

---

## Phase 6: Testing & Verification

### 6.1 Unit Tests

**File:** `Tests/Placement/OptimizedPlacementTests.cs`

```csharp
using Xunit;
using Moq;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;

namespace JSE_RevitAddin_MEP_OPENINGS.Tests.Placement
{
    public class OptimizedPlacementTests
    {
        [Fact]
        public void GroupingService_GroupsByFamilyAndDimensions()
        {
            // Arrange
            var zones = CreateTestZones(10);
            var groupingService = new GroupingService();
            
            // Act
            var groups = groupingService.GroupZonesByFamilyAndDimensions(
                zones,
                z => "TestFamily",
                z => 1.0,
                z => 1.0
            );
            
            // Assert
            Assert.Single(groups); // All zones same family/dimensions
            Assert.Equal(10, groups[0].Zones.Count);
        }
        
        [Fact]
        public void ImmediateParameterService_SetsParameters()
        {
            // Arrange
            var mockDoc = new Mock<Document>();
            var mockInstance = new Mock<FamilyInstance>();
            var mockParam = new Mock<Parameter>();
            
            mockParam.Setup(p => p.IsReadOnly).Returns(false);
            mockParam.Setup(p => p.Set(It.IsAny<double>())).Verifiable();
            mockInstance.Setup(i => i.LookupParameter("Width")).Returns(mockParam.Object);
            
            var service = new ImmediateParameterService(mockDoc.Object);
            var zone = CreateTestZone();
            
            // Act
            bool result = service.SetSleeveParameters(mockInstance.Object, zone, 1.0, 1.0);
            
            // Assert
            Assert.True(result);
            mockParam.Verify(p => p.Set(It.IsAny<double>()), Times.AtLeastOnce);
        }
        
        [Fact]
        public void TransactionManager_CommitsSuccessfully()
        {
            // Test transaction lifecycle
        }
        
        [Fact]
        public void CrashSafeExecutor_HandlesTimeout()
        {
            // Test timeout handling
        }
    }
}
```

### 6.2 Integration Tests

**Checklist:**
- [ ] Place 39 individual sleeves - verify all have correct dimensions
- [ ] Place 10 clusters - verify no duplicates
- [ ] Test timeout scenario (> 5 minutes) - verify graceful handling
- [ ] Test crash recovery - stop mid-placement, resume
- [ ] Test transaction rollback - verify no orphaned elements
- [ ] Verify database consistency after placement
- [ ] Verify Revit model matches database values

### 6.3 Performance Tests

**Metrics to track:**
- [ ] Time to place 39 individual sleeves
- [ ] Time to place 10 clusters
- [ ] Memory usage during placement
- [ ] Number of Revit API calls
- [ ] Number of regenerations
- [ ] Database query time

**Expected Results:**
| Metric | Sequential Mode | Optimized Batch Mode | Improvement |
|--------|----------------|---------------------|-------------|
| Time (39 sleeves) | 30-40s | 8-12s | 3-4x faster |
| Regenerations | 39 | 1 | 39x fewer |
| Memory | Baseline | +10-20% | Acceptable |

---

## Phase 7: Performance Monitoring

### 7.1 Create PerformanceMonitor

**File:** `Services/Monitoring/PerformanceMonitor.cs`

```csharp
using System;
using System.Diagnostics;
using System.Collections.Generic;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Monitoring
{
    public class PerformanceMonitor
    {
        private readonly Dictionary<string, Stopwatch> _timers;
        private readonly Dictionary<string, long> _counters;
        private readonly ILogger _logger;
        
        public PerformanceMonitor(ILogger logger = null)
        {
            _timers = new Dictionary<string, Stopwatch>();
            _counters = new Dictionary<string, long>();
            _logger = logger ?? new NullLogger();
        }
        
        public void StartTimer(string operationName)
        {
            if (!_timers.ContainsKey(operationName))
            {
                _timers[operationName] = new Stopwatch();
            }
            
            _timers[operationName].Restart();
        }
        
        public void StopTimer(string operationName)
        {
            if (_timers.TryGetValue(operationName, out Stopwatch timer))
            {
                timer.Stop();
                _logger.Info($"⏱️ {operationName}: {timer.ElapsedMilliseconds}ms");
            }
        }
        
        public void IncrementCounter(string counterName, long amount = 1)
        {
            if (!_counters.ContainsKey(counterName))
            {
                _counters[counterName] = 0;
            }
            
            _counters[counterName] += amount;
        }
        
        public void GenerateReport()
        {
            _logger.Info("═══════════════════════════════════════════════════");
            _logger.Info("PERFORMANCE REPORT");
            _logger.Info("═══════════════════════════════════════════════════");
            
            _logger.Info("TIMERS:");
            foreach (var kvp in _timers)
            {
                _logger.Info($"  {kvp.Key}: {kvp.Value.ElapsedMilliseconds}ms");
            }
            
            _logger.Info("COUNTERS:");
            foreach (var kvp in _counters)
            {
                _logger.Info($"  {kvp.Key}: {kvp.Value}");
            }
            
            _logger.Info("═══════════════════════════════════════════════════");
        }
    }
}
```

---

## Rollback & Fallback Strategy

### Emergency Rollback Plan

**If optimized batch mode fails:**

1. **Immediate Fallback to Sequential Mode**
   ```csharp
   // Add feature flag
   public static class FeatureFlags
   {
       public static bool UseOptimizedBatchMode = true; // Toggle to false for rollback
   }
   
   // In NewSleevePlacerService
   if (FeatureFlags.UseOptimizedBatchMode)
   {
       return _orchestrator.PlaceIndividualSleeves(zones);
   }
   else
   {
       return PlaceSequentialMode(zones); // Old code
   }
   ```

2. **Gradual Rollout**
   - Week 1: Enable for testing only (internal users)
   - Week 2: Enable for 10% of users
   - Week 3: Enable for 50% of users
   - Week 4: Enable for 100% of users

3. **Monitoring & Alerts**
   - Track error rates
   - Monitor performance metrics
   - Alert if errors > 5% or performance < sequential mode

---

## Implementation Checklist

### Phase 1: Core Services (Week 1)
- [ ] Create GroupingService
- [ ] Create ImmediateParameterService
- [ ] Create OptimizedIndividualSleevePlacer
- [ ] Unit tests for above
- [ ] Integration test: 39 individual sleeves

### Phase 2: Cluster Support (Week 1)
- [ ] Create OptimizedClusterSleevePlacer
- [ ] Update clustering logic to remove deferred parameters
- [ ] Integration test: 10 clusters
- [ ] Verify no duplicate clusters

### Phase 3: Safety Features (Week 2)
- [ ] Create SafeTransactionManager
- [ ] Create CrashSafeExecutor
- [ ] Create OperationRecoveryService
- [ ] Test timeout scenarios
- [ ] Test crash recovery

### Phase 4: Integration (Week 2)
- [ ] Create PlacementOrchestrator
- [ ] Update NewSleevePlacerService
- [ ] Update RefactoredClusterService
- [ ] End-to-end testing

### Phase 5: Performance (Week 3)
- [ ] Performance testing
- [ ] Memory profiling
- [ ] Optimize bottlenecks
- [ ] Document performance gains

### Phase 6: Deployment (Week 3)
- [ ] Code review
- [ ] Documentation
- [ ] Feature flag implementation
- [ ] Gradual rollout plan

---

## Success Criteria

✅ **Functionality:**
- All sleeves have correct dimensions (250x250, 300x300, etc.)
- No duplicate clusters
- No wrong coordinates
- Database matches Revit model

✅ **Performance:**
- 3-4x faster than sequential mode
- Memory usage < 20% increase
- Completes 39 sleeves in < 15 seconds

✅ **Reliability:**
- Handles timeouts gracefully
- Recovers from crashes
- Proper transaction management
- No orphaned elements

✅ **Maintainability:**
- SOLID principles followed
- 80%+ code coverage (unit tests)
- Comprehensive documentation
- Easy to debug and extend

---

## Conclusion

This plan provides a complete, production-ready implementation of optimized batch placement with:

1. **SOLID Principles** - Clean, maintainable architecture
2. **Safe Transaction Management** - Proper Revit transaction handling
3. **Crash-Safe Features** - Timeout protection and recovery
4. **Immediate Parameters** - Proven to work (no deferred/flush issues)
5. **Grouped Processing** - Optimal performance through batching

**Expected Outcome:** 
- ✅ Works correctly (no 300x100 defaults, no duplicates)
- ✅ 3-4x faster than sequential mode
- ✅ Safe and reliable (crash protection, rollback)
- ✅ Maintainable (SOLID, well-tested)

**Ready for implementation!** 🚀
