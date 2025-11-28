# Team D: Persistence & Error/Logging Framework - Usage Examples

## Overview
Team D has created persistence and error/logging abstractions that enable centralized error handling and structured logging while maintaining full backward compatibility with existing infrastructure.

## Key Deliverables

### 1. ISleevePersistenceService
Encapsulates all database persistence operations:
- `PersistPlacementAsync` - Batch persistence of placed sleeves
- `UpdateInstanceAsync` - Individual sleeve instance updates
- `UpdateBoundingBoxAsync` - Bounding box coordinate updates
- `SaveSnapshotsAsync` - Parameter snapshot persistence

### 2. IErrorPolicy with FailSafeErrorPolicy
Centralizes error handling decisions:
- Determines log level based on exception type
- Decides whether to continue processing
- Aggregates errors for batch reporting
- Fail-safe: Continues on non-critical errors

### 3. ILogger Abstraction
Wraps DebugLogger + SafeFileLogger:
- Preserves existing log file names
- Preserves existing message prefixes
- Adds correlation ID support
- Structured logging capability

## Usage Examples

### Example 1: Using ISleevePersistenceService

```csharp
public class MyPlacementService
{
    private readonly ISleevePersistenceService _persistence;
    private readonly ILogger _logger;
    
    public MyPlacementService(
        ISleevePersistenceService persistence = null,
        ILogger logger = null)
    {
        // ✅ COEXISTENCE: Use adapter if not injected
        _persistence = persistence ?? new SleevePersistenceServiceAdapter();
        _logger = logger ?? LoggerAdapter.Default;
    }
    
    public async Task PlaceSleeves(List<ClashZone> zones)
    {
        var placedSleeves = new List<(FamilyInstance, ClashZone)>();
        
        foreach (var zone in zones)
        {
            try
            {
                var instance = PlaceSleeve(zone);
                if (instance != null)
                {
                    placedSleeves.Add((instance, zone));
                }
            }
            catch (Exception ex)
            {
                _logger.Error($"Failed to place sleeve for zone {zone.Id}", ex, "PlaceSleeves");
                // Continue with next zone (fail-safe)
            }
        }
        
        // ✅ PERSISTENCE: All persistence operations route through abstraction
        int persisted = await _persistence.PersistPlacementAsync(
            placedSleeves,
            filterName: "MyFilter",
            category: "Ducts"
        );
        
        _logger.Info($"Persisted {persisted} of {placedSleeves.Count} sleeves", "PlaceSleeves");
    }
}
```

### Example 2: Using IErrorPolicy

```csharp
public class MyPlacementOrchestrator
{
    private readonly IErrorPolicy _errorPolicy;
    private readonly ILogger _logger;
    
    public MyPlacementOrchestrator(
        IErrorPolicy errorPolicy = null,
        ILogger logger = null)
    {
        _errorPolicy = errorPolicy ?? new FailSafeErrorPolicy(logger);
        _logger = logger ?? LoggerAdapter.Default;
    }
    
    public void ExecuteStages(List<IPlacementStage> stages, PlacementContext context)
    {
        var errors = new List<ErrorHandlingResult>();
        
        foreach (var stage in stages)
        {
            try
            {
                stage.Execute(context);
            }
            catch (Exception ex)
            {
                // ✅ ERROR POLICY: Centralized error handling
                var result = _errorPolicy.Handle(
                    ex,
                    scope: stage.Name,
                    context: context,
                    critical: false // Non-critical - continue with next stage
                );
                
                errors.Add(result);
                
                if (!result.ShouldContinue)
                {
                    _logger.Error($"Critical error in {stage.Name}, aborting", ex, "Orchestrator");
                    break;
                }
            }
        }
        
        // ✅ ERROR AGGREGATION: Get summary of all errors
        var summary = _errorPolicy.AggregateErrors(errors);
        if (summary.TotalErrors > 0)
        {
            _logger.Warning($"Completed with {summary.TotalErrors} errors ({summary.CriticalErrors} critical)", "Orchestrator");
        }
    }
}
```

### Example 3: Using ILogger with Correlation ID

```csharp
public class MyService
{
    private readonly ILogger _logger;
    
    public MyService(ILogger logger = null)
    {
        _logger = logger ?? new LoggerAdapter(correlationId: Guid.NewGuid().ToString("N").Substring(0, 8));
    }
    
    public void ProcessZones(List<ClashZone> zones)
    {
        // ✅ STRUCTURED LOGGING: Log with correlation ID
        _logger.Info($"Processing {zones.Count} zones", "ProcessZones");
        
        foreach (var zone in zones)
        {
            try
            {
                ProcessZone(zone);
                
                // ✅ STRUCTURED LOGGING: Log with structured data
                _logger.LogStructured(
                    "INFO",
                    "Zone processed",
                    new { zoneId = zone.Id, sleeveId = zone.SleeveInstanceId },
                    "ProcessZones"
                );
            }
            catch (Exception ex)
            {
                _logger.Error($"Failed to process zone {zone.Id}", ex, "ProcessZones");
            }
        }
    }
}
```

### Example 4: Custom Error Policy

```csharp
public class StrictErrorPolicy : IErrorPolicy
{
    private readonly ILogger _logger;
    
    public StrictErrorPolicy(ILogger logger)
    {
        _logger = logger;
    }
    
    public ErrorHandlingResult Handle(Exception exception, string scope, object context = null, bool critical = false)
    {
        // Strict policy: Always abort on any error
        return new ErrorHandlingResult
        {
            ShouldContinue = false, // Always stop
            LogLevel = "Error",
            Message = $"Strict policy: {exception.Message}",
            Scope = scope,
            Exception = exception
        };
    }
    
    public ErrorSummary AggregateErrors(IEnumerable<ErrorHandlingResult> errors)
    {
        // Strict: Any error means abort
        var summary = new ErrorSummary();
        var errorList = errors?.ToList() ?? new List<ErrorHandlingResult>();
        summary.TotalErrors = errorList.Count;
        summary.ShouldAbort = errorList.Count > 0; // Abort on any error
        return summary;
    }
}
```

### Example 5: Integration with PlacementContext

```csharp
// In PlacementContext (Team A will create this)
public class PlacementContext
{
    public ILogger Logger { get; set; }
    public IErrorPolicy ErrorPolicy { get; set; }
    public ISleevePersistenceService Persistence { get; set; }
    
    // ... other properties
}

// Usage in stage
public class PersistenceStage : IPlacementStage
{
    public string Name => "Persistence";
    
    public void Execute(PlacementContext context)
    {
        context.Logger.Info("Starting persistence stage", Name);
        
        try
        {
            var placedSleeves = context.PlacedInstances
                .Select(inst => (inst, context.GetZoneForInstance(inst)))
                .ToList();
            
            var persisted = context.Persistence
                .PersistPlacementAsync(placedSleeves, context.FilterName, context.Category)
                .GetAwaiter()
                .GetResult();
            
            context.Logger.Info($"Persisted {persisted} sleeves", Name);
        }
        catch (Exception ex)
        {
            var result = context.ErrorPolicy.Handle(ex, Name, context, critical: false);
            if (!result.ShouldContinue)
            {
                throw; // Re-throw if critical
            }
        }
    }
}
```

## Migration Strategy

### Phase 1: New Services (Current)
- New services should accept `ISleevePersistenceService`, `IErrorPolicy`, and `ILogger` via constructor injection
- Use adapter defaults if not provided
- Existing services continue using direct repository/logging calls

### Phase 2: Gradual Migration (Future)
- Update existing services to accept persistence/error/logging via constructor
- Keep direct calls as fallback during transition
- Both approaches work simultaneously

### Phase 3: Full Migration (Future)
- Remove direct repository/logging calls from all services
- All services use injected abstractions
- Centralized error handling in orchestrator

## Benefits

1. **Centralized Error Handling**: All errors handled through one policy
2. **Structured Logging**: Correlation IDs and structured data
3. **Testability**: Easy to inject test implementations
4. **Fail-Safe**: Continues processing on non-critical errors
5. **Coexistence**: No breaking changes to existing code
6. **Preserved Logs**: Existing log file names and formats maintained

## Constraints Met

✅ **Must preserve existing log file names** - LoggerAdapter uses same file names  
✅ **Must preserve message prefixes** - Format maintained with correlation ID  
✅ **DB schemas untouched** - Adapter uses existing repository methods  
✅ **All persistence operations route through abstraction** - ISleevePersistenceService  
✅ **Error handling no longer scattered** - IErrorPolicy centralizes decisions  

## Log File Names Preserved

- `placement_debug.log` - Debug and info messages
- `placement_errors.log` - Errors and warnings
- `placement_performance.log` - Performance metrics (via PerformanceMonitor)
- `placement_summary.log` - Summary statistics (via PerformanceMonitor)

## Acceptance Criteria Status

✅ **ISleevePersistenceService created** with async methods  
✅ **IErrorPolicy created** with FailSafeErrorPolicy implementation  
✅ **ILogger abstraction created** wrapping DebugLogger + SafeFileLogger  
✅ **Adapters created** with coexistence pattern  
✅ **Existing log file names preserved**  
✅ **Existing message prefixes preserved**  
✅ **All persistence operations route through abstraction**  
✅ **Error handling centralized** via IErrorPolicy  

