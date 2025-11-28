# Team B: Configuration & Environment Abstraction - Usage Examples

## Overview
Team B has created configuration and environment abstractions that coexist with existing `OptimizationFlags` and `DeploymentConfiguration` static classes. These abstractions enable Dependency Injection (DI) without breaking existing code.

## Key Deliverables

### 1. IOptimizationConfig Interface
Groups optimization flags into immutable records:
- `PlacementOptions` - Placement-related optimizations
- `DetectionOptions` - Spatial indexing and filtering
- `PerformanceOptions` - Performance monitoring

### 2. IAppEnvironment Interface
Encapsulates application environment:
- Deployment mode status
- Directory paths (logs, database, app data)
- Logging policy

### 3. Adapters
- `OptimizationConfigAdapter` - Reads from `OptimizationFlags`
- `AppEnvironmentAdapter` - Reads from `DeploymentConfiguration`

## Usage Examples

### Example 1: Basic Dependency Injection

```csharp
// In your service constructor
public class MyPlacementService
{
    private readonly IOptimizationConfig _config;
    private readonly IAppEnvironment _environment;
    
    public MyPlacementService(
        IOptimizationConfig config = null,
        IAppEnvironment environment = null)
    {
        // ✅ COEXISTENCE: Use factory defaults if not injected
        _config = config ?? OptimizationConfigFactory.Default;
        _environment = environment ?? AppEnvironmentFactory.Default;
    }
    
    public void DoWork()
    {
        // Use configuration instead of static flags
        if (_config.Placement.UseParameterBatching)
        {
            // Use batching
        }
        
        // Use environment instead of static DeploymentConfiguration
        if (!_environment.IsDeploymentMode)
        {
            // Log debug info
        }
    }
}
```

### Example 2: Testing with Custom Configuration

```csharp
[Test]
public void TestWithCustomConfig()
{
    // Create test-specific configuration
    var placementOptions = new PlacementOptions(
        UseParallelClearance: false,  // Disable for test
        UseParameterBatching: true,
        UseSmartReplay: false,
        UseNewSleevePlacerService: true,
        UseNewSleeveRepository: false
    );
    
    var detectionOptions = new DetectionOptions(
        UseRTreeIndex: false,
        UseSpatialGrid: false,
        UseRTreeDatabaseIndex: false,
        UseBoundingBoxSectionBoxFilter: false,
        UseCurveInBoundingBoxFilter: false,
        UseViewIndependentFilter: false
    );
    
    var performanceOptions = new PerformanceOptions(
        EnablePerformanceLogging: true,  // Enable for test
        EnableDetailedTiming: true
    );
    
    var testConfig = OptimizationConfigFactory.Create(
        placementOptions,
        detectionOptions,
        performanceOptions
    );
    
    // Use test config in service
    var service = new MyPlacementService(testConfig);
    // ... test code
}
```

### Example 3: Testing with Custom Environment

```csharp
[Test]
public void TestWithCustomEnvironment()
{
    // Create test environment
    var testEnvironment = AppEnvironmentFactory.Create(
        isDeploymentMode: false,  // Enable logging for test
        appDataDirectory: @"C:\Test\AppData",
        logDirectory: @"C:\Test\Logs",
        databaseDirectory: @"C:\Test\Database",
        loggingPolicy: LoggingPolicy.DevelopmentMode
    );
    
    var service = new MyPlacementService(environment: testEnvironment);
    // ... test code
}
```

### Example 4: Gradual Migration Pattern

```csharp
// OLD CODE (still works):
if (OptimizationFlags.UseParameterBatching)
{
    // ...
}

// NEW CODE (preferred for new services):
if (_config.Placement.UseParameterBatching)
{
    // ...
}

// ✅ COEXISTENCE: Both work simultaneously during migration
```

### Example 5: Logging Policy Usage

```csharp
public class MyService
{
    private readonly IAppEnvironment _environment;
    
    public MyService(IAppEnvironment environment)
    {
        _environment = environment;
    }
    
    public void LogSomething(string message)
    {
        // Check logging policy instead of DeploymentMode directly
        if (_environment.LoggingPolicy.EnableInfoLogging)
        {
            DebugLogger.Info(message);
        }
        
        if (_environment.LoggingPolicy.EnablePerformanceLogging)
        {
            _performanceMonitor.LogMetric("Something", message);
        }
    }
}
```

## Migration Strategy

### Phase 1: New Services (Current)
- New services should accept `IOptimizationConfig` and `IAppEnvironment` via constructor injection
- Use factory defaults if not provided: `OptimizationConfigFactory.Default`
- Existing services continue using static classes

### Phase 2: Gradual Migration (Future)
- Update existing services to accept config/environment via constructor
- Keep static class access as fallback during transition
- Both approaches work simultaneously

### Phase 3: Full Migration (Future)
- Remove static class access from all services
- Deprecate `OptimizationFlags` and `DeploymentConfiguration` (keep for backward compatibility)
- All services use injected configuration

## Benefits

1. **Testability**: Easy to inject test configurations
2. **SOLID Compliance**: Dependency Inversion Principle (DIP)
3. **Coexistence**: No breaking changes to existing code
4. **Type Safety**: Immutable records prevent accidental mutations
5. **Organization**: Related flags grouped logically

## Constraints (Team B Requirements)

✅ **Must NOT delete or alter `OptimizationFlags`** - Adapter reads from it  
✅ **Must NOT delete or alter `DeploymentConfiguration`** - Adapter reads from it  
✅ **No business logic moves** - Only configuration abstraction  
✅ **All new services resolve via DI** - Factory pattern provides defaults  

## Acceptance Criteria Status

- ✅ `IOptimizationConfig` interface created with grouped records
- ✅ Adapter reading from `OptimizationFlags` implemented
- ✅ `IAppEnvironment` interface created
- ✅ Adapter reading from `DeploymentConfiguration` implemented
- ✅ Factory classes for easy DI integration
- ✅ Coexistence with existing static classes maintained
- ✅ No business logic moved (only configuration abstraction)

## Next Steps (After Team A Interfaces Finalized)

Once Team A finalizes `ISleevePlacementOrchestrator` and `IPlacementStage` interfaces:
- Replace direct static flag reads in non-Team-A files via injected `IOptimizationConfig`
- Update services to use `IAppEnvironment` instead of `DeploymentConfiguration.DeploymentMode`
- Maintain backward compatibility during transition

