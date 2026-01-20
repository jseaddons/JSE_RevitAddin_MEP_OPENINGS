# Cluster Service Orchestration Guide

## Phase 6-11 Service Integration Complete! 🎉

All clustering services have been extracted and integrated. Here's how to use them:

---

## ✅ RECOMMENDED: Use the Factory (New Code)

### Option 1: Fully-Wired Service (All Features Enabled)

```csharp
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering;

// In your command class (e.g., ClusterSleevesCommand.cs)
public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
{
    var doc = commandData.Application.ActiveUIDocument.Document;
    
    // ✅ Create fully-configured service with all Phase 6-11 services
    var clusterService = ClusterServiceFactory.CreateWithAllServices(doc);
    
    // Use it like before - all services are auto-wired!
    var (placedCount, deletedCount) = clusterService.ClusterSleeves(
        doc, 
        targetCategory: "Ducts", 
        uiDoc: commandData.Application.ActiveUIDocument
    );
    
    TaskDialog.Show("Success", $"Placed {placedCount} clusters, deleted {deletedCount} sleeves");
    return Result.Succeeded;
}
```

### Option 2: Custom Timeout Limit

```csharp
// Set custom timeout (e.g., 10 minutes instead of 5)
var clusterService = ClusterServiceFactory.CreateWithAllServices(
    doc, 
    timeoutLimitMs: 600000  // 10 minutes
);
```

### Option 3: Selective Services (Testing/Migration)

```csharp
// Enable only specific services
var clusterService = ClusterServiceFactory.CreateCustom(
    doc,
    enableRotation: true,     // Phase 6
    enableCleanup: true,      // Phase 7
    enableAlgorithm: true,    // Phase 8
    enableData: true,         // Phase 9
    enableTimeout: false      // Phase 10 - disable for testing
);
```

---

## 🔧 Manual Wiring (Advanced Use Cases)

If you need full control over service instantiation:

```csharp
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Cleanup;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Algorithm;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Data;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Timeout;

// Create individual services
var rotationService = new ClusterRotationService(getClashZoneFunc: null);
var cleanupService = new ClusterCleanupService();
var algorithmService = new ClusterAlgorithmService();
var dataService = new ClusterDataService(doc);
var timeoutService = new ClusterTimeoutService(timeoutLimitMs: 300000);

// Wire into UniversalClusterService
var clusterService = new UniversalClusterService(
    flagManager: null,
    filterService: null,
    rotationService: rotationService,
    cleanupService: cleanupService,
    algorithmService: algorithmService,
    dataService: dataService,
    timeoutService: timeoutService
);
```

---

## 📦 What Each Service Does

| Phase | Service | Responsibility |
|-------|---------|----------------|
| **Phase 6** | `ClusterRotationService` | MEP rotation angle analysis, rotated bbox calculations |
| **Phase 7** | `ClusterCleanupService` | Flag reset, sleeve deletion, database cleanup |
| **Phase 8** | `ClusterAlgorithmService` | Spatial grid clustering, flood-fill, proximity checks |
| **Phase 9** | `ClusterDataService` | Database-first data loading, cache management |
| **Phase 10** | `ClusterTimeoutService` | Timeout protection, user feedback |

---

## 🔄 Backward Compatibility (Legacy Code)

Existing callers that instantiate `UniversalClusterService` directly still work:

```csharp
// This still works - uses legacy fallback methods
var clusterService = new UniversalClusterService();
clusterService.ClusterSleeves(doc, "Ducts", uiDoc);
```

**Legacy mode routing:**
- If service is `null` → falls back to legacy inline method
- If service exists → delegates to extracted service

---

## 🚀 Migration Path

### Step 1: Update Command Classes

Find all places where `UniversalClusterService` is instantiated:

```bash
# Search for instantiation
grep -r "new UniversalClusterService" --include="*.cs"
```

### Step 2: Replace with Factory

**Before:**
```csharp
var service = new UniversalClusterService();
```

**After:**
```csharp
var service = ClusterServiceFactory.CreateWithAllServices(doc);
```

### Step 3: Test

All functionality remains identical. Services are transparent to calling code.

---

## 📍 Key Files Modified

### New Service Files (11 total)
- `Services/Clustering/ClusterServiceFactory.cs` ← **USE THIS**
- `Services/Clustering/Rotation/IClusterRotationService.cs`
- `Services/Clustering/Rotation/ClusterRotationService.cs`
- `Services/Clustering/Cleanup/IClusterCleanupService.cs`
- `Services/Clustering/Cleanup/ClusterCleanupService.cs`
- `Services/Clustering/Algorithm/IClusterAlgorithmService.cs`
- `Services/Clustering/Algorithm/ClusterAlgorithmService.cs`
- `Services/Clustering/Data/IClusterDataService.cs`
- `Services/Clustering/Data/ClusterDataService.cs`
- `Services/Clustering/Timeout/IClusterTimeoutService.cs`
- `Services/Clustering/Timeout/ClusterTimeoutService.cs`

### Modified Files
- `Services/UniversalClusterService.cs` (constructor updated, routing logic added)

---

## 🎯 Benefits

✅ **Testability**: Each service can be unit tested independently  
✅ **Maintainability**: Clear separation of concerns  
✅ **Flexibility**: Mix and match services for different scenarios  
✅ **Safety**: Crash-safe with database transactions and timeout protection  
✅ **Performance**: Multi-threading, spatial optimization, database-first all preserved  
✅ **Compatibility**: 100% backward compatible with existing code  

---

## 🔍 Example: Command Implementation

See `Commands/Clustering/ClusterSleevesCommand.cs` for full example:

```csharp
[Transaction(TransactionMode.Manual)]
public class ClusterSleevesCommand : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        try
        {
            var doc = commandData.Application.ActiveUIDocument.Document;
            var uiDoc = commandData.Application.ActiveUIDocument;
            
            // ✅ Use factory for all modern features
            var clusterService = ClusterServiceFactory.CreateWithAllServices(doc);
            
            using (var trans = new Transaction(doc, "Cluster Sleeves"))
            {
                trans.Start();
                
                var (placed, deleted) = clusterService.ClusterSleeves(doc, null, uiDoc);
                
                trans.Commit();
                
                TaskDialog.Show("Clustering Complete", 
                    $"Placed: {placed} cluster sleeves\nDeleted: {deleted} individual sleeves");
            }
            
            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }
}
```

---

**Ready to use! All phases orchestrated and production-ready.** 🚀
