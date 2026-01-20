# Data Persistence Code Flow Architecture
#
# Transaction Management & WPF/Revit Safety Checklist

## WPF/Revit Transaction Safety Checklist

When implementing transaction management and UI interactions in Revit add-ins, always follow these best practices to ensure stability and prevent crashes:

- [ ] **Use ExternalEvent for all Revit API calls from WPF**
    - Never call the Revit API directly from WPF event handlers or background threads.
- [ ] **Close or hide all modal dialogs before calling ExternalEvent.Raise()**
    - Modal dialogs block the Revit UI thread and can cause deadlocks or crashes if you raise an external event while they are open.
- [ ] **Use Dispatcher.BeginInvoke to defer ExternalEvent.Raise()**
    - This ensures the UI is fully closed before the event is raised, preventing re-entrancy issues.
- [ ] **Never use async/await for Revit API calls unless you fully marshal back to the Revit context**
    - Async code can easily break Revit’s single-threaded model.
- [ ] **Wrap all transaction code in try/catch and always call Transaction.Commit() or Transaction.RollBack()**
    - Uncommitted or abandoned transactions can leave Revit in an unstable state.
- [ ] **Avoid long-running operations in the UI thread**
    - Use background threads for heavy computation, but marshal results back to the UI and then to the Revit API context via ExternalEvent.

**Summary:**
Your implementation should always close/hide modal dialogs before raising external events, marshal all Revit API calls through ExternalEvent, and ensure all transaction code is exception-safe and properly committed or rolled back.

## Overview

This document describes the comprehensive data persistence architecture in the MEP Openings system, focusing on how data flows from business logic creation to database storage. The system uses a multi-layered approach with SQLite as the primary persistence store, featuring repository patterns, transaction management, and optimized bulk operations.

## Architecture Layers

### 1. Database Layer (SQLite)
- **Primary Store**: SQLite database with project-specific files (`{ProjectName}_SleevePersistence.db`)
- **Schema**: Auto-migrating with version tracking via `SchemaMigrations` table
- **Performance**: WAL mode, memory-mapped I/O, optimized cache sizes
- **Key Tables**:
  - `ClashZones`: Core clash detection data
  - `Filters`: Filter configurations
  - `FileCombos`: File combination metadata
  - `SleeveSnapshots`: Parameter snapshots for placed sleeves
  - `ClusterSleeves`: Cluster sleeve data
  - `ClashZonesRTree`: Spatial indexing (optional)

### 2. Repository Layer
- **Pattern**: Repository pattern with `IClashZoneRepository` interface
- **Implementation**: `ClashZoneRepository` with bulk operations
- **Features**: Transaction management, error handling, logging

### 3. Service Layer
- **Business Logic**: Services like `ClashZonePersistenceService`, `UniversalSleevePlacerService`
- **Data Transformation**: Convert business objects to database entities
- **Transaction Coordination**: Manage complex multi-table operations

### 4. Business Logic Layer
- **Data Creation**: Generate `ClashZone` objects during refresh/detection
- **State Management**: Track resolved/unresolved states, sleeve placements
- **Parameter Collection**: Gather MEP and host element parameters

## Data Flow: Complete Persistence Pipeline

### Phase 1: Data Creation (Business Logic)

```csharp
// 1. Clash Detection Service creates ClashZone objects
var clashZone = new ClashZone
{
    Id = GenerateDeterministicGuid(mepId, hostId, intersectionPoint),
    MepElementId = new ElementId(mepId),
    StructuralElementId = new ElementId(hostId),
    IntersectionPoint = intersectionPoint,
    MepElementCategory = "Pipes",
    // ... other properties
};

// 2. Parameter collection during refresh
clashZone.MepParameterValues = CollectMepParameters(mepElement);
clashZone.HostParameterValues = CollectHostParameters(hostElement);

// 3. Geometry calculations
clashZone.WallCenterlinePointX = CalculateWallCenterline(intersectionPoint, wallDirection);
clashZone.SleeveBoundingBoxMinX = CalculateBoundingBox(sleeveGeometry).Min.X;
// ... other geometry fields
```

### Phase 2: Service Layer Processing

```csharp
// ClashZonePersistenceService.cs
public void SaveClashZones(List<ClashZone> zones, string filterName, string category)
{
    // 1. Validate and filter zones
    var validZones = zones.Where(z => z != null && IsValidZone(z)).ToList();

    // 2. Group by category for batch processing
    var zonesByCategory = validZones.GroupBy(z => z.MepElementCategory);

    // 3. Process each category
    foreach (var categoryGroup in zonesByCategory)
    {
        // 4. Call repository for persistence
        _sqliteRepository.InsertOrUpdateClashZones(
            categoryGroup.ToList(),
            filterName,
            categoryGroup.Key
        );
    }
}
```

### Phase 3: Repository Layer Persistence

```csharp
// ClashZoneRepository.cs
public void InsertOrUpdateClashZones(IEnumerable<ClashZone> clashZones, string filterName, string category)
{
    // 1. Use bulk operations for performance
    if (OptimizationFlags.UseBulkSqliteUpdates)
    {
        InsertOrUpdateClashZonesBulk(clashZones, filterName, category);
        return;
    }

    // 2. Legacy individual processing (fallback)
    using (var transaction = _context.Connection.BeginTransaction())
    {
        try
        {
            // Get/create filter metadata
            var filterId = GetOrCreateFilter(filterName, category, transaction);

            // Get/create file combo
            var comboId = GetOrCreateFileCombo(filterId, category, hostCategories, clashZone, transaction);

            // Insert/update each zone
            foreach (var zone in clashZones)
            {
                InsertOrUpdateClashZone(comboId, zone, transaction);
            }

            // Save parameter snapshots
            InsertOrUpdateSleeveSnapshots(filterId, processedZones, transaction);

            transaction.Commit();
        }
        catch (Exception ex)
        {
            transaction.Rollback();
            throw;
        }
    }
}
```

### Phase 4: Database Operations

```csharp
// Individual zone persistence
private bool InsertOrUpdateClashZone(int comboId, ClashZone clashZone, SQLiteTransaction transaction)
{
    // 1. GUID-based lookup (deterministic matching)
    var existingId = FindExistingByGuid(clashZone.Id.ToString());

    if (existingId.HasValue)
    {
        // Update existing
        UpdateClashZone(existingId.Value, comboId, clashZone, transaction);
        return false;
    }
    else
    {
        // Insert new
        InsertClashZone(comboId, clashZone, transaction);
        return true;
    }
}

// Bulk operations for performance
private void BulkUpdateClashZones(List<ClashZone> zones, Dictionary<Guid, int> existingMap, Dictionary<Guid, int> comboMap, SQLiteTransaction transaction)
{
    // Build dynamic SQL with CASE statements for bulk update
    var sql = BuildBulkUpdateSql(zones, existingMap, comboMap);
    using (var cmd = _context.Connection.CreateCommand())
    {
        cmd.Transaction = transaction;
        cmd.CommandText = sql.ToString();

        // Add parameters for all zones in batch
        AddBulkUpdateParameters(cmd, zones, existingMap, comboMap);

        cmd.ExecuteNonQuery();
    }
}
```

## Example: Persisting Sleeve Corners for Cluster Sleeves

### Business Context
When a cluster sleeve is placed, its 4 corner coordinates (in world space) need to be persisted for downstream processes like parameter transfer and visualization. These corners are calculated once during placement and stored for reuse.

### Code Flow Example

#### 1. Corner Calculation (Business Logic)
```csharp
// CombinedClusterFormationService.cs - After cluster placement
private void CalculateAndPersistClusterCorners(ClusterSleeve cluster, List<ClashZone> zonesInCluster)
{
    // 1. Calculate 4 corners in world space
    var corners = CalculateClusterCorners(cluster.BoundingBox, cluster.RotationAngle);

    // 2. Update cluster object with corner data
    cluster.Corner1X = corners[0].X;
    cluster.Corner1Y = corners[0].Y;
    cluster.Corner1Z = corners[0].Z;
    // ... Corner2X, Corner2Y, etc.

    // 3. Persist to database
    _repository.UpdateClusterSleeveCorners(
        cluster.ClusterInstanceId,
        corners[0].X, corners[0].Y, corners[0].Z,  // Corner 1
        corners[1].X, corners[1].Y, corners[1].Z,  // Corner 2
        corners[2].X, corners[2].Y, corners[2].Z,  // Corner 3
        corners[3].X, corners[3].Y, corners[3].Z   // Corner 4
    );
}
```

#### 2. Repository Method (Data Access)
```csharp
// ClashZoneRepository.cs
public void UpdateClusterSleeveCorners(int clusterInstanceId,
    double corner1X, double corner1Y, double corner1Z,
    double corner2X, double corner2Y, double corner2Z,
    double corner3X, double corner3Y, double corner3Z,
    double corner4X, double corner4Y, double corner4Z)
{
    using (var cmd = _context.Connection.CreateCommand())
    {
        cmd.CommandText = @"
            UPDATE ClusterSleeves SET
                Corner1X = @Corner1X, Corner1Y = @Corner1Y, Corner1Z = @Corner1Z,
                Corner2X = @Corner2X, Corner2Y = @Corner2Y, Corner2Z = @Corner2Z,
                Corner3X = @Corner3X, Corner3Y = @Corner3Y, Corner3Z = @Corner3Z,
                Corner4X = @Corner4X, Corner4Y = @Corner4Y, Corner4Z = @Corner4Z,
                UpdatedAt = CURRENT_TIMESTAMP
            WHERE ClusterInstanceId = @ClusterInstanceId";

        // Add all 12 coordinate parameters
        cmd.Parameters.AddWithValue("@ClusterInstanceId", clusterInstanceId);
        cmd.Parameters.AddWithValue("@Corner1X", corner1X);
        cmd.Parameters.AddWithValue("@Corner1Y", corner1Y);
        cmd.Parameters.AddWithValue("@Corner1Z", corner1Z);
        // ... 8 more corner coordinate parameters

        var rowsAffected = cmd.ExecuteNonQuery();
        if (rowsAffected == 0)
        {
            _logger($"⚠️ UpdateClusterSleeveCorners: No rows updated for ClusterInstanceId {clusterInstanceId}");
        }
        else
        {
            _logger($"✅ Saved 4 corners for ClusterInstanceId {clusterInstanceId}");
        }
    }
}
```

#### 3. Database Schema (Storage)
```sql
-- ClusterSleeves table schema for corner storage
CREATE TABLE ClusterSleeves (
    ClusterSleeveId INTEGER PRIMARY KEY,
    ClusterInstanceId INTEGER NOT NULL UNIQUE,
    -- ... other cluster fields ...

    -- 4 corner coordinates in world space (Phase 3)
    Corner1X REAL DEFAULT 0.0,
    Corner1Y REAL DEFAULT 0.0,
    Corner1Z REAL DEFAULT 0.0,
    Corner2X REAL DEFAULT 0.0,
    Corner2Y REAL DEFAULT 0.0,
    Corner2Z REAL DEFAULT 0.0,
    Corner3X REAL DEFAULT 0.0,
    Corner3Y REAL DEFAULT 0.0,
    Corner3Z REAL DEFAULT 0.0,
    Corner4X REAL DEFAULT 0.0,
    Corner4Y REAL DEFAULT 0.0,
    Corner4Z REAL DEFAULT 0.0,

    CreatedAt DATETIME DEFAULT CURRENT_TIMESTAMP,
    UpdatedAt DATETIME DEFAULT CURRENT_TIMESTAMP
);
```

#### 4. Retrieval for Downstream Use
```csharp
// ParameterTransferService.cs - Using persisted corners
public void TransferParametersToClusterSleeve(int clusterInstanceId)
{
    // 1. Load cluster data including corners
    var clusterData = _repository.GetClusterSleeveByInstanceId(clusterInstanceId);

    // 2. Use corners for geometric calculations
    var corners = new List<XYZ>
    {
        new XYZ(clusterData.Corner1X, clusterData.Corner1Y, clusterData.Corner1Z),
        new XYZ(clusterData.Corner2X, clusterData.Corner2Y, clusterData.Corner2Z),
        new XYZ(clusterData.Corner3X, clusterData.Corner3Y, clusterData.Corner3Z),
        new XYZ(clusterData.Corner4X, clusterData.Corner4Y, clusterData.Corner4Z)
    };

    // 3. Calculate sleeve center from corners
    var sleeveCenter = CalculateCenterFromCorners(corners);

    // 4. Use center for parameter transfer logic
    TransferParametersToSleeve(sleeveCenter, clusterData.MepElementIds);
}
```

## Key Design Patterns

### 1. Repository Pattern
- **Interface Segregation**: `IClashZoneRepository` defines contract
- **Implementation**: `ClashZoneRepository` handles SQLite operations
- **Testability**: Easy to mock for unit testing

### 2. Unit of Work Pattern
- **Transactions**: All related operations in single transaction
- **Consistency**: Either all operations succeed or all fail
- **Performance**: Reduced database round trips

### 3. Bulk Operations
- **Performance**: Batch INSERT/UPDATE with CASE statements
- **Scalability**: Handle thousands of records efficiently
- **Memory Management**: Process data in chunks

### 4. Deterministic GUIDs
- **Consistency**: Same data always generates same GUID
- **Merging**: Detect duplicate zones across sessions
- **Referential Integrity**: Reliable foreign key relationships

## Critical Optimization: Place All → Regen Once → Batch Read → Batch Write

### Overview
The system implements a crucial performance optimization pattern that dramatically improves placement efficiency:

1. **Place ALL Sleeves First**: Create all Revit family instances in a single transaction
2. **Regen ONE TIME**: Single document regeneration instead of per-sleeve regeneration
3. **Batch Read Corner Data**: Parallel calculation of geometric corner coordinates from Revit
4. **Batch Update Database**: Bulk persistence of all geometric data

### Performance Impact
- **Before**: N regenerations (one per sleeve) + N individual DB writes
- **After**: 1 regeneration + parallel corner calculation + 1 bulk DB write
- **Improvement**: 10-20x faster for large placements (100+ sleeves)

### Implementation Details

#### 1. Place All Sleeves First
```csharp
// NewSleevePlacerService.cs - ExecutePlacementInternal()
foreach (var clashZone in filteredZones)
{
    // Place sleeve in Revit (no regeneration yet)
    placedSleeve = PlaceSleeveNormal(clashZone);
    placedSleeveData.Add((placedSleeve, clashZone, width, height, depth));
}

// Single regeneration for ALL sleeves
_doc.Regenerate();
```

#### 2. Single Regeneration
```csharp
// After ALL sleeves placed, regenerate ONCE
if (placedSleeveData.Count > 0)
{
    var regenTimer = System.Diagnostics.Stopwatch.StartNew();
    _doc.Regenerate(); // ✅ Single regeneration for all sleeves
    regenTimer.Stop();
    DebugLogger.Info($"✅ Regenerated {placedSleeveData.Count} sleeves in {regenTimer.ElapsedMilliseconds}ms");
}
```

#### 3. Batch Read Corner Data (Parallel)
```csharp
// SleevePersistenceService.cs - PersistSleeveData()
// Pre-calculate ALL geometry in parallel (corners + rotated bboxes simultaneously)
if (OptimizationFlags.UseParallelProcessing && placedSleeveData.Count >= 12)
{
    var geometryTasks = placedSleeveData.Select(item => Task.Run(() =>
    {
        // Calculate corners for each sleeve in parallel
        var corners = _cornerCalculationService.CalculateCornersFromZone(zone, width, height);
        return (zone.Id, corners);
    })).ToArray();

    Task.WaitAll(geometryTasks);

    // Aggregate results
    foreach (var task in geometryTasks)
    {
        var (zoneId, corners) = task.Result;
        cornerData[zoneId] = corners;
    }
}
```

#### 4. Batch Write to Database
```csharp
// After regeneration, batch save ALL data
if (placedSleeveData.Count > 0)
{
    int persistedCount = _persistenceService.PersistSleeveData(placedSleeveData, filterName);
    // Saves: instance IDs, placement points, bounding boxes, corners, snapshots
}
```

### Why This Optimization Works

#### Problem Solved
- **Multiple Regenerations**: Each `doc.Regenerate()` is expensive (100-500ms)
- **Stale Bounding Boxes**: Reading geometry before regeneration gives wrong data
- **Sequential Processing**: No parallelism in geometric calculations
- **Individual DB Writes**: High latency for per-row operations

#### Solution Benefits
- **Single Regeneration**: Eliminates 99% of regeneration overhead
- **Fresh Geometry**: All bounding boxes calculated after regeneration
- **Parallel Corner Calculation**: CPU-intensive math done in parallel
- **Bulk Database Operations**: Minimize DB round trips with CASE statements

### Technical Implementation

#### Parallel Corner Calculation
```csharp
// Pre-calculate corners in parallel before DB writes
var cornerData = new Dictionary<Guid, (double c1x, double c1y, double c1z, ...)>();

// Parallel processing for 12+ sleeves
if (placedSleeveData.Count >= 12)
{
    var tasks = placedSleeveData.Select(item => Task.Run(() => {
        var corners = CalculateCorners(item.zone, item.width, item.height);
        return (item.zone.Id, corners);
    })).ToArray();

    Task.WaitAll(tasks);
    cornerData = tasks.ToDictionary(t => t.Result.Item1, t => t.Result.Item2);
}
```

#### Bulk Database Updates
```csharp
// Build SQL with CASE statements for bulk update
var sql = @"
    UPDATE ClashZones SET
        Corner1X = CASE ClashZoneId
            WHEN @Id1 THEN @C1X1
            WHEN @Id2 THEN @C1X2
            -- ... for all zones
        END,
        Corner1Y = CASE ClashZoneId
            WHEN @Id1 THEN @C1Y1
            WHEN @Id2 THEN @C1Y2
            -- ... for all zones
        END
    WHERE ClashZoneId IN (@Id1, @Id2, ...)
";

// Single DB call updates all zones
cmd.ExecuteNonQuery();
```

## Performance Optimizations

### 1. Database Level
- **Indexes**: Optimized for common query patterns
- **WAL Mode**: Concurrent reads during writes
- **Memory Mapping**: Faster I/O for large datasets
- **Prepared Statements**: Reusable SQL execution plans

### 2. Application Level
- **Bulk Operations**: Reduce network round trips
- **Lazy Loading**: Load data only when needed
- **Caching**: Session-level schema verification cache
- **Async Processing**: Non-blocking database operations

### 3. Query Optimization
- **R-Tree Indexes**: Spatial queries for section box filtering
- **Parameterized Queries**: SQL injection prevention and plan reuse
- **Batch Processing**: Group similar operations

## Error Handling & Recovery

### 1. Transaction Rollback
```csharp
using (var transaction = _context.Connection.BeginTransaction())
{
    try
    {
        // Multiple database operations
        InsertClashZone(comboId, zone, transaction);
        UpdateSleeveSnapshots(snapshot, transaction);

        transaction.Commit();
    }
    catch (Exception ex)
    {
        transaction.Rollback();
        _logger($"❌ Persistence failed: {ex.Message}");
        throw;
    }
}
```

### 2. Graceful Degradation
- **Fallback Paths**: Bulk → Individual operations
- **Partial Success**: Log warnings for failed individual items
- **Data Integrity**: Validate critical fields before persistence

### 3. Logging & Diagnostics
- **Operation Tracking**: Log all database operations with timing
- **Parameter Validation**: Verify data integrity before persistence
- **Performance Metrics**: Track operation duration and success rates

## Migration & Schema Evolution

### 1. Automatic Schema Upgrades
```csharp
private void EnsureSchemaUpgraded()
{
    // Add new columns without breaking existing data
    AddColumnIfMissing("ClashZones", "SleeveCorner1X", "REAL", transaction);
    AddColumnIfMissing("ClashZones", "SleeveCorner1Y", "REAL", transaction);
    // ... etc for all new fields
}
```

### 2. Backward Compatibility
- **Optional Fields**: New columns with DEFAULT values
- **Version Tracking**: `SchemaMigrations` table prevents duplicate upgrades
- **Safe Rollbacks**: ALTER TABLE operations are non-destructive

## Monitoring & Maintenance

### 1. Health Checks
- **Connection Validation**: Test database connectivity
- **Schema Verification**: Ensure all required tables/columns exist
- **Performance Monitoring**: Track query execution times

### 2. Diagnostic Tools
- **Query Logging**: Log all SQL operations with parameters
- **Data Validation**: Verify persisted data integrity
- **Performance Profiling**: Identify slow operations

### 3. Cleanup Operations
- **Orphaned Records**: Remove invalid references
- **Index Maintenance**: Rebuild indexes periodically
- **Space Reclamation**: VACUUM database files

## Debugging Lessons: Cluster Corner Persistence

### Real-World Case Study: Cluster Sleeve Corners Not Persisting

#### Problem
Cluster sleeve corner coordinates were being calculated correctly (e.g., `C1:(44.00,0.25,6.61)`) but showing as `0.0` in the database.

#### Debugging Journey (6 False Starts)

1. **Logging Suppression Issue**
   - **Symptom**: No debug logs appearing
   - **Root Cause**: `DeploymentConfiguration.DeploymentMode = true` suppressed all `SafeFileLogger` calls
   - **Lesson**: Always verify logging infrastructure is enabled before trusting absence of logs
   - **Fix**: Used direct `File.AppendAllText` to bypass deployment mode

2. **Execution Order Issue**
   - **Symptom**: Batch save not executing
   - **Root Cause**: Code was inside an `if` block that never executed
   - **Lesson**: Verify control flow - add explicit entry/exit logging
   - **Fix**: Moved code outside problematic conditional block

3. **SQL Column Duplication**
   - **Symptom**: Corners calculated but database showed `0.0`
   - **Root Cause**: Duplicate columns in INSERT statement
   ```sql
   -- WRONG:
   ClashZoneGuids, MepSizes, MepSystemNames, MepElementIds,
   ClashZoneGuids, MepSizes, MepSystemNames, MepElementIds,  -- DUPLICATES!
   Corner1X, Corner1Y, Corner1Z, ...
   ```
   - **Lesson**: SQL column count must exactly match parameter count
   - **Fix**: Removed duplicate columns

4. **Parameter Binding Duplication**
   - **Symptom**: Still `0.0` after fixing SQL
   - **Root Cause**: Duplicate `@MepSystemNames` parameter causing position shift
   ```csharp
   // WRONG:
   cmd.Parameters.AddWithValue($"@MepSystemNames{index}", ...);
   cmd.Parameters.AddWithValue($"@MepSystemNames{index}", ...); // DUPLICATE!
   // All corner parameters shifted by 1 → wrong values
   ```
   - **Lesson**: Parameter binding order MUST match SQL placeholder order exactly
   - **Fix**: Removed duplicate parameter

5. **Wrong Persistence Pattern**
   - **Symptom**: No UPDATE operations logged for `ClusterSleeves` table
   - **Root Cause**: Trying to use `INSERT OR REPLACE` instead of following existing pattern
   - **Lesson**: **CRITICAL - Always check how similar working code persists data first**
   - **Discovery**: Individual sleeves use separate UPDATE method:
   ```csharp
   // Individual sleeves (WORKING):
   repository.UpdateSleeveCorners(zone.Id, c1x, c1y, c1z, ...);
   
   // Cluster sleeves (SHOULD MATCH):
   repository.UpdateClusterSleeveCorners(clusterId, c1x, c1y, c1z, ...);
   ```

6. **Final Solution: Follow Established Pattern**
   ```csharp
   // In BatchSaveClusterDataToDatabase method:
   
   // 1. Create repository instances
   var clusterRepository = new ClusterSleeveRepository(dbContext);
   var clashZoneRepository = new ClashZoneRepository(dbContext); // For corners
   
   // 2. Calculate corners
   var cornersResult = _cornerService.CalculateCorners(placementPoint, width, height, rotationAngleDeg);
   
   // 3. Extract corner values
   if (cornersResult.HasValue)
   {
       var c = cornersResult.Value;
       c1x = c.corner1.X; c1y = c.corner1.Y; c1z = c.corner1.Z;
       // ... etc
       
       // 4. ✅ CRITICAL: Update corners in database (same pattern as individual sleeves)
       clashZoneRepository.UpdateClusterSleeveCorners(
           clusterInstanceId,
           c1x, c1y, c1z,
           c2x, c2y, c2z,
           c3x, c3y, c3z,
           c4x, c4y, c4z);
   }
   ```

### Key Takeaways

1. **Follow Existing Patterns**
   - Before implementing new persistence logic, check how similar data is already persisted
   - If individual sleeves use `UpdateSleeveCorners`, cluster sleeves should use `UpdateClusterSleeveCorners`
   - Don't reinvent the wheel - reuse proven patterns

2. **Verify Each Layer**
   - Calculation → Assignment → Parameter Binding → SQL Execution → Database Write
   - Add logging at each step to pinpoint where data is lost
   - Don't assume multiple steps work - verify each independently

3. **SQL/Parameter Alignment is Critical**
   - Column count must match parameter count exactly
   - Parameter binding order must match SQL placeholder order
   - Duplicates cause silent data corruption (no errors, just wrong data)

4. **Logging Infrastructure Matters**
   - Verify logging is enabled before trusting absence of logs
   - Use direct file I/O for critical debugging when logger may be suppressed
   - Check for deployment mode, feature flags, or conditional logging

5. **Database Operation Logging**
   - Always log database operations (INSERT, UPDATE, DELETE)
   - Include table name, operation type, and key parameters
   - Check logs to verify operations are actually executing

### Prevention Checklist

When implementing data persistence:

- [ ] Check how similar data is already persisted in the codebase
- [ ] Follow the same pattern (don't invent new approaches)
- [ ] Verify logging is enabled for debugging
- [ ] Add explicit entry/exit logs for critical methods
- [ ] Verify SQL column count matches parameter count
- [ ] Verify parameter binding order matches SQL placeholders
- [ ] Check for duplicate columns or parameters
- [ ] Log database operations with table name and operation type
- [ ] Verify each step independently (calculation → binding → execution → persistence)
- [ ] Check database operation logs to confirm execution
- [ ] Test with actual database queries to verify data is correct

## Conclusion

The data persistence architecture provides a robust, scalable foundation for the MEP Openings system. The layered approach with repository patterns, transaction management, and performance optimizations ensures reliable data storage while maintaining high performance even with large datasets. The sleeve corners example demonstrates how complex geometric data flows from business logic calculation through service layer processing to efficient database storage.

**Most Important Lesson**: When debugging persistence issues, always start by checking how similar working code persists data. Don't invent new patterns when proven patterns already exist. This single insight could have saved hours of debugging time.

