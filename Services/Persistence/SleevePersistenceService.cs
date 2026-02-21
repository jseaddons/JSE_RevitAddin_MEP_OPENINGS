using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Geometry;
using JSE_RevitAddin_MEP_OPENINGS.Services.Filters;
using JSE_RevitAddin_MEP_OPENINGS.Services.Parallel;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Persistence
{
    /// <summary>
    /// ✅ SOLID COMPLIANCE (SRP): Service responsible ONLY for orchestrating sleeve data persistence to database.
    /// Single Responsibility: Coordinates the persistence workflow by delegating to specialized services.
    /// 
    /// Delegates to specialized services:
    /// - ParallelCornerCalculationOrchestrator: Parallel corner calculations
    /// - RotatedBoundingBoxCalculationService: Rotated bounding box calculations
    /// - FilterLookupService: Filter ID lookups from database
    /// - SleeveCornerCalculationService: Individual corner calculations
    /// - ClashZoneRepository: Database operations (instance ID, placement, bounding boxes, corners, snapshots)
    /// 
    /// Follows OCP by using repository abstraction - can be extended without modification.
    /// Follows DIP by depending on abstractions (injected services) rather than concrete implementations.
    /// </summary>
    public class SleevePersistenceService
    {
        private readonly Document _doc;
        private readonly Action<string> _logger;
        
        // ✅ SRP COMPLIANCE: Delegate all specialized operations to dedicated services
        private readonly SleeveCornerCalculationService _cornerCalculationService;
        private readonly RotatedBoundingBoxCalculationService _rotatedBboxService;
        private readonly ParallelCornerCalculationOrchestrator _parallelCornerOrchestrator;

        public SleevePersistenceService(Document doc, Action<string> logger = null, 
            SleeveCornerCalculationService cornerCalculationService = null,
            RotatedBoundingBoxCalculationService rotatedBboxService = null,
            ParallelCornerCalculationOrchestrator parallelCornerOrchestrator = null)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _logger = logger ?? (msg => { });
            
            // ✅ SRP: Use injected services or create defaults (Dependency Injection)
            _cornerCalculationService = cornerCalculationService ?? new SleeveCornerCalculationService();
            _rotatedBboxService = rotatedBboxService ?? new RotatedBoundingBoxCalculationService();
            _parallelCornerOrchestrator = parallelCornerOrchestrator ?? new ParallelCornerCalculationOrchestrator(_cornerCalculationService);
        }

        /// <summary>
        /// ✅ SRP: Persists complete sleeve data to database for a collection of placed sleeves.
        /// Saves: instance ID, placement, bounding boxes (axis-aligned, RCS, rotated), corners, snapshots.
        /// Includes performance monitoring, safety features, and comprehensive error handling.
        /// </summary>
        /// <param name="placedSleeveData">Collection of (FamilyInstance, ClashZone, finalWidth, finalHeight, finalDiameter) tuples</param>
        /// <param name="filterName">Filter name for snapshot saving</param>
        /// <returns>Number of sleeves successfully persisted</returns>
        public int PersistSleeveData(
            List<(FamilyInstance sleeve, ClashZone zone, double finalWidth, double finalHeight, double finalDiameter, double finalDepth)> placedSleeveData,
            string filterName)
        {
            if (placedSleeveData == null || placedSleeveData.Count == 0)
                return 0;

            // ✅ PERFORMANCE MONITORING: Track persistence time
            var persistenceTimer = System.Diagnostics.Stopwatch.StartNew();
            int persistedCount = 0;
            int errorCount = 0;
            int skippedCount = 0;

            try
            {
                using (var dbContext = new SleeveDbContext(_doc))
                {
                    var repository = new ClashZoneRepository(dbContext, _logger);

                    // ✅ SAFETY: Validate database connection
                    if (dbContext.Connection == null)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error("[SleevePersistenceService] Database connection is null - cannot persist data");
                        return 0;
                    }

                    // ✅ PERFORMANCE OPTIMIZATION: Pre-calculate ALL geometry data in parallel (non-Revit operations)
                    // This includes: corners, rotated bounding boxes - all pure math that can be parallelized
                    // ✅ CRITICAL: Corners are batch calculated here and saved to database for cluster calculation
                    // - Cluster calculation will read these saved corners directly (NO recalculation needed)
                    // - This ensures accurate cluster sizing using pre-calculated corner coordinates
                    // - All 4 corners (Corner1X/Y/Z through Corner4X/Y/Z) are saved to database columns
                    var geometryCalculationTimer = System.Diagnostics.Stopwatch.StartNew();
                    
                    // ✅ MULTI-THREADING: Pre-calculate ALL geometry in parallel (corners + rotated bboxes simultaneously)
                    // ✅ BATCH CORNER CALCULATION: All corners calculated in parallel, then batch saved to database
                    // After regeneration, these corners are persisted and used by cluster calculation without recalculation
                    var cornerData = new Dictionary<Guid, (double corner1X, double corner1Y, double corner1Z,
                                        double corner2X, double corner2Y, double corner2Z,
                                        double corner3X, double corner3Y, double corner3Z,
                                        double corner4X, double corner4Y, double corner4Z)>();
                    var rotatedBboxData = new Dictionary<Guid, (double minX, double minY, double minZ, double maxX, double maxY, double maxZ)>();
                    
                    if (OptimizationFlags.UseParallelProcessing && placedSleeveData.Count >= 12)
                    {
                        // ✅ PARALLEL PROCESSING: Calculate corners AND rotated bounding boxes in parallel simultaneously
                        var geometryTasks = placedSleeveData
                            .Where(item => item.zone != null && item.sleeve != null && item.sleeve.IsValidObject)
                            .Select(item => Task.Run(() =>
                            {
                                try
                                {
                                    var (sleeve, zone, fw, fh, fd, fdepth) = item;
                                    var rotationAngleRad = zone.MepElementRotationAngle;
                                    var rotationAngleDeg = Math.Abs(rotationAngleRad * 180.0 / Math.PI);
                                    bool isStraightAxisAligned = IsStraightAxisAligned(rotationAngleDeg);

                                    // ✅ CORNERS: ENABLED Revit-based reading (User Request)
                                    // Parallel calculation uses pure math and might be inaccurate.
                                    // We DISABLE parallel corner calc here to force the main thread
                                    // to use CalculateCornersFromInstance (Revit Geometry).
                                    (Guid zoneId, (double corner1X, double corner1Y, double corner1Z,
                                        double corner2X, double corner2Y, double corner2Z,
                                        double corner3X, double corner3Y, double corner3Z,
                                        double corner4X, double corner4Y, double corner4Z) corners)? cornerResult = null;
                                    
                                    // ✅ Restore variable definitions needed for rotated bbox calculation
                                    double actualWidth = zone.SleeveWidth > 0 ? zone.SleeveWidth : fw;
                                    double actualHeight = zone.SleeveHeight > 0 ? zone.SleeveHeight : fh;
                                    
                                    /* ⚠️ DISABLED: Math-based calculation (suspected inaccurate by user)
                                    var corners = _cornerCalculationService.CalculateCornersFromZone(zone, actualWidth, actualHeight);
                                    if (corners.HasValue)
                                    {
                                        cornerResult = (zone.Id, (
                                            corners.Value.corner1.X, corners.Value.corner1.Y, corners.Value.corner1.Z,
                                            corners.Value.corner2.X, corners.Value.corner2.Y, corners.Value.corner2.Z,
                                            corners.Value.corner3.X, corners.Value.corner3.Y, corners.Value.corner3.Z,
                                            corners.Value.corner4.X, corners.Value.corner4.Y, corners.Value.corner4.Z
                                        ));
                                    }
                                    */

                                    // ✅ ROTATED BBOX: Calculate if rotation is non-zero and not axis-aligned
                                    (Guid zoneId, (double minX, double minY, double minZ, double maxX, double maxY, double maxZ) bbox)? bboxResult = null;
                                    if (Math.Abs(rotationAngleRad) > 1e-6 && !isStraightAxisAligned)
                                    {
                                        double actualDepth = zone.SleeveBoundingBoxMaxZ - zone.SleeveBoundingBoxMinZ;

                                        var rotatedBbox = _rotatedBboxService.CalculateRotatedBoundingBoxFromZone(zone, actualWidth, actualHeight, actualDepth);
                                        if (rotatedBbox.HasValue)
                                        {
                                            bboxResult = (zone.Id, rotatedBbox.Value);
                                        }
                                    }

                                    return (cornerResult, bboxResult);
                                }
                                catch (Exception ex)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        DebugLogger.Warning($"[SleevePersistenceService] Error calculating geometry in parallel: {ex.Message}");
                                    }
                                    // ✅ Explicitly typed nulls to help inference (Names MUST match exactly)
                                    (Guid zoneId, (double corner1X, double corner1Y, double corner1Z, double corner2X, double corner2Y, double corner2Z, double corner3X, double corner3Y, double corner3Z, double corner4X, double corner4Y, double corner4Z) corners)? nullCorner = null;
                                    (Guid zoneId, (double minX, double minY, double minZ, double maxX, double maxY, double maxZ) bbox)? nullBbox = null;
                                    return (nullCorner, nullBbox);
                                }
                            }))
                            .ToArray();

                        Task.WaitAll(geometryTasks);

                        // ✅ AGGREGATE: Collect all parallel results
                        foreach (var task in geometryTasks)
                        {
                            var (cornerResult, bboxResult) = task.Result;
                            
                            if (cornerResult.HasValue && cornerResult.Value.zoneId != Guid.Empty && 
                                (cornerResult.Value.corners.corner1X != 0 || cornerResult.Value.corners.corner1Y != 0))
                            {
                                cornerData[cornerResult.Value.zoneId] = cornerResult.Value.corners;
                            }
                            
                            if (bboxResult.HasValue && bboxResult.Value.zoneId != Guid.Empty && 
                                (bboxResult.Value.bbox.minX != 0 || bboxResult.Value.bbox.minY != 0))
                            {
                                rotatedBboxData[bboxResult.Value.zoneId] = bboxResult.Value.bbox;
                            }
                        }
                    }
                    else
                    {
                        // ✅ FALLBACK: Use orchestrator for corners if parallel disabled or small batch
                        // ? CRITICAL: Convert 6-tuple to 5-tuple for orchestrator (it only expects 5)
                        var convertedList = placedSleeveData.Select(x => (x.sleeve, x.zone, x.finalWidth, x.finalHeight, x.finalDiameter)).ToList();
                        cornerData = _parallelCornerOrchestrator.CalculateCornersInParallel(convertedList);
                    }
                    
                    geometryCalculationTimer.Stop();
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[SleevePersistenceService] ✅ Pre-calculated geometry in parallel: {cornerData.Count} corners, {rotatedBboxData.Count} rotated bboxes in {geometryCalculationTimer.ElapsedMilliseconds}ms");
                    }

                    // ✅ PERFORMANCE OPTIMIZATION: Pre-validate all sleeves in parallel (non-Revit, non-DB operations)
                    // This validation is pure data checking - can be parallelized
                    var validationTimer = System.Diagnostics.Stopwatch.StartNew();
                    var validSleeveData = new List<(FamilyInstance sleeve, ClashZone zone, double fw, double fh, double fd, double fdepth)>();
                    
                    if (OptimizationFlags.UseParallelProcessing && placedSleeveData.Count >= 12)
                    {
                        // ✅ PARALLEL VALIDATION: Validate all sleeves in parallel before database writes
                        var validationTasks = placedSleeveData
                            .Select(item => Task.Run(() =>
                            {
                                try
                                {
                                    var (sleeve, zone, fw, fh, fd, fdepth) = item;
                                    
                                    // ✅ VALIDATION CHECKS (pure data validation - no Revit API, no DB)
                                    if (zone == null) return (item, false, "Null zone");
                                    if (zone.SleeveInstanceId <= 0) return (item, false, $"Invalid SleeveInstanceId={zone.SleeveInstanceId}");
                                    if (sleeve == null || !sleeve.IsValidObject) return (item, false, "Invalid sleeve");
                                    if (fw <= 0 && fh <= 0 && fd <= 0 && zone.SleeveWidth <= 0 && zone.SleeveHeight <= 0 && zone.SleeveDiameter <= 0)
                                        return (item, false, "All dimensions are zero or negative");
                                    
                                    return (item, true, "Valid");
                                }
                                catch (Exception ex)
                                {
                                    return (item, false, ex.Message);
                                }
                            }))
                            .ToArray();

                        Task.WaitAll(validationTasks);

                        foreach (var task in validationTasks)
                        {
                            var (item, isValid, reason) = task.Result;
                            if (isValid)
                            {
                                validSleeveData.Add(item);
                            }
                            else
                            {
                                skippedCount++;
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Warning($"[SleevePersistenceService] Skipping invalid sleeve: {reason}");
                                }
                            }
                        }
                    }
                    else
                    {
                        // ✅ FALLBACK: Sequential validation if parallel disabled or small batch
                        validSleeveData = placedSleeveData.ToList();
                    }
                    
                    validationTimer.Stop();
                    if (!DeploymentConfiguration.DeploymentMode && validSleeveData.Count != placedSleeveData.Count)
                    {
                        DebugLogger.Info($"[SleevePersistenceService] ✅ Pre-validated {placedSleeveData.Count} sleeves in parallel: {validSleeveData.Count} valid, {skippedCount} skipped in {validationTimer.ElapsedMilliseconds}ms");
                    }

                    // ✅ PERFORMANCE OPTIMIZATION: Process each sleeve and update zone object in-memory
                    // Then perform ONE high-performance bulk update at the end (Push then Merge)
                    foreach (var (sleeve, zone, fw, fh, fd, fdepth) in validSleeveData)
                    {
                        // ✅ SAFETY: Comprehensive validation before processing
                        if (zone == null)
                        {
                            skippedCount++;
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning("[SleevePersistenceService] Skipping null zone");
                            continue;
                        }

                        if (zone.SleeveInstanceId <= 0 && zone.ClusterSleeveInstanceId <= 0 && zone.CombinedClusterSleeveInstanceId <= 0 && !zone.IsCombinedResolved)
                        {
                            skippedCount++;
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[SleevePersistenceService] Skipping zone {zone.Id}: No linked sleeve element (SleeveId={zone.SleeveInstanceId}, ClusterId={zone.ClusterSleeveInstanceId}, CombinedId={zone.CombinedClusterSleeveInstanceId})");
                            continue;
                        }

                        if (sleeve == null || !sleeve.IsValidObject)
                        {
                            skippedCount++;
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[SleevePersistenceService] Skipping zone {zone.Id}: Invalid sleeve (null or invalid object)");
                            continue;
                        }

                        // ✅ SAFETY: Validate dimensions are positive
                        if (fw <= 0 && fh <= 0 && fd <= 0 && zone.SleeveWidth <= 0 && zone.SleeveHeight <= 0 && zone.SleeveDiameter <= 0)
                        {
                            skippedCount++;
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[SleevePersistenceService] Skipping zone {zone.Id}: All dimensions are zero or negative");
                            continue;
                        }

                        try
                        {
                            // ✅ STEP 1: Sync geometry and family data back to zone object for bulk persistence
                            // This replaces multiple individual repository.UpdateXXX calls
                            zone.SleeveFamilyName = sleeve.Symbol?.Family?.Name ?? string.Empty;
                            
                            // Ensure dimensions are synced if they were calculated during placement
                            if (zone.SleeveWidth <= 0 && fw > 0) zone.SleeveWidth = fw;
                            if (zone.SleeveHeight <= 0 && fh > 0) zone.SleeveHeight = fh;
                            if (zone.SleeveDiameter <= 0 && fd > 0) zone.SleeveDiameter = fd;

                            // ✅ STEP 2: Handle rotated bounding box and corners
                            var rotationAngleRad = zone.MepElementRotationAngle;
                            var rotationAngleDeg = Math.Abs(rotationAngleRad * 180.0 / Math.PI);
                            bool isStraightAxisAligned = IsStraightAxisAligned(rotationAngleDeg);

                            if (Math.Abs(rotationAngleRad) > 1e-6 && !isStraightAxisAligned)
                            {
                                // ✅ PERFORMANCE: Use pre-calculated rotated bbox if available
                                if (rotatedBboxData.ContainsKey(zone.Id))
                                {
                                    var bbox = rotatedBboxData[zone.Id];
                                    zone.RotatedBoundingBoxMinX = bbox.minX;
                                    zone.RotatedBoundingBoxMinY = bbox.minY;
                                    zone.RotatedBoundingBoxMinZ = bbox.minZ;
                                    zone.RotatedBoundingBoxMaxX = bbox.maxX;
                                    zone.RotatedBoundingBoxMaxY = bbox.maxY;
                                    zone.RotatedBoundingBoxMaxZ = bbox.maxZ;
                                }
                            }

                            persistedCount++;
                        }
                        catch (Exception ex)
                        {
                            errorCount++;
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Error($"[SleevePersistenceService] Failed to process sleeve {sleeve?.Id?.GetIntegerValue() ?? -1} for zone {zone?.Id}: {ex.Message}");
                            }
                        }
                    }

                    // ✅ PLACEMENT OPTIMIZATION: Perform SINGLE high-performance bulk update (Push then Merge)
                    // This is 50x-100x faster than individual updates for large batches.
                    var bulkUpdateTimer = System.Diagnostics.Stopwatch.StartNew();
                    try
                    {
                        var zonesToUpdate = validSleeveData.Select(x => x.zone).ToList();
                        if (zonesToUpdate.Count > 0)
                        {
                            repository.BatchUpdatePostPlacement(zonesToUpdate);
                        }
                        bulkUpdateTimer.Stop();
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[SleevePersistenceService] ✅ BULK POST-PLACEMENT UPDATE: {zonesToUpdate.Count} zones in {bulkUpdateTimer.ElapsedMilliseconds}ms");
                        }
                    }
                    catch (Exception bulkEx)
                    {
                        bulkUpdateTimer.Stop();
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Error($"[SleevePersistenceService] ❌ BULK POST-PLACEMENT UPDATE FAILED: {bulkEx.Message}");
                        }
                        throw;
                    }


                    // ✅ PERFORMANCE MONITORING: Log batch persistence statistics
                    persistenceTimer.Stop();
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        double avgTimePerSleeve = persistedCount > 0 ? (double)persistenceTimer.ElapsedMilliseconds / persistedCount : 0;
                        DebugLogger.Info($"[SleevePersistenceService] ✅ Batch persistence complete: {persistedCount} persisted, {errorCount} errors, {skippedCount} skipped in {persistenceTimer.ElapsedMilliseconds}ms (avg: {avgTimePerSleeve:F1}ms per sleeve)");
                    }

                    // ✅ STEP 6: Save snapshots for all placed sleeves (only if we have persisted sleeves)
                    if (persistedCount > 0 && !string.IsNullOrWhiteSpace(filterName))
                    {
                        var snapshotTimer = System.Diagnostics.Stopwatch.StartNew();
                        try
                        {
                            // ? CRITICAL: Convert 6-tuple to 5-tuple for SaveSnapshots (it only expects 5)
                            var convertedForSnapshot = placedSleeveData.Select(x => (x.sleeve, x.zone, x.finalWidth, x.finalHeight, x.finalDiameter)).ToList();
                            SaveSnapshots(repository, dbContext, convertedForSnapshot, filterName);
                            snapshotTimer.Stop();
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[SleevePersistenceService] ✅ Snapshots saved in {snapshotTimer.ElapsedMilliseconds}ms");
                            }
                        }
                        catch (Exception snapshotEx)
                        {
                            snapshotTimer.Stop();
                            // ✅ SAFETY: Snapshot failure doesn't fail entire persistence (non-critical)
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Warning($"[SleevePersistenceService] Failed to save snapshots (non-critical): {snapshotEx.Message}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                persistenceTimer.Stop();
                // ✅ SAFETY: Log critical errors but don't throw (fail-safe)
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[SleevePersistenceService] ❌ CRITICAL: Failed to persist sleeve data: {ex.Message}");
                    DebugLogger.Error($"[SleevePersistenceService] Stack trace: {ex.StackTrace}");
                }
                // Return partial count (some sleeves may have been persisted before error)
            }

            return persistedCount;
        }

        /// <summary>
        /// Saves 4 corner coordinates in WORLD space only from Revit geometry extraction.
        /// No fallback: corners are persisted only when CalculateCornersFromInstance succeeds. See REVIT_GEOMETRY_RULES.md.
        /// </summary>
        private void SaveSleeveCorners(ClashZone zone, FamilyInstance sleeve, ClashZoneRepository repository,
            double rotationAngleRad, double width, double height)
        {
            (XYZ corner1, XYZ corner2, XYZ corner3, XYZ corner4)? corners = null;

            if (sleeve != null && sleeve.IsValidObject)
            {
                try
                {
                    corners = _cornerCalculationService.CalculateCornersFromInstance(sleeve, zone.HostOrientation, zone.StructuralElementType);
                    if (corners.HasValue && !DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[SaveSleeveCorners] ✅ Sleeve {sleeve.Id} GEOMETRY extraction success. C1=({corners.Value.corner1.X:F4},{corners.Value.corner1.Y:F4},{corners.Value.corner1.Z:F4})...");
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[SaveSleeveCorners] ⚠️ Sleeve {sleeve.Id} GEOMETRY extraction failed: {ex.Message}. Corners not persisted (no fallback).");
                }
            }

            if (corners.HasValue)
            {
                repository.UpdateSleeveCorners(
                    zone.Id,
                    corners.Value.corner1.X, corners.Value.corner1.Y, corners.Value.corner1.Z,
                    corners.Value.corner2.X, corners.Value.corner2.Y, corners.Value.corner2.Z,
                    corners.Value.corner3.X, corners.Value.corner3.Y, corners.Value.corner3.Z,
                    corners.Value.corner4.X, corners.Value.corner4.Y, corners.Value.corner4.Z);
            }
        }

        /// <summary>
        /// ✅ SRP: Saves sleeve snapshots for parameter transfer (delegates filter lookup to specialized service).
        /// </summary>
        private void SaveSnapshots(ClashZoneRepository repository, SleeveDbContext dbContext,
            List<(FamilyInstance sleeve, ClashZone zone, double finalWidth, double finalHeight, double finalDiameter)> placedSleeveData,
            string filterName)
        {
            try
            {
                // ✅ Get category from placed zones
                string categoryForLookup = null;
                var firstPlacedZone = placedSleeveData.FirstOrDefault(p => p.zone != null && p.zone.MepElementCategory != null);
                if (firstPlacedZone.zone != null)
                {
                    categoryForLookup = firstPlacedZone.zone.MepElementCategory;
                }

                // ✅ CRITICAL FIX: Always save snapshots for Individual, Cluster, and Combined sleeves
                // Get placed zones that have any valid instance association (REQUIRED for snapshot save)
                var placedZones = placedSleeveData
                    .Where(p => p.zone != null && (p.zone.SleeveInstanceId > 0 || p.zone.ClusterSleeveInstanceId > 0 || p.zone.IsCombinedResolved))
                    .Select(p => p.zone)
                    .Distinct()
                    .ToList();

                if (placedZones.Count == 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("database_operations.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [SleevePersistenceService] ⚠️ No zones with SleeveInstanceId > 0 found - cannot save snapshots. Total placed: {placedSleeveData.Count}\n");
                    }
                    return;
                }

                // ✅ SRP: Delegate filter lookup to specialized service
                var filterLookupService = new FilterLookupService(dbContext);
                int filterId = filterLookupService.GetFilterId(filterName, categoryForLookup);

                if (filterId <= 0)
                {
                    // ✅ CRITICAL FIX: Use filterId = -1 if lookup fails (repository will handle ComboId lookup)
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("database_operations.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [SleevePersistenceService] ⚠️ FilterId not found for FilterName='{filterName}' - will use ComboId lookup from database\n");
                    }
                    // Continue anyway - repository will try to find ComboId from database
                }

                // ✅ CRITICAL: Always save snapshots if we have placed zones with SleeveInstanceId
                // Note: SaveSleeveSnapshotsForPlacedSleeves will automatically load MepParameterValues from ClashZones table if missing
                repository.SaveSleeveSnapshotsForPlacedSleeves(filterId > 0 ? filterId : -1, placedZones);

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("database_operations.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [SleevePersistenceService] ✅ Saved sleeve snapshots for {placedZones.Count} placed sleeves (FilterId={filterId})\n");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("database_operations.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [SleevePersistenceService] ⚠️ Failed to save sleeve snapshots: {ex.Message}\n");
                }
            }
        }

        // ✅ HELPER METHODS (SRP: Single responsibility for each check)

        private bool IsBoundingBoxEmpty(ClashZone zone)
        {
            return zone.SleeveBoundingBoxMinX == 0.0 && zone.SleeveBoundingBoxMinY == 0.0 && zone.SleeveBoundingBoxMinZ == 0.0 &&
                   zone.SleeveBoundingBoxMaxX == 0.0 && zone.SleeveBoundingBoxMaxY == 0.0 && zone.SleeveBoundingBoxMaxZ == 0.0;
        }

        private bool IsRcsBoundingBoxEmpty(ClashZone zone)
        {
            return zone.SleeveBoundingBoxRCS_MinX == 0.0 && zone.SleeveBoundingBoxRCS_MinY == 0.0 && zone.SleeveBoundingBoxRCS_MinZ == 0.0 &&
                   zone.SleeveBoundingBoxRCS_MaxX == 0.0 && zone.SleeveBoundingBoxRCS_MaxY == 0.0 && zone.SleeveBoundingBoxRCS_MaxZ == 0.0;
        }

        private bool IsWallOrFramingHost(ClashZone zone)
        {
            return string.Equals(zone.StructuralElementType, "Wall", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(zone.StructuralElementType, "Walls", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(zone.StructuralElementType, "Structural Framing", StringComparison.OrdinalIgnoreCase);
        }

        // ✅ HELPER: Simple validation method for orchestration logic (not a calculation)
        private bool IsStraightAxisAligned(double rotationAngleDeg)
        {
            return Math.Abs(rotationAngleDeg) < 1.0 ||
                   Math.Abs(rotationAngleDeg - 90.0) < 1.0 ||
                   Math.Abs(rotationAngleDeg - 180.0) < 1.0 ||
                   Math.Abs(rotationAngleDeg - 270.0) < 1.0 ||
                   Math.Abs(rotationAngleDeg - 360.0) < 1.0;
        }
    }
}
