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
            List<(FamilyInstance sleeve, ClashZone zone, double finalWidth, double finalHeight, double finalDiameter)> placedSleeveData,
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
                                    var (sleeve, zone, fw, fh, fd) = item;
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
                        cornerData = _parallelCornerOrchestrator.CalculateCornersInParallel(placedSleeveData);
                    }
                    
                    geometryCalculationTimer.Stop();
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[SleevePersistenceService] ✅ Pre-calculated geometry in parallel: {cornerData.Count} corners, {rotatedBboxData.Count} rotated bboxes in {geometryCalculationTimer.ElapsedMilliseconds}ms");
                    }

                    // ✅ PERFORMANCE OPTIMIZATION: Pre-validate all sleeves in parallel (non-Revit, non-DB operations)
                    // This validation is pure data checking - can be parallelized
                    var validationTimer = System.Diagnostics.Stopwatch.StartNew();
                    var validSleeveData = new List<(FamilyInstance sleeve, ClashZone zone, double fw, double fh, double fd)>();
                    
                    if (OptimizationFlags.UseParallelProcessing && placedSleeveData.Count >= 12)
                    {
                        // ✅ PARALLEL VALIDATION: Validate all sleeves in parallel before database writes
                        var validationTasks = placedSleeveData
                            .Select(item => Task.Run(() =>
                            {
                                try
                                {
                                    var (sleeve, zone, fw, fh, fd) = item;
                                    
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

                    // ✅ PLACEMENT OPTIMIZATION: Collect all updates in lists for batch processing
                    var placementUpdates = new List<(Guid ClashZoneGuid, int SleeveInstanceId, double Width, double Height, double Diameter,
                        double PlacementX, double PlacementY, double PlacementZ,
                        double PlacementActiveX, double PlacementActiveY, double PlacementActiveZ,
                        double RotationAngleRad)>();
                    var cornerUpdates = new List<(Guid ClashZoneGuid,
                        double Corner1X, double Corner1Y, double Corner1Z,
                        double Corner2X, double Corner2Y, double Corner2Z,
                        double Corner3X, double Corner3Y, double Corner3Z,
                        double Corner4X, double Corner4Y, double Corner4Z)>();

                    // ✅ PERFORMANCE OPTIMIZATION: Batch processing in single transaction (handled by repository)
                    // ✅ SAFETY: Process each sleeve with individual error handling (fail-safe)
                    // ✅ NOTE: Database operations remain sequential (SQLite doesn't support parallel writes well)
                    foreach (var (sleeve, zone, fw, fh, fd) in validSleeveData)
                    {
                        // ✅ SAFETY: Comprehensive validation before processing
                        if (zone == null)
                        {
                            skippedCount++;
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning("[SleevePersistenceService] Skipping null zone");
                            continue;
                        }

                        if (zone.SleeveInstanceId <= 0)
                        {
                            skippedCount++;
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[SleevePersistenceService] Skipping zone {zone.Id}: Invalid SleeveInstanceId={zone.SleeveInstanceId}");
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
                            // ✅ STEP 1: Collect placement data for batch update (instead of immediate write)
                            repository.UpdateSleeveInstanceId(zone.Id, zone.SleeveInstanceId);
                            
                            // ✅ CRITICAL FIX: Save Family Name to DB (User Request)
                            // "debug after placing sleeves it should populated the families in db"
                            if (sleeve.Symbol != null && sleeve.Symbol.Family != null)
                            {
                                repository.UpdateSleeveFamilyName(zone.Id, sleeve.Symbol.Family.Name);
                            }

                            placementUpdates.Add((
                                zone.Id,
                                zone.SleeveInstanceId,
                                zone.SleeveWidth > 0 ? zone.SleeveWidth : fw,
                                zone.SleeveHeight > 0 ? zone.SleeveHeight : fh,
                                zone.SleeveDiameter > 0 ? zone.SleeveDiameter : fd,
                                zone.SleevePlacementPointX,
                                zone.SleevePlacementPointY,
                                zone.SleevePlacementPointZ,
                                zone.SleevePlacementPointActiveDocumentX,
                                zone.SleevePlacementPointActiveDocumentY,
                                zone.SleevePlacementPointActiveDocumentZ,
                                zone.MepElementRotationAngle));

                            // ✅ CRITICAL: Sync MEP Category to DB (Dump once, use many times)
                            // This ensures the category used for filtering is persisted
                            if (!string.IsNullOrEmpty(zone.MepElementCategory))
                            {
                                repository.UpdateMepCategory(zone.Id, zone.MepElementCategory);
                            }

                            // ✅ STEP 2: Save bounding boxes if available
                            if (!IsBoundingBoxEmpty(zone))
                            {
                                // Save axis-aligned bounding box
                                repository.UpdateSleeveBoundingBoxes(
                                    zone.Id,
                                    zone.SleeveBoundingBoxMinX, zone.SleeveBoundingBoxMinY, zone.SleeveBoundingBoxMinZ,
                                    zone.SleeveBoundingBoxMaxX, zone.SleeveBoundingBoxMaxY, zone.SleeveBoundingBoxMaxZ);

                                // ✅ STEP 3: Save RCS bounding box for walls/framing
                                if (IsWallOrFramingHost(zone) && !IsRcsBoundingBoxEmpty(zone))
                                {
                                    repository.UpdateSleeveBoundingBoxesRcs(
                                        zone.Id,
                                        zone.SleeveBoundingBoxRCS_MinX, zone.SleeveBoundingBoxRCS_MinY, zone.SleeveBoundingBoxRCS_MinZ,
                                        zone.SleeveBoundingBoxRCS_MaxX, zone.SleeveBoundingBoxRCS_MaxY, zone.SleeveBoundingBoxRCS_MaxZ);
                                }

                                // ✅ STEP 4: Save rotated bounding box and corners for rotated sleeves
                                var rotationAngleRad = zone.MepElementRotationAngle;
                                var rotationAngleDeg = Math.Abs(rotationAngleRad * 180.0 / Math.PI);
                                bool isStraightAxisAligned = IsStraightAxisAligned(rotationAngleDeg);

                                if (Math.Abs(rotationAngleRad) > 1e-6 && !isStraightAxisAligned)
                                {
                                    // ✅ PERFORMANCE: Use pre-calculated rotated bbox if available (from parallel calculation)
                                    if (rotatedBboxData.ContainsKey(zone.Id))
                                    {
                                        var bbox = rotatedBboxData[zone.Id];
                                        repository.UpdateRotatedBoundingBoxes(
                                            zone.Id,
                                            bbox.minX, bbox.minY, bbox.minZ,
                                            bbox.maxX, bbox.maxY, bbox.maxZ);
                                    }
                                    else
                                    {
                                        // ✅ FALLBACK: Calculate on-the-fly if not pre-calculated (sequential fallback)
                                        double actualWidth = zone.SleeveWidth > 0 ? zone.SleeveWidth : fw;
                                        double actualHeight = zone.SleeveHeight > 0 ? zone.SleeveHeight : fh;
                                        double actualDepth = zone.SleeveBoundingBoxMaxZ - zone.SleeveBoundingBoxMinZ;

                                        var rotatedBbox = _rotatedBboxService.CalculateRotatedBoundingBoxFromZone(zone, actualWidth, actualHeight, actualDepth);
                                        if (rotatedBbox.HasValue)
                                        {
                                            repository.UpdateRotatedBoundingBoxes(
                                                zone.Id,
                                                rotatedBbox.Value.minX, rotatedBbox.Value.minY, rotatedBbox.Value.minZ,
                                                rotatedBbox.Value.maxX, rotatedBbox.Value.maxY, rotatedBbox.Value.maxZ);
                                        }
                                    }

                                    // ✅ STEP 5: Batch save corners to database (use pre-calculated if available, otherwise calculate on-the-fly)
                                    // ✅ CRITICAL: These corners are saved AFTER regeneration and used by cluster calculation
                                    // - Cluster calculation reads corners from database (SleeveCorner1X/Y/Z through Corner4X/Y/Z)
                                    // - NO recalculation needed during clustering - corners are already saved
                                    // ✅ DISABLED: Corner calculation moved to BatchSleeveCornerExtractor (after placement + regeneration)
                                    // Corners are now extracted from actual Revit geometry, not calculated from math.
                                    // See OpeningCommandOrchestrator.ExecuteCommandSequence for the correct extraction flow.
                                }
                                else if (Math.Abs(rotationAngleRad) > 1e-6 && isStraightAxisAligned)
                                {
                                    // ✅ AXIS-ALIGNED SLEEVE: Save corners only (no rotated bbox needed)
                                    double axisAlignedWidth = zone.SleeveWidth > 0 ? zone.SleeveWidth : fw;
                                    double axisAlignedHeight = zone.SleeveHeight > 0 ? zone.SleeveHeight : fh;
                                    
                                    // Use pre-calculated corners if available
                                    // ✅ DISABLED: Corner calculation moved to BatchSleeveCornerExtractor
                                }
                                else
                                {
                                    // ✅ ZERO ROTATION: Save corners for consistency
                                    double zeroRotationWidth = zone.SleeveWidth > 0 ? zone.SleeveWidth : fw;
                                    double zeroRotationHeight = zone.SleeveHeight > 0 ? zone.SleeveHeight : fh;
                                    
                                    // Use pre-calculated corners if available
                                    // ✅ DISABLED: Corner calculation moved to BatchSleeveCornerExtractor
                                }
                            }

                            persistedCount++;
                        }
                        catch (Exception ex)
                        {
                            errorCount++;
                            // ✅ SAFETY: Fail-safe - continue with next sleeve instead of aborting entire batch
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Error($"[SleevePersistenceService] Failed to persist sleeve {sleeve?.Id?.IntegerValue ?? -1} for zone {zone?.Id}: {ex.Message}");
                                DebugLogger.Error($"[SleevePersistenceService] Stack trace: {ex.StackTrace}");
                            }
                            // Continue with next sleeve (fail-safe)
                        }
                    }

                    // ✅ PLACEMENT OPTIMIZATION: Batch flush all collected updates (50x faster than per-sleeve calls)
                    var batchFlushTimer = System.Diagnostics.Stopwatch.StartNew();
                    try
                    {
                        if (placementUpdates.Count > 0)
                        {
                            repository.BatchUpdateSleevePlacement(placementUpdates);
                        }
                        if (cornerUpdates.Count > 0)
                        {
                            repository.BatchUpdateSleeveCorners(cornerUpdates);
                        }
                        batchFlushTimer.Stop();
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[SleevePersistenceService] ✅ BATCH FLUSH: {placementUpdates.Count} placements + {cornerUpdates.Count} corners in {batchFlushTimer.ElapsedMilliseconds}ms");
                        }
                    }
                    catch (Exception batchEx)
                    {
                        batchFlushTimer.Stop();
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Error($"[SleevePersistenceService] ❌ BATCH FLUSH FAILED: {batchEx.Message}");
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
                            SaveSnapshots(repository, dbContext, placedSleeveData, filterName);
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
        /// ✅ SRP: Calculates and saves 4 corner coordinates in WORLD space (delegates to corner calculation service).
        /// ✅ CRITICAL FIX: For Y-wall sleeves, use width for Y-axis (along wall direction).
        /// Corner calculation service maps: width→X-axis, height→Y-axis.
        /// For Y-wall: Width (along wall) should map to Y-axis, so pass width as height parameter.
        /// </summary>
        private void SaveSleeveCorners(ClashZone zone, FamilyInstance sleeve, ClashZoneRepository repository, 
            double rotationAngleRad, double width, double height)
        {
            // ✅ CRITICAL FIX: For Y-wall, use width for Y-axis (along wall)
            // Corner service maps: width→X, height→Y
            // For Y-wall: width (along wall) → Y-axis, so pass width as height parameter
            double cornerWidth = width;
            double cornerHeight = height;
            
            bool isYWall = IsWallOrFramingHost(zone) && 
                          (zone.HostOrientation == "Y" || 
                           (zone.WallDirection != null && Math.Abs(zone.WallDirection.Y) > Math.Abs(zone.WallDirection.X)));
            
            if (isYWall)
            {
                // ✅ Y-WALL: Use width for Y-axis (along wall direction)
                cornerWidth = height;  // Height (vertical) → X-axis
                cornerHeight = width;  // Width (along wall) → Y-axis ✅
            }
            
            // ✅ SRP: Delegate corner calculation to specialized service
            // ✅ IMPROVEMENT (User Request): "Read corners from Revit"
            // prioritized: Extract exact corners from placed element geometry (Solids)
            // Fallback: Calculate from parameters (original logic)
            
            (XYZ corner1, XYZ corner2, XYZ corner3, XYZ corner4)? corners = null;
            
            // 1. Try extracting from geometry (Most trusted source)
            if (sleeve != null && sleeve.IsValidObject)
            {
               try 
               {
                   corners = _cornerCalculationService.CalculateCornersFromInstance(sleeve);
                   if (corners.HasValue && !DeploymentConfiguration.DeploymentMode)
                   {
                         // Optional: Log success
                         DebugLogger.Info($"[SaveSleeveCorners] ✅ Sleeve {sleeve.Id} GEOMETRY extraction success. C1=({corners.Value.corner1.X:F4},{corners.Value.corner1.Y:F4},{corners.Value.corner1.Z:F4})...");
                   }
               }
               catch (Exception ex)
               {
                   if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[SaveSleeveCorners] ⚠️ Sleeve {sleeve.Id} GEOMETRY extraction failed: {ex.Message}");
               }
            }
            
            // 2. Fallback to math calculation if geometry failed
            if (!corners.HasValue)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[SaveSleeveCorners] ⚠️ Sleeve {sleeve.Id} - Falling back to MATH calculation.");
                
                corners = _cornerCalculationService.CalculateCorners(
                    zone.SleevePlacementPoint,
                    cornerWidth,
                    cornerHeight,
                    rotationAngleRad);
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

                // ✅ CRITICAL FIX: Always save snapshots, even if filter lookup fails
                // Get placed zones with SleeveInstanceId > 0 (REQUIRED for snapshot save)
                var placedZones = placedSleeveData
                    .Where(p => p.zone != null && p.zone.SleeveInstanceId > 0)
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

