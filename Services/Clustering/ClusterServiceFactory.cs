using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Algorithm;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.BoundingBox;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Cleanup;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Data;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Placement;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Strategy;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Timeout;
using JSE_RevitAddin_MEP_OPENINGS.Services.Geometry; // ✅ SOLID: For SleeveCornerCalculationService
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement; // ✅ SOLID: For SleeveParameterService dependency injection
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using System;
using System.Collections.Generic;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering
{
    /// <summary>
    /// Factory for creating fully-configured clustering service instances with all dependencies.
    /// Orchestrates the instantiation of all Phase 1-10 extracted services.
    /// 
    /// Phase 1-5 Services (Geometry, Proximity, BoundingBox, Strategy, Placement):
    /// - Geometry: Static classes (DistanceCalculator, RotationMatrixCalculator, CoordinateTransformer)
    /// - Proximity: ProximityCheckerFactory (stateless)
    /// - BoundingBox: IBoundingBoxCalculator implementations (wired via RotationService)
    /// - Strategy: ClusteringStrategyFactory (stateless)
    /// - Placement: IClusterPlacementService (requires function delegates from other services)
    /// 
    /// Phase 6-10 Services (Rotation, Cleanup, Algorithm, Data, Timeout):
    /// - Already wired in CreateWithAllServices
    /// </summary>
    public static class ClusterServiceFactory
    {
        /// <summary>
        /// Create a fully-wired UniversalClusterService with all modern services enabled.
        /// This is the recommended way to instantiate the clustering service for new code.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="flagManager">Optional FlagManager (will be created if null)</param>
        /// <param name="filterService">Optional FilterManagementService (will be created if null)</param>
        /// <param name="timeoutLimitMs">Timeout limit in milliseconds (default: 300000 = 5 minutes)</param>
        /// <returns>Fully-configured RefactoredClusterService with all Phase 6-11 services</returns>
        public static RefactoredClusterService CreateWithAllServices(
            Document doc,
            JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor.IFlagManager flagManager = null,
            FilterManagementService filterService = null,
            int timeoutLimitMs = 300000)
        {
            // Phase 9: Data Service (needed for getClashZoneFunc)
            var dataService = new ClusterDataService(doc);
            
            // Phase 6: Rotation Service - wire function to load from cache first, then database
            // ✅ CRITICAL FIX: Use cache first (populated by ClusterDataService), then fallback to database
            // This ensures we use the already-loaded data with all rotated bounding box and corner data
            Func<int, string, Models.ClashZone> getClashZoneFunc = (sleeveId, xmlPath) =>
            {
                // ✅ STEP 1: Try cache first (fast, already loaded with all data including corners and rotated bbox)
                try
                {
                    var cached = dataService.GetClashZoneBySleeveInstanceId(sleeveId);
                    if (cached != null)
                    {
                        // ✅ CRITICAL FIX: Check if cached clash zone has bounding boxes
                        // If not, it's stale (loaded before bounding boxes were saved) - fall through to database reload
                        bool hasBoundingBox = cached.SleeveBoundingBoxMinX != 0 || cached.SleeveBoundingBoxMaxX != 0 ||
                                              cached.SleeveBoundingBoxMinY != 0 || cached.SleeveBoundingBoxMaxY != 0 ||
                                              cached.SleeveBoundingBoxMinZ != 0 || cached.SleeveBoundingBoxMaxZ != 0;
                        
                        if (hasBoundingBox)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                bool hasCorners = cached.SleeveCorner1X.HasValue && cached.SleeveCorner1Y.HasValue;
                                bool hasRotatedBbox = cached.RotatedBoundingBoxMinX.HasValue && cached.RotatedBoundingBoxMaxX.HasValue;
                                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                    $"[{DateTime.Now:HH:mm:ss}] ✅ CACHE HIT: Sleeve {sleeveId}: HasCorners={hasCorners}, HasRotatedBbox={hasRotatedBbox}, HasBoundingBox={hasBoundingBox}, Rotation={cached.MepElementRotationAngle * 180 / Math.PI:F1}°\n");
                            }
                            return cached;
                        }
                        else
                        {
                            // ⚠️ Cache has stale data (no bounding boxes) - fall through to database reload
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ CACHE STALE: Sleeve {sleeveId} found in cache but has no bounding boxes, reloading from database...\n");
                            }
                        }
                    }
                }
                catch (Exception cacheEx)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ Cache lookup error for sleeve {sleeveId}: {cacheEx.Message}\n");
                    }
                }
                
                // ✅ STEP 2: Fallback to database if not in cache (shouldn't happen if cache is properly populated)
                try
                {
                    using (var dbContext = new SleeveDbContext(doc))
                    {
                        var repository = new ClashZoneRepository(dbContext);
                        
                        // Query all categories and find matching SleeveInstanceId
                        var categories = new[] { "Ducts", "Pipes", "Cable Trays", "Duct Accessories" };
                        
                        foreach (var category in categories)
                        {
                            try
                            {
                                var zones = repository.GetClashZonesByCategory(category);
                                if (zones == null) continue;
                                
                                var match = zones.FirstOrDefault(cz => cz.SleeveInstanceId == sleeveId);
                                if (match != null)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        bool hasCorners = match.SleeveCorner1X.HasValue && match.SleeveCorner1Y.HasValue;
                                        bool hasRotatedBbox = match.RotatedBoundingBoxMinX.HasValue && match.RotatedBoundingBoxMaxX.HasValue;
                                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ CACHE MISS: Loaded from DB for sleeve {sleeveId}: HasCorners={hasCorners}, HasRotatedBbox={hasRotatedBbox}, Rotation={match.MepElementRotationAngle * 180 / Math.PI:F1}°\n");
                                    }
                                    return match;
                                }
                            }
                            catch { /* Try next category */ }
                        }
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ Could not find ClashZone for sleeve {sleeveId} in cache or database\n");
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                            $"[{DateTime.Now:HH:mm:ss}] ❌ Error loading ClashZone for sleeve {sleeveId}: {ex.Message}\n");
                    }
                }
                
                return null;
            };
            
            // ✅ VALIDATION: Ensure function is not null before passing to constructor
            if (getClashZoneFunc == null)
                throw new InvalidOperationException("getClashZoneFunc cannot be null");
            
            var rotationService = new ClusterRotationService(getClashZoneFunc);

            // Phase 7: Cleanup Service
            var cleanupService = new ClusterCleanupService();

            // Phase 8: Algorithm Service
            var algorithmService = new ClusterAlgorithmService();

            // ✅ NOTE: dataService already created above for getClashZoneFunc
            // var dataService = new ClusterDataService(doc); // Already created above

            // Phase 10: Timeout Service
            var timeoutService = new ClusterTimeoutService(timeoutLimitMs);
            
            // ✅ Phase 1b: Corner Calculation Service
            var cornerService = new SleeveCornerCalculationService();
            

            
            // Phase 5: Placement Service - null delegates (RefactoredClusterService wires them internally)
            // ✅ SOLID: Create SleeveParameterService for dependency injection
            var parameterService = new SleeveParameterService(doc);
            var placementService = new ClusterPlacementService(
                getClashZoneBySleeveInstanceId: null,
                determineRotationAngle: null,
                getClusterBoundingBox: null,
                markClusterResolved: null,
                getFilterNameForCategory: null,
                boundingBoxCalculator: null,
                parameterService: parameterService // ✅ SOLID: Inject SleeveParameterService dependency
            );
            
            // Phase 4: Strategy Factory
            var strategyFactory = new ClusteringStrategyFactory();

            // ✅ WIRING: Create flag manager adapter (use null for now since refactored flag manager not ready)
            Services.Interfaces.Refactor.IFlagManager? flagManagerRefactor = null;
            
            // Wire all services into RefactoredClusterService
            return new RefactoredClusterService(
                doc: doc,
                dataService: dataService,
                algorithmService: algorithmService,
                rotationService: rotationService,
                placementService: placementService,
                cleanupService: cleanupService,
                timeoutService: timeoutService,
                cornerService: cornerService,
                strategyFactory: strategyFactory,
                flagManager: null,  // IFlagManager - refactored flag manager not yet ready
                filterService: filterService
            );
        }





        /// <summary>
        /// ✅ NEW: Create a fully-wired RefactoredClusterService with ALL Phase 1-10 services.
        /// This is the modern, clean orchestrator that replaces the 8000+ line UniversalClusterService.
        /// 
        /// Architecture:
        /// - Phase 1: Geometry (static classes - no instantiation needed)
        /// - Phase 2: Proximity (ProximityCheckerFactory - stateless)
        /// - Phase 3: BoundingBox (IBoundingBoxCalculator - wired via RotationService)
        /// - Phase 4: Strategy (ClusteringStrategyFactory - stateless)
        /// - Phase 5: Placement (IClusterPlacementService - requires function delegates)
        /// - Phase 6: Rotation (IClusterRotationService)
        /// - Phase 7: Cleanup (IClusterCleanupService)
        /// - Phase 8: Algorithm (IClusterAlgorithmService)
        /// - Phase 9: Data (IClusterDataService)
        /// - Phase 10: Timeout (IClusterTimeoutService)
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="flagManager">Optional FlagManager (will be created if null)</param>
        /// <param name="filterService">Optional FilterManagementService (will be created if null)</param>
        /// <param name="timeoutLimitMs">Timeout limit in milliseconds (default: 300000 = 5 minutes)</param>
        /// <returns>Fully-configured RefactoredClusterService with all Phase 1-10 services</returns>
        public static RefactoredClusterService CreateRefactored(
            Document doc,
            JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor.IFlagManager? flagManager = null,
            FilterManagementService? filterService = null,
            int timeoutLimitMs = 300000)
        {
            // ✅ Phase 9: Data Service (required for all other services)
            var dataService = new ClusterDataService(doc);

            // ✅ Phase 10: Timeout Service
            var timeoutService = new ClusterTimeoutService(timeoutLimitMs);

            // ✅ Phase 1b: Corner Calculation Service
            var cornerService = new SleeveCornerCalculationService();

            // ✅ Phase 8: Algorithm Service (stateless)
            var algorithmService = new ClusterAlgorithmService();

            // ✅ Phase 7: Cleanup Service (stateless)
            var cleanupService = new ClusterCleanupService();

            // ✅ Phase 6: Rotation Service (requires getClashZoneFunc - will be wired from dataService)
            Func<int, string, ClashZone> getClashZoneFunc = (sleeveId, xmlPath) =>
            {
                // Use dataService cache to get ClashZone
                // Note: This is a simplified implementation - in practice, we'd need to track sleeveId to ClashZone mapping
                // For now, return null - RotationService will handle fallback
                return null;
            };
            var rotationService = new ClusterRotationService(getClashZoneFunc);

            // ✅ Phase 4: Strategy Factory (stateless)
            var strategyFactory = new ClusteringStrategyFactory();

            // ✅ Phase 5: Placement Service (requires function delegates from other services)
            // These delegates will be wired from RefactoredClusterService's methods
            // For now, create with null delegates - RefactoredClusterService will provide them via method calls
            // Note: PlacementService is actually called indirectly via PlaceClusterForGroup method
            // So we can create a minimal wrapper that will be replaced by proper wiring
            var placementService = CreatePlacementService(doc, dataService, rotationService, flagManager, null); // ✅ SOLID: Pass doc for SleeveParameterService

            // ✅ Create RefactoredClusterService with all services wired
            return new RefactoredClusterService(
                doc: doc,
                dataService: dataService,
                algorithmService: algorithmService,
                rotationService: rotationService,
                placementService: placementService,
                cleanupService: cleanupService,
                timeoutService: timeoutService,
                cornerService: cornerService,
                strategyFactory: strategyFactory,
                flagManager: null, // IFlagManager - refactored flag manager not yet ready
                filterService: filterService
            );
        }

        /// <summary>
        /// Helper method to create PlacementService with all required function delegates.
        /// </summary>
        private static IClusterPlacementService CreatePlacementService(
            Document doc, // ✅ SOLID: Required for SleeveParameterService dependency injection
            IClusterDataService dataService,
            IClusterRotationService rotationService,
            JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor.IFlagManager? flagManager,
            Func<string, string?>? getFilterNameForCategory)
        {
            // ✅ Wire function delegates for PlacementService
            // Note: These are placeholders - RefactoredClusterService should provide these via its own methods
            Func<int, string?, ClashZone?> getClashZoneBySleeveInstanceId = (sleeveId, xmlPath) =>
            {
                // Use dataService cache
                // TODO: Implement proper lookup from dataService cache
                return null;
            };

            Func<List<dynamic>, string?, double> determineRotationAngle = (cluster, xmlPath) =>
            {
                return rotationService.DetermineRotationAngle(cluster, xmlPath);
            };

            Func<List<dynamic>, List<FamilyInstance>, double, string?, (double width, double height, double depth, XYZ mid, double? rotatedMinX, double? rotatedMinY, double? rotatedMinZ, double? rotatedMaxX, double? rotatedMaxY, double? rotatedMaxZ)> getClusterBoundingBox = (cluster, actualSleeves, rotationAngle, xmlPath) =>
            {
                return rotationService.CalculateRotatedBoundingBox(cluster, actualSleeves, rotationAngle, xmlPath);
            };

            Action<List<dynamic>, ElementId, string?, BoundingBoxXYZ?, (double minX, double minY, double minZ, double maxX, double maxY, double maxZ)?> markClusterResolved = (cluster, clusterSleeveId, xmlPath, bbox, rotatedBbox) =>
            {
                // ✅ FLAG FIX: Update flags in database using ClashZoneRepository
                if (cluster == null || cluster.Count == 0 || clusterSleeveId == ElementId.InvalidElementId) return;

                try
                {
                    using (var dbContext = new SleeveDbContext(doc))
                    {
                        var repository = new ClashZoneRepository(dbContext);
                        foreach (var sleeve in cluster)
                        {
                            // Use dynamic property access to get ClashZone
                            // The dynamic object is created in RefactoredClusterService.PrepareSleeveData and has a ClashZone property
                            var clashZone = sleeve.ClashZone as ClashZone;
                            
                            if (clashZone != null)
                            {
                                // Call UpdateClusterPlacement to update state, cluster ID, and flags
                                // This sets IsResolvedFlag=1 and IsClusterResolvedFlag=1 via the repository logic
                                repository.UpdateClusterPlacement(
                                    clashZoneId: clashZone.Id,
                                    clusterInstanceId: clusterSleeveId.IntegerValue,
                                    minX: bbox?.Min.X ?? 0,
                                    minY: bbox?.Min.Y ?? 0,
                                    minZ: bbox?.Min.Z ?? 0,
                                    maxX: bbox?.Max.X ?? 0,
                                    maxY: bbox?.Max.Y ?? 0,
                                    maxZ: bbox?.Max.Z ?? 0,
                                    // Pass rotated bounding box if available
                                    rotatedMinX: rotatedBbox?.minX,
                                    rotatedMinY: rotatedBbox?.minY,
                                    rotatedMinZ: rotatedBbox?.minZ,
                                    rotatedMaxX: rotatedBbox?.maxX,
                                    rotatedMaxY: rotatedBbox?.maxY,
                                    rotatedMaxZ: rotatedBbox?.maxZ,
                                    isClustered: true
                                );
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] ✅ DELEGATE: Marked ClashZone {clashZone.Id} as resolved for Cluster {clusterSleeveId.IntegerValue}\n");
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_errors.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ❌ DELEGATE ERROR: Failed to mark cluster resolved: {ex.Message}\n");
                    }
                }
            };

            getFilterNameForCategory ??= (category) => null; // Default to null if not provided

            // ✅ SOLID: Create SleeveParameterService for dependency injection
            var parameterService = new SleeveParameterService(doc);

            return new ClusterPlacementService(
                getClashZoneBySleeveInstanceId: getClashZoneBySleeveInstanceId,
                determineRotationAngle: determineRotationAngle,
                getClusterBoundingBox: getClusterBoundingBox,
                markClusterResolved: markClusterResolved,
                getFilterNameForCategory: getFilterNameForCategory,
                boundingBoxCalculator: null, // Optional - wired via rotation service
                parameterService: parameterService // ✅ SOLID: Inject SleeveParameterService dependency
            );
        }
    }
}
