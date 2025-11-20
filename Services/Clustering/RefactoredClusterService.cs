using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Algorithm;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.BoundingBox;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Cleanup;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Data;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Placement;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Proximity;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Strategy;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Timeout;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Safety;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering
{
    /// <summary>
    /// ✅ REFACTORED CLUSTER SERVICE: Clean orchestrator for all Phase 1-10 services.
    /// Replaces the 8000+ line UniversalClusterService with clean delegation to extracted services.
    /// 
    /// Architecture:
    /// - Phase 1: Geometry (DistanceCalculator, RotationMatrixCalculator, CoordinateTransformer)
    /// - Phase 2: Proximity (ProximityCheckerFactory, IProximityChecker implementations)
    /// - Phase 3: BoundingBox (IBoundingBoxCalculator implementations)
    /// - Phase 4: Strategy (IClusteringStrategy implementations, ClusteringStrategyFactory)
    /// - Phase 5: Placement (IClusterPlacementService)
    /// - Phase 6: Rotation (IClusterRotationService)
    /// - Phase 7: Cleanup (IClusterCleanupService)
    /// - Phase 8: Algorithm (IClusterAlgorithmService)
    /// - Phase 9: Data (IClusterDataService)
    /// - Phase 10: Timeout (IClusterTimeoutService)
    /// </summary>
    public class RefactoredClusterService
    {
        // Phase 1-5 Services (Geometry, Proximity, BoundingBox, Strategy, Placement)
        private readonly ClusteringStrategyFactory _strategyFactory;
        private readonly IClusterPlacementService _placementService;
        
        // Phase 6-10 Services (Rotation, Cleanup, Algorithm, Data, Timeout)
        private readonly IClusterRotationService _rotationService;
        private readonly IClusterCleanupService _cleanupService;
        private readonly IClusterAlgorithmService _algorithmService;
        private readonly IClusterDataService _dataService;
        private readonly IClusterTimeoutService _timeoutService;
        
        // Supporting Services
        private readonly FlagManager _flagManager;
        private readonly FilterManagementService _filterService;
        private readonly Document _doc;
        
        // Internal state
        private string? _filterName;
        private readonly Dictionary<int, List<Guid>> _clusterToClashZoneIds = new Dictionary<int, List<Guid>>();

        /// <summary>
        /// Constructor with dependency injection for all Phase 1-10 services.
        /// </summary>
        public RefactoredClusterService(
            Document doc,
            IClusterDataService dataService,
            IClusterAlgorithmService algorithmService,
            IClusterRotationService rotationService,
            IClusterPlacementService placementService,
            IClusterCleanupService cleanupService,
            IClusterTimeoutService timeoutService,
            ClusteringStrategyFactory? strategyFactory = null,
            FlagManager? flagManager = null,
            FilterManagementService? filterService = null)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _dataService = dataService ?? throw new ArgumentNullException(nameof(dataService));
            _algorithmService = algorithmService ?? throw new ArgumentNullException(nameof(algorithmService));
            _rotationService = rotationService ?? throw new ArgumentNullException(nameof(rotationService));
            _placementService = placementService ?? throw new ArgumentNullException(nameof(placementService));
            _cleanupService = cleanupService ?? throw new ArgumentNullException(nameof(cleanupService));
            _timeoutService = timeoutService ?? throw new ArgumentNullException(nameof(timeoutService));
            
            _strategyFactory = strategyFactory ?? new ClusteringStrategyFactory();
            _flagManager = flagManager ?? new FlagManager(doc);
            _filterService = filterService ?? new FilterManagementService(
                doc,
                msg => { if (!DeploymentConfiguration.DeploymentMode) DebugLogger.Info(msg); },
                msg => { if (!DeploymentConfiguration.DeploymentMode) DebugLogger.Error(msg); });
        }

        /// <summary>
        /// Main entry point: Cluster sleeves for a specific category.
        /// Clean orchestration of all Phase 1-10 services.
        /// </summary>
        public (int placedCount, int deletedCount) ClusterSleeves(
            Document doc,
            string targetCategory,
            UIDocument? uiDoc = null,
            string? xmlFilePath = null,
            string? filterName = null,
            List<FamilyInstance>? placedClusterSleevesOut = null,
            bool isPath1Replay = false,
            int? comboId = null,
            int? filterId = null)
        {
            _filterName = filterName;
            int placedCount = 0;
            int deletedCount = 0;
            var placedClusters = new List<FamilyInstance>();

            try
            {
                // ✅ STEP 1: Path 1 Replay - Load from database if available
                if (isPath1Replay && comboId.HasValue && filterId.HasValue)
                {
                    var path1Result = HandlePath1Replay(doc, comboId.Value, filterId.Value, targetCategory, uiDoc, placedClusterSleevesOut, xmlFilePath);
                    if (path1Result.hasData)
                        return (path1Result.placedCount, path1Result.deletedCount);
                }

                // ✅ STEP 2: Load clash zones (Phase 9: Data Service)
                var allClashZones = _dataService.LoadClashZonesFromRegularXml(xmlFilePath, targetCategory, doc);
                if (allClashZones == null || allClashZones.Count == 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Log($"[RefactoredClusterService] No clash zones loaded for category '{targetCategory}'");
                    return (0, 0);
                }

                // ✅ STEP 3: Populate cache (Phase 9: Data Service)
                _dataService.LoadClashZoneCacheFromLoadedClashZones(allClashZones, targetCategory);

                // ✅ STEP 4: Filter and prepare sleeves for clustering
                var filteredClashZones = allClashZones
                    .Where(cz => cz.SleeveInstanceId > 0)
                    .Where(cz => !cz.IsClusterResolved)
                    .Where(cz => string.IsNullOrEmpty(targetCategory) || 
                                string.Equals(cz.MepElementCategory, targetCategory, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (filteredClashZones.Count == 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Log($"[RefactoredClusterService] No sleeves to cluster after filtering");
                    return (0, 0);
                }

                // ✅ STEP 5: Validate transaction
                if (!doc.IsModifiable)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[RefactoredClusterService] Document is not modifiable - transaction must be started by caller");
                    throw new InvalidOperationException("Document must be in a transaction before calling ClusterSleeves");
                }

                // ✅ STEP 6: Get tolerance from settings
                double toleranceDist = GetToleranceFromSettings(targetCategory);

                // ✅ STEP 7: Prepare sleeve data (create dynamic objects with ClashZone references)
                var rawSleeves = PrepareSleeveData(filteredClashZones, allClashZones);

                // ✅ STEP 8: Group sleeves by host type, system type, and orientation
                var sleeveGroups = rawSleeves.GroupBy(sleeve => new SleeveGroupKey(
                    sleeve.HostType,
                    sleeve.Category,
                    sleeve.Orientation
                ));

                // ✅ STEP 9: Start timeout protection (Phase 10: Timeout Service)
                _timeoutService.StartTimer();

                // ✅ STEP 10: Form clusters using algorithm service (Phase 8: Algorithm Service)
                var clustersByGroup = _algorithmService.FormClusters(sleeveGroups, toleranceDist, doc, enableParallel: true);

                // ✅ STEP 11: Check timeout after clustering
                if (_timeoutService.IsTimedOut())
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[RefactoredClusterService] ⏱ TIMEOUT: Clustering exceeded {_timeoutService.TimeoutLimitMs / 1000} second limit");
                    _timeoutService.ShowTimeoutWarning("during FormClusters");
                    return (placedCount, deletedCount);
                }

                // ✅ STEP 12: Process each cluster group
                int clusterProcessedCount = 0;
                foreach (var groupEntry in clustersByGroup)
                {
                    // Check timeout every 5 clusters
                    clusterProcessedCount++;
                    if (clusterProcessedCount % 5 == 0 && _timeoutService.IsTimedOut())
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error($"[RefactoredClusterService] ⏱ TIMEOUT: Exceeded limit after processing {clusterProcessedCount} clusters");
                        _timeoutService.ShowTimeoutWarning($"after processing {clusterProcessedCount} clusters");
                        break;
                    }

                    var groupKey = groupEntry.Key;
                    var clusters = groupEntry.Value;

                    // ✅ STEP 13: Select clustering strategy (Phase 4: Strategy Factory)
                    var strategy = _strategyFactory.GetStrategy(groupKey, clusters.FirstOrDefault() ?? new List<dynamic>());

                    // Process each cluster
                    foreach (var cluster in clusters)
                    {
                        if (cluster.Count <= 1) continue; // Skip individual sleeves

                        try
                        {
                // ✅ STEP 14: Place cluster sleeve (Phase 5: Placement Service + Phase 6: Rotation Service + Phase 3: BoundingBox)
                // Note: Placement service needs to be wired with functions from rotation service and data service
                var placementResult = PlaceClusterForGroup(
                    doc,
                    cluster,
                    groupKey,
                    targetCategory,
                    xmlFilePath);

                            if (placementResult.success)
                            {
                                placedCount += placementResult.placedCount;
                                deletedCount += placementResult.deletedCount;
                                
                                if (placementResult.placedClusterSleeve != null && placementResult.capturedClusterSleeveId.HasValue)
                                {
                                    placedClusters.Add(placementResult.placedClusterSleeve);
                                    
                                    // Track ClashZoneIds for database save
                                    var clusterClashZoneIds = cluster
                                        .Select(s => (s.ClashZone as ClashZone)?.Id)
                                        .Where(id => id.HasValue)
                                        .Select(id => id!.Value)
                                        .ToList();
                                    
                                    _clusterToClashZoneIds[placementResult.capturedClusterSleeveId.Value] = clusterClashZoneIds;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Error($"[RefactoredClusterService] Error placing cluster: {ex.Message}");
                                SafeFileLogger.SafeAppendText("cluster_errors.log",
                                    $"[{DateTime.Now:HH:mm:ss.fff}] [RefactoredClusterService] ❌ Error placing cluster: {ex.Message}\nStackTrace: {ex.StackTrace}\n");
                            }
                            // Continue with next cluster
                        }
                    }
                }

                // ✅ STEP 15: Cleanup individual sleeves within placed clusters (Phase 7: Cleanup Service)
                if (placedClusters.Count > 0)
                {
                    int deletedInCleanup = _cleanupService.CleanupSleevesWithinClusters(doc, placedClusters);
                    deletedCount += deletedInCleanup;
                }

                // ✅ STEP 16: Return placed cluster sleeves if requested
                if (placedClusterSleevesOut != null)
                {
                    placedClusterSleevesOut.AddRange(placedClusters);
                }

                // ✅ STEP 17: Save cluster data to database (if comboId and filterId are available)
                if (comboId.HasValue && filterId.HasValue && _clusterToClashZoneIds.Count > 0)
                {
                    SaveClusterDataToDatabase(doc, comboId.Value, filterId.Value, targetCategory, _clusterToClashZoneIds);
                }

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[RefactoredClusterService] ✅ Completed: Placed {placedCount} clusters, Deleted {deletedCount} individual sleeves");
                }

                return (placedCount, deletedCount);
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[RefactoredClusterService] ❌ Critical error: {ex.Message}");
                    SafeFileLogger.SafeAppendText("cluster_errors.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [RefactoredClusterService] ❌ CRITICAL ERROR: {ex.Message}\nStackTrace: {ex.StackTrace}\n");
                }
                throw;
            }
        }

        /// <summary>
        /// Handle Path 1 Replay: Load pre-calculated clusters from database.
        /// </summary>
        private (bool hasData, int placedCount, int deletedCount) HandlePath1Replay(
            Document doc,
            int comboId,
            int filterId,
            string targetCategory,
            UIDocument? uiDoc,
            List<FamilyInstance>? placedClusterSleevesOut,
            string? xmlFilePath)
        {
            try
            {
                using (var dbContext = new SleeveDbContext(doc))
                {
                    var clusterRepository = new Data.Repositories.ClusterSleeveRepository(dbContext);
                    var existingClusters = clusterRepository.LoadClusterSleevesForCombo(comboId, targetCategory);
                    
                    if (existingClusters != null && existingClusters.Count > 0)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[RefactoredClusterService] PATH 1: Found {existingClusters.Count} pre-calculated clusters in database");
                        
                        // Use placement service to place from database
                        // TODO: Implement PlaceClustersFromDatabase method in PlacementService
                        // For now, return (false, 0, 0) to fall through to normal calculation
                        return (false, 0, 0);
                    }
                    else
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[RefactoredClusterService] PATH 1: No cluster data found, skipping clustering");
                        return (true, 0, 0); // Skip clustering
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[RefactoredClusterService] PATH 1: Error loading cluster data: {ex.Message}, falling back to calculation");
                return (false, 0, 0); // Fall through to normal calculation
            }
        }

        /// <summary>
        /// Get tolerance distance from settings (ClusterConfigurationManager).
        /// </summary>
        private double GetToleranceFromSettings(string? targetCategory)
        {
            var settingsService = new SettingsService();
            var currentProfile = ApplicationProfileService.Instance.GetCurrentProfile();
            if (currentProfile != null)
            {
                var settings = settingsService.GetSettings(currentProfile);
                ClusterConfigurationManager.Instance.LoadFromSettings(settings, "User Settings");
            }
            
            double toleranceMm = ClusterConfigurationManager.Instance.JoinOpeningsDistance;
            
            // Clamp pipes to 100mm
            if (!string.IsNullOrEmpty(targetCategory) && targetCategory.IndexOf("Pipe", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                toleranceMm = Math.Min(toleranceMm, 100);
            }
            
            return UnitUtils.ConvertToInternalUnits(toleranceMm, UnitTypeId.Millimeters);
        }

        /// <summary>
        /// Prepare sleeve data from clash zones (create dynamic objects with ClashZone references).
        /// </summary>
        private List<dynamic> PrepareSleeveData(List<ClashZone> filteredClashZones, List<ClashZone> allClashZones)
        {
            var rawSleeves = new List<dynamic>();
            
            foreach (var cz in filteredClashZones)
            {
                try
                {
                    // Get host type and orientation from ClashZone
                    string hostType = GetHostTypeFromClashZone(cz);
                    string orientation = GetEffectiveOrientationForClustering(cz);
                    
                    // Create bounding box from ClashZone data
                    BoundingBoxXYZ? bbox = GetBoundingBoxFromClashZone(cz);
                    
                    if (bbox == null) continue; // Skip sleeves without valid bounding boxes
                    
                    dynamic sleeveData = new
                    {
                        SleeveInstanceId = cz.SleeveInstanceId,
                        Category = cz.MepElementCategory,
                        HostType = hostType,
                        Orientation = orientation,
                        BoundingBox = bbox,
                        ClashZone = cz
                    };
                    
                    rawSleeves.Add(sleeveData);
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[RefactoredClusterService] Error preparing sleeve data for ClashZone {cz?.Id}: {ex.Message}");
                }
            }
            
            return rawSleeves;
        }

        /// <summary>
        /// Place a cluster sleeve for a group using all relevant services.
        /// </summary>
        private (bool success, int placedCount, int deletedCount, FamilyInstance? placedClusterSleeve, int? capturedClusterSleeveId) PlaceClusterForGroup(
            Document doc,
            List<dynamic> cluster,
            SleeveGroupKey groupKey,
            string targetCategory,
            string? xmlFilePath)
        {
            try
            {
                // ✅ Step 1: Determine rotation angle (Phase 6: Rotation Service)
                double rotationAngle = _rotationService.DetermineRotationAngle(cluster, xmlFilePath);
                
                // ✅ Step 2: Calculate bounding box (Phase 6: Rotation Service + Phase 3: BoundingBox)
                var actualSleeves = cluster
                    .Select(s => doc.GetElement(new ElementId(s.SleeveInstanceId)) as FamilyInstance)
                    .Where(fi => fi != null)
                    .ToList();
                
                var bboxResult = _rotationService.CalculateRotatedBoundingBox(cluster, actualSleeves, rotationAngle, xmlFilePath);
                
                if (bboxResult.width <= 0 || bboxResult.height <= 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[RefactoredClusterService] Invalid bounding box for cluster, skipping");
                    return (false, 0, 0, null, null);
                }
                
                // ✅ Step 3: Place cluster sleeve (Phase 5: Placement Service)
                bool placementSuccess = _placementService.PlaceClusterSleeve(
                    doc,
                    cluster,
                    groupKey,
                    targetCategory,
                    bboxResult.mid, // placement point
                    bboxResult.width,
                    bboxResult.height,
                    bboxResult.depth,
                    rotationAngle,
                    xmlFilePath,
                    out FamilyInstance? placedClusterSleeve,
                    out int? capturedClusterSleeveId);
                
                if (!placementSuccess || placedClusterSleeve == null)
                {
                    return (false, 0, 0, null, null);
                }
                
                // ✅ Step 4: Store rotation data (Phase 6: Rotation Service)
                if (capturedClusterSleeveId.HasValue)
                {
                    _rotationService.StoreRotationData(
                        capturedClusterSleeveId.Value,
                        rotationAngle * 180.0 / Math.PI,
                        Math.Abs(rotationAngle) > 1e-6,
                        new XYZ(bboxResult.rotatedMinX ?? 0, bboxResult.rotatedMinY ?? 0, bboxResult.rotatedMinZ ?? 0),
                        new XYZ(bboxResult.rotatedMaxX ?? 0, bboxResult.rotatedMaxY ?? 0, bboxResult.rotatedMaxZ ?? 0),
                        bboxResult.width,
                        bboxResult.height,
                        bboxResult.depth);
                }
                
                return (true, 1, 0, placedClusterSleeve, capturedClusterSleeveId);
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[RefactoredClusterService] Error in PlaceClusterForGroup: {ex.Message}");
                return (false, 0, 0, null, null);
            }
        }

        /// <summary>
        /// Save cluster data to database.
        /// </summary>
        private void SaveClusterDataToDatabase(
            Document doc,
            int comboId,
            int filterId,
            string targetCategory,
            Dictionary<int, List<Guid>> clusterToClashZoneIds)
        {
            try
            {
                using (var dbContext = new SleeveDbContext(doc))
                {
                    var clusterRepository = new Data.Repositories.ClusterSleeveRepository(dbContext);
                    
                    foreach (var kvp in clusterToClashZoneIds)
                    {
                        int clusterInstanceId = kvp.Key;
                        List<Guid> clashZoneIds = kvp.Value;
                        
                        // Save cluster data to database
                        // TODO: Implement SaveClusterSleeve method in ClusterSleeveRepository if needed
                        // For now, this is a placeholder
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[RefactoredClusterService] Error saving cluster data to database: {ex.Message}");
            }
        }

        #region Helper Methods (Extracted from UniversalClusterService)

        /// <summary>
        /// Get host type from ClashZone.
        /// </summary>
        private string GetHostTypeFromClashZone(ClashZone cz)
        {
            try
            {
                // Get host type from structural element type
                if (!string.IsNullOrEmpty(cz.StructuralElementType))
                {
                    if (cz.StructuralElementType.Contains("Wall", StringComparison.OrdinalIgnoreCase))
                        return "Wall";
                    else if (cz.StructuralElementType.Contains("Floor", StringComparison.OrdinalIgnoreCase) || 
                             cz.StructuralElementType.Contains("Slab", StringComparison.OrdinalIgnoreCase))
                        return "Floor";
                    else if (cz.StructuralElementType.Contains("Structural Framing", StringComparison.OrdinalIgnoreCase))
                        return "Structural Framing";
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[RefactoredClusterService] Error getting host type from clash zone: {ex.Message}");
            }
            
            return "Floor"; // Default fallback
        }

        /// <summary>
        /// Get effective orientation for clustering from ClashZone.
        /// </summary>
        private string GetEffectiveOrientationForClustering(ClashZone cz)
        {
            try
            {
                string hostType = GetHostTypeFromClashZone(cz);
                
                // For walls and framing, use HostOrientation (X/Y)
                if (hostType == "Wall" || hostType == "Structural Framing")
                {
                    return cz.HostOrientation ?? "Unknown";
                }
                // For floors, return "Floor" to group ALL floor sleeves together for clustering
                else if (hostType == "Floor")
                {
                    return "Floor"; // Unified "Floor" instead of X/Y split
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[RefactoredClusterService] Error getting effective orientation: {ex.Message}");
            }
            
            return "Vertical"; // Default fallback
        }

        /// <summary>
        /// Get bounding box from ClashZone.
        /// </summary>
        private BoundingBoxXYZ? GetBoundingBoxFromClashZone(ClashZone cz)
        {
            try
            {
                if (cz.SleeveBoundingBoxMinX == 0 && cz.SleeveBoundingBoxMaxX == 0 &&
                    cz.SleeveBoundingBoxMinY == 0 && cz.SleeveBoundingBoxMaxY == 0 &&
                    cz.SleeveBoundingBoxMinZ == 0 && cz.SleeveBoundingBoxMaxZ == 0)
                {
                    return null; // Invalid bounding box
                }
                
                return new BoundingBoxXYZ
                {
                    Min = new XYZ(cz.SleeveBoundingBoxMinX, cz.SleeveBoundingBoxMinY, cz.SleeveBoundingBoxMinZ),
                    Max = new XYZ(cz.SleeveBoundingBoxMaxX, cz.SleeveBoundingBoxMaxY, cz.SleeveBoundingBoxMaxZ),
                    Enabled = true
                };
            }
            catch
            {
                return null;
            }
        }

        #endregion
    }
}

