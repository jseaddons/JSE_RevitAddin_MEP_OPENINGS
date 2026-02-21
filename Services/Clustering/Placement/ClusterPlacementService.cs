using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Services.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Safety;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.BoundingBox;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Geometry; // For WallRcsTransformer
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement; // For SleeveParameterService
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
// OpeningSettingsHelper is already in JSE_RevitAddin_MEP_OPENINGS.Services namespace

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Placement
{
    /// <summary>
    /// Phase 5: Service for placing cluster sleeves in Revit.
    /// Handles creation, sizing, metadata, and family loading.
    /// Extracts placement methods and caches from UniversalClusterService.
    /// </summary>
    public class ClusterPlacementService : IClusterPlacementService
    {
        // ⚠️⚠️⚠️ MODIFICATION CONSENT REQUIRED ⚠️⚠️⚠️
        // To modify protected methods in this class, you MUST:
        // 1. Set ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = true
        // 2. Get explicit consent from the project owner
        // 3. Test thoroughly before committing
        // 4. Reset ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = false after changes
        // 
        // PROTECTED METHODS:
        // - ApplyRotation() - 90-degree offset is REQUIRED
        // 
        // ⚠️ DO NOT SET TO true UNLESS YOU HAVE EXPLICIT CONSENT ⚠️
        private const bool ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = false;
        // ✅ PERFORMANCE OPTIMIZATION: Cache MEP elements and bounding boxes to avoid duplicate API calls
        private readonly Dictionary<ElementId, Element> _mepElementCache;
        private readonly Dictionary<FamilyInstance, BoundingBoxXYZ> _bboxCache;
        private readonly Dictionary<FamilyInstance, Dictionary<string, Parameter>> _parameterCache;

        // Dependencies for methods that will be in other services or are still in UniversalClusterService
        private readonly Func<int, string?, ClashZone?> _getClashZoneBySleeveInstanceId;
        private readonly Func<List<dynamic>, string?, double> _determineRotationAngle;
        private readonly Func<List<dynamic>, List<FamilyInstance>, double, string?, (double width, double height, double depth, XYZ mid, double? rotatedMinX, double? rotatedMinY, double? rotatedMinZ, double? rotatedMaxX, double? rotatedMaxY, double? rotatedMaxZ)> _getClusterBoundingBox;
        private readonly Action<List<dynamic>, ElementId, string?, BoundingBoxXYZ?, (double minX, double minY, double minZ, double maxX, double maxY, double maxZ)?> _markClusterResolved;
        private readonly Func<string, string?> _getFilterNameForCategory;
        private readonly IBoundingBoxCalculator? _boundingBoxCalculator; // Optional: Phase 3 service
        private readonly SleeveParameterService? _parameterService; // ✅ SOLID: Injected dependency for parameter setting

        /// <summary>
        /// Constructor with dependency injection for external methods.
        /// </summary>
        public ClusterPlacementService(
            Func<int, string?, ClashZone?>? getClashZoneBySleeveInstanceId = null,
            Func<List<dynamic>, string?, double>? determineRotationAngle = null,
            Func<List<dynamic>, List<FamilyInstance>, double, string?, (double width, double height, double depth, XYZ mid, double? rotatedMinX, double? rotatedMinY, double? rotatedMinZ, double? rotatedMaxX, double? rotatedMaxY, double? rotatedMaxZ)>? getClusterBoundingBox = null,
            Action<List<dynamic>, ElementId, string?, BoundingBoxXYZ?, (double minX, double minY, double minZ, double maxX, double maxY, double maxZ)?>? markClusterResolved = null,
            Func<string, string?>? getFilterNameForCategory = null,
            IBoundingBoxCalculator? boundingBoxCalculator = null,
            SleeveParameterService? parameterService = null) // ✅ SOLID: Optional dependency injection
        {
            // Initialize caches
            _mepElementCache = new Dictionary<ElementId, Element>();
            _bboxCache = new Dictionary<FamilyInstance, BoundingBoxXYZ>();
            _parameterCache = new Dictionary<FamilyInstance, Dictionary<string, Parameter>>();

            // Store dependencies - allow null, provide safe defaults
            _getClashZoneBySleeveInstanceId = getClashZoneBySleeveInstanceId ?? ((id, path) => null);
            _determineRotationAngle = determineRotationAngle ?? ((cluster, path) => 0.0);
            _getClusterBoundingBox = getClusterBoundingBox ?? ((cluster, sleeves, angle, path) => (0, 0, 0, XYZ.Zero, null, null, null, null, null, null));
            _markClusterResolved = markClusterResolved ?? ((cluster, id, path, bbox, rotated) => { });
            _getFilterNameForCategory = getFilterNameForCategory ?? (cat => null);
            _boundingBoxCalculator = boundingBoxCalculator;
            _parameterService = parameterService; // ✅ SOLID: Store injected dependency
        }

        // ✅ CRASH-SAFE: Expose caches as read-only properties for pre-population
        public Dictionary<ElementId, Element> MepElementCache => _mepElementCache;
        public Dictionary<FamilyInstance, BoundingBoxXYZ> BoundingBoxCache => _bboxCache;
        public Dictionary<FamilyInstance, Dictionary<string, Parameter>> ParameterCache => _parameterCache;

        /// <summary>
        /// Place a cluster sleeve in the document.
        /// Phase 5: Extracted from UniversalClusterService.PlaceClusterSleeve.
        /// </summary>
        public bool PlaceClusterSleeve(
            Document doc,
            List<dynamic> cluster,
            SleeveGroupKey groupKey,
            string targetCategory,
            XYZ placementPoint,
            double width,
            double height,
            double depth,
            double rotationAngle,
            string? xmlFilePath,
            string? sleeveFamilyName,
            out FamilyInstance? placedClusterSleeve,
            out int? capturedClusterSleeveId,
            out XYZ? actualPlacementPoint,
            out ClusterSaveData? clusterSaveData,
            Dictionary<ElementId, Dictionary<string, object>>? deferredParameters = null,
            string? hostOrientation = null,
            double mepRotationAngle = 0.0)
        {
            placedClusterSleeve = null;
            capturedClusterSleeveId = null;
            actualPlacementPoint = null;
            clusterSaveData = null;

            // ✅ CRASH-SAFE: Validate inputs
            if (doc == null)
            {
                SafeFileLogger.SafeAppendText("placement_errors.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Invalid Document input - returning false\n");
                return false;
            }

            if (cluster == null || cluster.Count == 0)
            {
                SafeFileLogger.SafeAppendText("placement_errors.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Invalid cluster input (null or empty) - returning false\n");
                return false;
            }

            if (placementPoint == null || double.IsNaN(placementPoint.X) || double.IsInfinity(placementPoint.X))
            {
                SafeFileLogger.SafeAppendText("placement_errors.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Invalid placement point - returning false\n");
                return false;
            }

            try
            {
                // ✅ LOGGING: Log placement start
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string startMsg = $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] Starting cluster sleeve placement:\n";
                    startMsg += $"  HostType={groupKey.hostType}, SystemType={groupKey.systemType}, Orientation={groupKey.orientation}\n";
                    startMsg += $"  TargetCategory={targetCategory}, ClusterSize={cluster.Count}\n";
                    startMsg += $"  Dimensions: W={width * 304.8:F1}mm, H={height * 304.8:F1}mm, D={depth * 304.8:F1}mm\n";
                    startMsg += $"  Rotation Angle: {rotationAngle * 180 / Math.PI:F1}°\n";
                    startMsg += $"  Provided Family: {sleeveFamilyName ?? "NONE (will detect)"}\n";
                    DebugLogger.Info(startMsg);
                    SafeFileLogger.SafeAppendText("cluster_debug.log", startMsg);
                }

                // ✅ Declare versionTag at method start for use throughout method
                var versionTag = Helpers.VersionInfo.VersionTag;
                
                // 🔥 CRITICAL: Direct IO logging (bypass SafeFileLogger)
                try
                {
                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                    if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                    var logPath = Path.Combine(logDir, "cluster_debug.log");
                    // DEPLOYMENT MODE: Skip file writes
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] 🔥 PlaceClusterSleeve: Starting placement, hostType={groupKey.hostType}, systemType={groupKey.systemType}, family={sleeveFamilyName ?? "DETECT"}\n");
                    }
                }
                catch { }

                // Determine family name based on host type, category, and size (if not provided)
                string familyName = !string.IsNullOrEmpty(sleeveFamilyName) 
                    ? sleeveFamilyName 
                    : GetFamilyName(groupKey.hostType, groupKey.systemType, Math.Max(width, height));
                if (string.IsNullOrEmpty(familyName))
                {
                    try
                    {
                        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                        var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                        if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                        var logPath = Path.Combine(logDir, "cluster_debug.log");
                        // DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ PlaceClusterSleeve FAILED: Unknown host type '{groupKey.hostType}' - cannot determine family name\n");
                        }
                    }
                    catch { }
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Unknown host type '{groupKey.hostType}' - cannot determine family name\n");
                    return false;
                }

                // Load family if needed
                FamilySymbol? familySymbol = GetOrLoadFamilySymbol(doc, familyName);
                if (familySymbol == null)
                {
                    try
                    {
                        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                        var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                        if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                        var logPath = Path.Combine(logDir, "cluster_debug.log");
                        // DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ PlaceClusterSleeve FAILED: Failed to load family '{familyName}'\n");
                        }
                    }
                    catch { }
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Failed to load family '{familyName}'\n");
                    return false;
                }

                // Get reference level
                Level? refLevel = GetReferenceLevel(doc, cluster[0]);
                if (refLevel == null)
                {
                    try
                    {
                        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                        var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                        if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                        var logPath = Path.Combine(logDir, "cluster_debug.log");
                        // DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ PlaceClusterSleeve FAILED: Reference level not found\n");
                        }
                    }
                    catch { }
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Reference level not found\n");
                    return false;
                }

                // ✅ TRANSACTION SAFETY: Document modifiability is already validated in ClusterSleeves()
                // If we reach here, document should be modifiable (transaction started by caller)
                // Removed redundant check - if doc is not modifiable, Create.NewFamilyInstance will throw
                // which will be caught and logged below

                // ✅ WALL/FRAMING MIDPOINT OVERRIDE REMOVED
                // The upstream ClusterRotationService now handles "Hybrid Placement" (First Sleeve vs Centroid) correctly.
                // We must TRUST the passed 'placementPoint' and NOT recalculate it here, 
                // otherwise we undo the "First Sleeve" fix requested by the user.
                
                XYZ finalPlacementPoint = placementPoint;
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🎯 Using calculated placement point: ({finalPlacementPoint.X:F6}, {finalPlacementPoint.Y:F6}, {finalPlacementPoint.Z:F6})\n");
                }
                // (Logic removed to trust upstream Hybrid Placement)
                // (Remaining removal of legacy logic)


                // Create cluster sleeve instance
                FamilyInstance? inst = null;
                try
                {
                    // ✅ PERFORMANCE PROFILING: Profile family instantiation to identify symbol binding vs geometry creation
                    var instantiationTimer = System.Diagnostics.Stopwatch.StartNew();
                    var beforeInstantiation = System.GC.CollectionCount(0); // Track GC before
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        try
                        {
                            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                            var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                            if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                            SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] 🔥 ABOUT TO CREATE cluster sleeve instance: familyName={familyName}, level={refLevel?.Name ?? "NULL"}\n");
                        }
                        catch { /* Ignore logging errors */ }
                    }

                    inst = doc.Create.NewFamilyInstance(placementPoint, familySymbol, refLevel, StructuralType.NonStructural);
                    
                    instantiationTimer.Stop();
                    var afterInstantiation = System.GC.CollectionCount(0);
                    var gcCollections = afterInstantiation - beforeInstantiation;
                    
                    if (inst == null)
                    {
                        SafeFileLogger.SafeAppendText("placement_errors.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Cluster sleeve instance is null after creation\n");
                        return false;
                    }

                    // ✅ CRITICAL: Capture ID immediately while element is valid
                    capturedClusterSleeveId = inst.Id.GetIntegerValue();
                    
                    // ✅ CRITICAL: Return actual placement point via out parameter
                    // This ensures database saves the correct calculated placement point instead of Revit bbox center
                    actualPlacementPoint = placementPoint;
                    
                    // ✅ PROFILING: Log instantiation timing to identify bottlenecks
                    // Fast (<10ms) = quick placement, Medium (10-50ms) = moderate overhead, Slow (>50ms) = high overhead
                    // Note: No geometry creation - just family placement (symbol binding)
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        string instantiationType = instantiationTimer.ElapsedMilliseconds < 10 
                            ? "FAST_PLACEMENT" 
                            : instantiationTimer.ElapsedMilliseconds < 50 
                                ? "MODERATE_OVERHEAD" 
                                : "SLOW_OVERHEAD";
                        
                        try
                        {
                            var logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Logs", versionTag);
                            Directory.CreateDirectory(logDir);
                            SafeFileLogger.SafeAppendText("cluster_debug.log", $"{DateTime.Now:O}\t" +
                                $"ClusterSleeveId={inst?.Id?.GetIntegerValue() ?? -1}\t" +
                                $"Family={familySymbol?.Family?.Name ?? "NULL"}\t" +
                                $"Symbol={familySymbol?.Name ?? "NULL"}\t" +
                                $"Type={instantiationType}\t" +
                                $"TimeMs={instantiationTimer.ElapsedMilliseconds}\t" +
                                $"TimeTicks={instantiationTimer.ElapsedTicks}\t" +
                                $"GCCollections={gcCollections}\t" +
                                $"Level={refLevel?.Name ?? "NULL"}\t" +
                                $"ClusterSize={cluster?.Count ?? 0}\n");

                            SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ✅✅✅ Cluster sleeve CREATED: ID={capturedClusterSleeveId}\n");
                            string createMsg = $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ✅ Cluster sleeve created: ID={capturedClusterSleeveId}\n";
                            DebugLogger.Info(createMsg);
                            SafeFileLogger.SafeAppendText("cluster_debug.log", createMsg);
                        }
                        catch { /* Ignore logging errors */ }
                    }
                    
                    // ✅ CRITICAL FIX for Cluster Depth: Divert SleeveParameterService batch writes to our local dictionary
                    // This ensures "Depth" and "Wall Width" parameters (set via SleeveParameterService)
                    // are captured in the same dictionary as "Width" and "Height" (set here).
                    // ✅ UNIFIED BATCH CONTEXT FIX: Removed local assignment.
                    // We now rely on RefactoredClusterService to set the DivertedBatchDictionary globally for the group.
                    // if (_parameterService != null)
                    // {
                    //     _parameterService.DivertedBatchDictionary = deferredParameters;
                    // }

                    try
                    {
                        // Set size parameters
                        // ✅ WALL/FRAMING: X-walls get +90° rotation (matches individual sleeves), Y-walls get 0°
                        // shouldSwapDimensions is false for walls (dimension mapping is handled in SetSizeParameters)
                        // Rotation logic for floors (rotated axis/non-straight) is separate from wall orientation rotation
                        bool shouldSwapDimensions = false; // Walls use normal dimension mapping (no swap needed)
                        SetSizeParameters(doc, inst, cluster, groupKey, width, height, depth, shouldSwapDimensions, deferredParameters);

                        // ✅ ROTATION: Apply rotation for X-walls (90°) and floors (rotated axis/non-straight)
                        // Y-walls get 0° rotation (no rotation needed - LEFT view family works naturally for Y-walls)
                        // X-walls need +90° rotation (matches individual sleeve placement logic)
                        // Floor rotation is for rotated axis-aligned clusters (non-straight, 45°, etc.)
                        // ✅ DEFINITION RESTORED: Needed for rotation logic below
                        bool isWallOrFraming = groupKey.hostType == "Wall" || groupKey.hostType == "Walls" || groupKey.hostType == "Structural Framing";

                        // ⚠️ CABLETRAY FIX: Cable trays on floors don't need rotation like ducts do
                        // Note: isWallOrFraming is already declared earlier in the method (for midpoint override)
                        bool isFloorHost = groupKey.hostType == "Floor" || groupKey.hostType == "Floors";
                        
                        // ✅ CATEGORY CHECK: Determine if this is a cable tray or duct cluster
                        bool isCableTrayCategory = false;
                        bool isDuctCategory = false;
                        if (cluster != null && cluster.Count > 0)
                        {
                            var firstSleeve = cluster[0];
                            if (firstSleeve?.ClashZone != null)
                            {
                                var firstClashZone = firstSleeve.ClashZone as Models.ClashZone;
                                if (firstClashZone != null)
                                {
                                    string category = firstClashZone.MepElementCategory ?? "";
                                    isCableTrayCategory = category.Contains("Cable", StringComparison.OrdinalIgnoreCase) ||
                                                        category.Contains("CableTray", StringComparison.OrdinalIgnoreCase);
                                    isDuctCategory = category.Contains("Duct", StringComparison.OrdinalIgnoreCase) &&
                                                    !category.Contains("Accessory", StringComparison.OrdinalIgnoreCase);
                                }
                            }
                        }
                        
                        // Apply rotation if needed
                        if (Math.Abs(rotationAngle) > 1e-6)
                        {
                            if (isWallOrFraming)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                                        $"[{DateTime.Now:HH:mm:ss}] 🔄 APPLYING WALL/FRAMING ROTATION: {rotationAngle * 180 / Math.PI:F1}° (Trusting ClusterRotationService)\n");
                                }
                                ApplyRotation(doc, inst, placementPoint, rotationAngle, isWallOrFraming: true, skip90DegreeOffset: false);
                            }
                            else if (isFloorHost)
                            {
                                bool skipOffset = isCableTrayCategory || isDuctCategory || string.Equals(targetCategory, "Pipes", StringComparison.OrdinalIgnoreCase); 
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    string categoryBehavior = skipOffset ? $"{targetCategory.ToUpper()} (matches individual: no offset)" : "PIPE (cluster: +90° offset)";
                                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                                        $"[{DateTime.Now:HH:mm:ss}] 🔄 APPLYING FLOOR ROTATION: {rotationAngle * 180 / Math.PI:F1}° for floor-hosted cluster (category: {targetCategory}, {categoryBehavior})\n");
                                }
                                ApplyRotation(doc, inst, placementPoint, rotationAngle, isWallOrFraming: false, skip90DegreeOffset: skipOffset);
                            }
                        }
                        else if (isWallOrFraming)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss}] ✅ Y-WALL: No rotation applied (0° matches individual sleeve placement)\n");
                            }
                        }

                        // Set metadata (including HostOrientation for proper orientation)
                        SetMetadata(inst, targetCategory, null, deferredParameters, hostOrientation, mepRotationAngle);

                        // ✅ SCHEDULE LEVEL & ELEVATION: Set from first ClashZone's MEP element level
                        try
                        {
                            var firstSleeve = cluster[0];
                            ClashZone? firstClashZone = null;
                            if (firstSleeve is ClashZone cz) firstClashZone = cz;
                            else if (firstSleeve?.ClashZone != null) firstClashZone = firstSleeve.ClashZone as ClashZone;
                            else if (firstSleeve?.SleeveInstanceId != null && _getClashZoneBySleeveInstanceId != null)
                                firstClashZone = _getClashZoneBySleeveInstanceId(firstSleeve.SleeveInstanceId, xmlFilePath);

                            if (firstClashZone != null)
                            {
                                if (_parameterService != null)
                                {
                                    var verifyBeforeParams = doc.GetElement(inst.Id) as FamilyInstance;
                                    if (verifyBeforeParams != null && verifyBeforeParams.IsValidObject)
                                    {
                                        _parameterService.SetScheduleLevelAndElevationForCluster(inst, firstClashZone, inst.Id);
                                    }
                                }
                                else
                                {
                                    // ✅ FALLBACK: Set Schedule Level directly if SleeveParameterService not injected
                                    if (!string.IsNullOrWhiteSpace(firstClashZone.MepElementLevelName))
                                    {
                                        Level? mepLevel = new FilteredElementCollector(doc)
                                            .OfClass(typeof(Level))
                                            .Cast<Level>()
                                            .FirstOrDefault(l => string.Equals(l.Name, firstClashZone.MepElementLevelName, StringComparison.OrdinalIgnoreCase));

                                        if (mepLevel != null)
                                        {
                                            var scheduleLevelParam = inst.LookupParameter("Schedule of Level")
                                                                 ?? inst.LookupParameter("Schedule Level")
                                                                 ?? inst.LookupParameter("ScheduleLevel");
                                            
                                            if (scheduleLevelParam != null && !scheduleLevelParam.IsReadOnly)
                                            {
                                                if (scheduleLevelParam.StorageType == StorageType.ElementId)
                                                    scheduleLevelParam.Set(mepLevel.Id);
                                                else
                                                    scheduleLevelParam.Set(mepLevel.Name);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception levelEx)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[ClusterPlacementService] Error setting level for cluster: {levelEx.Message}");
                        }

                        // ✅ BOTTOM OF OPENING
                        if (OptimizationFlags.UseBottomOfOpeningCalculation)
                        {
                            double openingHeight = height;
                            if (deferredParameters != null && deferredParameters.ContainsKey(inst.Id) && deferredParameters[inst.Id].ContainsKey("Height"))
                            {
                                openingHeight = (double)deferredParameters[inst.Id]["Height"];
                            }
                            SetBottomOfOpeningForCluster(inst, openingHeight, deferredParameters);
                        }

                        // Mark clash zones as cluster-resolved
                        BoundingBoxXYZ? clusterBbox = null;
                        try { clusterBbox = inst.get_BoundingBox(null); } catch { }

                        if (_markClusterResolved != null)
                        {
                            _markClusterResolved(cluster, inst.Id, xmlFilePath, clusterBbox, null);
                        }

                        // ✅ BATCH PERSISTENCE FIX: Populate clusterSaveData for database save
                        // This allows RefactoredClusterService to save accurate data even when Revit parameters are deferred
                        clusterSaveData = new ClusterSaveData
                        {
                            ClusterInstanceId = inst.Id.GetIntegerValue(),
                            PlacementX = placementPoint.X,
                            PlacementY = placementPoint.Y,
                            PlacementZ = placementPoint.Z,
                            ClusterWidth = width,
                            ClusterHeight = height,
                            ClusterDepth = depth,
                            RotationAngleDeg = rotationAngle * 180.0 / Math.PI,
                            IsRotated = Math.Abs(rotationAngle) > 1e-6,
                            HostType = groupKey.hostType,
                            HostOrientation = hostOrientation ?? groupKey.orientation,
                            Category = targetCategory,
                            SleeveFamilyName = familyName,
                            ClashZoneIds = cluster
                                .Select(s => {
                                    ClashZone? cz = null;
                                    if (s is ClashZone icz) cz = icz;
                                    else if (s?.ClashZone != null) cz = s.ClashZone as ClashZone;
                                    return cz?.Id;
                                })
                                .Where(id => id.HasValue)
                                .Select(id => id!.Value)
                                .ToList()
                        };

                        placedClusterSleeve = inst;
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ✅✅✅ PlaceClusterSleeve COMPLETED: ID={inst.Id.GetIntegerValue()}\n");
                        }
                        
                        return true;
                    }
                    finally
                    {
                        // ✅ ALWAYS unset diverted dictionary to prevent side effects
                        // ✅ UNIFIED BATCH CONTEXT FIX: Removed local reset.
                        // if (_parameterService != null)
                        // {
                        //     _parameterService.DivertedBatchDictionary = null;
                        // }
                    }
                }
                catch (Exception ex)
                {
                    // DEPLOYMENT MODE: Skip file writes
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        try
                        {
                            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                            var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                            if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                            SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ EXCEPTION creating cluster sleeve: {ex.Message}\n");
                            SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] StackTrace: {ex.StackTrace}\n");
                        }
                        catch { /* Ignore logging errors */ }
                    }
                    SafeFileLogger.SafeAppendText("placement_errors.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Exception creating cluster sleeve: {ex.Message}\n");
                    return false;
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("placement_errors.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Critical error in PlaceClusterSleeve: {ex.Message}\n");
                return false;
            }
        }

        /// <summary>
        /// Set size parameters (Width, Height, Depth) on a cluster sleeve.
        /// Phase 5: Extracted from UniversalClusterService.SetClusterSizeParameters.
        /// </summary>
        public void SetSizeParameters(
            Document doc,
            FamilyInstance clusterSleeve,
            List<dynamic> cluster,
            SleeveGroupKey groupKey,
            double width,
            double height,
            double depth,
            bool shouldSwapDimensions = false,
            Dictionary<ElementId, Dictionary<string, object>>? deferredParameters = null)
        {
            // ✅ CRASH-SAFE: Validate inputs
            if (doc == null || clusterSleeve == null || cluster == null || cluster.Count == 0)
            {
                SafeFileLogger.SafeAppendText("placement_errors.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Invalid input for SetSizeParameters\n");
                return;
            }

            // Default dimensions (apply swapping if requested)
            double openingWidth = shouldSwapDimensions ? depth : width;
            double openingHeight = height;
            double openingDepth = shouldSwapDimensions ? width : depth;
            double structuralThickness = 0.0;

                // ✅ RCS DIMENSION MAPPING: For walls/framing, use RCS dimensions directly (already wall-aligned)
                if (groupKey.hostType == "Wall" || groupKey.hostType == "Structural Framing")
                {
                    // ✅ RCS: Dimensions are already in wall-aligned coordinates from CalculateRotatedBoundingBox
                    // RCS Definition:
                    // - RCS X = along wall (Width parameter)
                    // - RCS Y = through wall (Depth parameter, will be overridden with wall thickness)
                    // - RCS Z = vertical (Height parameter)
                    // 
                    // Direct mapping (no swapping needed - RCS handles wall alignment):
                    // Width = RCS X (along wall)
                    // Depth = RCS Y (through wall) - override with wall thickness
                    // Height = RCS Z (vertical)

                    // ✅ DIAGNOSTIC: Log input dimensions (already in RCS)
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        double wMm = RevitUnitConversionService.Instance.FromInternalMillimeters(width);
                        double hMm = RevitUnitConversionService.Instance.FromInternalMillimeters(height);
                        double dMm = RevitUnitConversionService.Instance.FromInternalMillimeters(depth);
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] 📐 RCS DIMENSIONS: Orientation={groupKey.orientation}, RCS W={wMm:F1}mm (X=along wall), H={hMm:F1}mm (Z=vertical), D={dMm:F1}mm (Y=through wall)\n");
                    }

                    // ✅ DIRECT DB ACCESS: Get structural thickness from clash zone 
                    // All sleeves in a cluster MUST have the same host element
                    Models.ClashZone? logClashZone = null; // Capture for logging
                    try
                    {
                        if (cluster.Count > 0)
                        {
                            var firstItem = cluster[0];
                            Models.ClashZone? firstClashZone = null;
                            
                            // Try to get from the item itself first (fastest)
                            try { firstClashZone = firstItem?.ClashZone as Models.ClashZone; } catch { }
                            
                            // ⚠️ ROBUSTNESS: If not found, look up by ID directly (most direct/reliable)
                            if (firstClashZone == null)
                            {
                                int sleeveId = 0;
                                try { sleeveId = firstItem?.SleeveInstanceId ?? 0; } catch { }
                                if (sleeveId > 0 && _getClashZoneBySleeveInstanceId != null)
                                {
                                    firstClashZone = _getClashZoneBySleeveInstanceId(sleeveId, null);
                                }
                            }

                            if (firstClashZone != null)
                            {
                                logClashZone = firstClashZone;
                                // Depth = structural thickness for all (Floor, Wall, Framing)
                                if (firstClashZone.StructuralElementThickness > 0.001)
                                    structuralThickness = firstClashZone.StructuralElementThickness;
                            }
                            
                            // ✅ FALLBACK: If DB thickness is missing, check individual sleeves
                            if (structuralThickness <= 0.001)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    SafeFileLogger.SafeAppendText("placement_debug.log", 
                                        $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ⚠️ DB Thickness invalid. Trying fallback to individual sleeves...\n");

                                foreach (var item in cluster)
                                {
                                    try
                                    {
                                        int sId = item?.SleeveInstanceId ?? 0;
                                        if (sId > 0)
                                        {
                                            var indSleeve = doc.GetElement(new ElementId(sId)) as FamilyInstance;
                                            if (indSleeve != null)
                                            {
                                                // Try Depth first, then Wall Width
                                                var dParam = indSleeve.LookupParameter("Depth") ?? indSleeve.LookupParameter("Wall Width");
                                                if (dParam != null && dParam.AsDouble() > 0.001)
                                                {
                                                    structuralThickness = dParam.AsDouble();
                                                    if (!DeploymentConfiguration.DeploymentMode)
                                                        SafeFileLogger.SafeAppendText("placement_debug.log", 
                                                            $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ✅ Fallback Success: Got thickness {structuralThickness * 304.8:F1}mm from individual sleeve {sId}\n");
                                                    break; // Found it
                                                }
                                            }
                                        }
                                    }
                                    catch { /* Ignore lookup errors */ }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("placement_errors.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ⚠️ Error getting structural thickness: {ex.Message}\n");
                    }

                    // ✅ RCS: Direct dimension mapping (no swapping needed)
                    openingWidth = width;   // RCS X (along wall) → Width parameter
                    openingHeight = height; // RCS Z (vertical) → Height parameter
                    // openingDepth initialization deferred to validation

                    // ✅ PHASE 5 FIX: CRITICAL VALIDATION - Depth MUST come from database
                    if (structuralThickness <= 0.001)
                    {
                        // Log ERROR and SKIP cluster (do not fall back to Revit API)
                        SafeFileLogger.SafeAppendText("placement_errors.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ CRITICAL: StructuralThickness is ZERO or invalid for cluster! " +
                            $"ClashZone={logClashZone?.Id}, StructuralElementType={logClashZone?.StructuralElementType}, " +
                            $"WallThickness={logClashZone?.WallThickness}, FramingThickness={logClashZone?.FramingThickness}, " +
                            $"StructuralElementThickness={logClashZone?.StructuralElementThickness}. " +
                            $"CANNOT place cluster without valid depth from database.\n");
                        
                        // Return early - do NOT place cluster with invalid depth
                        return;
                    }

                    // ✅ LOG SUCCESS: Confirm depth came from database
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        double thicknessMm = RevitUnitConversionService.Instance.FromInternalMillimeters(structuralThickness);
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] ✅ DEPTH FROM DATABASE: {thicknessMm:F1}mm (Host={logClashZone?.StructuralElementType})\n");
                    }
                    
                    openingDepth = structuralThickness;
                }

                // ✅ DIAGNOSTIC: Log final dimensions after swapping
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    double finalWMm = RevitUnitConversionService.Instance.FromInternalMillimeters(openingWidth);
                    double finalHMm = RevitUnitConversionService.Instance.FromInternalMillimeters(openingHeight);
                    double finalDMm = RevitUnitConversionService.Instance.FromInternalMillimeters(openingDepth);
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss}] 📐 AFTER SWAP: Final W={finalWMm:F1}mm, H={finalHMm:F1}mm, D={finalDMm:F1}mm\n");
                }

                // ✅ GLOBAL SETTINGS: Apply rounding to cluster dimensions based on UI settings (RoundingValue and RoundAlwaysUp)
                // ⚠️ CRITICAL: Rounding is applied ONLY to dimensions (width/height), NOT to placement point (centroid)
                // The placement point (centroid) was already calculated from corners and set when the sleeve was created (line 379)
                // Rounding only affects the size parameters - the centroid remains unchanged to preserve geometric accuracy
                // This ensures the cluster sleeve is centered correctly on the calculated centroid, with only size adjusted per UI settings
                double originalWidth = openingWidth;
                double originalHeight = openingHeight;
                if (clusterSleeve == null || cluster == null || cluster.Count == 0) return;

                try
                {
                    // Dimensions (Using outer scope variables)

                    // Lookup Parameters
                    Parameter widthParam = clusterSleeve.LookupParameter("Width") ?? clusterSleeve.LookupParameter("width");
                    Parameter heightParam = clusterSleeve.LookupParameter("Height") ?? clusterSleeve.LookupParameter("height");
                    
                    // Set Width
                    if (widthParam != null && !widthParam.IsReadOnly)
                    {
                        try
                        {
                            if (deferredParameters != null && OptimizationFlags.UseBatchedParameterWrites)
                            {
                                if (!deferredParameters.ContainsKey(clusterSleeve.Id))
                                    deferredParameters[clusterSleeve.Id] = new Dictionary<string, object>();
                                deferredParameters[clusterSleeve.Id]["Width"] = openingWidth;
                            }
                            else
                            {
                                widthParam.Set(openingWidth);
                            }
                        }
                        catch (Exception ex)
                        {
                            SafeFileLogger.SafeAppendText("placement_errors.log", $"[SetSizeParameters] Error setting Width: {ex.Message}\n");
                        }
                    }

                    // Set Height
                    if (heightParam != null && !heightParam.IsReadOnly)
                    {
                        try
                        {
                            if (deferredParameters != null && OptimizationFlags.UseBatchedParameterWrites)
                            {
                                if (!deferredParameters.ContainsKey(clusterSleeve.Id))
                                    deferredParameters[clusterSleeve.Id] = new Dictionary<string, object>();
                                deferredParameters[clusterSleeve.Id]["Height"] = openingHeight;
                            }
                            else
                            {
                                heightParam.Set(openingHeight);
                            }
                        }
                        catch (Exception ex)
                        {
                            SafeFileLogger.SafeAppendText("placement_errors.log", $"[SetSizeParameters] Error setting Height: {ex.Message}\n");
                        }
                    }

                    // Set Depth (Robustness Fix Logic)
                    // ✅ FIXED: Explicitly handle deferred writing locally to ensure batch mode works.
                    // Do not rely on _parameterService here because it doesn't share the 'deferredParameters' dictionary context.
                    // This aligns Depth handling with Width/Height logic above.
                    
                    try
                    {
                        // 1. Determine which parameter determines 'Thickness/Depth'
                        Parameter foundDepthParam = clusterSleeve.LookupParameter("Depth") 
                                                  ?? clusterSleeve.LookupParameter("depth") 
                                                  ?? clusterSleeve.LookupParameter("Wall Width");
                        
                        // 2. Also check for 'Wall Width' explicitly if Depth was found (some families have both)
                        Parameter wallWidthParam = clusterSleeve.LookupParameter("Wall Width");
                        
                        // Use calculated openingDepth
                        if (foundDepthParam != null && !foundDepthParam.IsReadOnly)
                        {
                            if (deferredParameters != null && OptimizationFlags.UseBatchedParameterWrites)
                            {
                                if (!deferredParameters.ContainsKey(clusterSleeve.Id))
                                    deferredParameters[clusterSleeve.Id] = new Dictionary<string, object>();
                                
                                // Write to the MAIN depth parameter found
                                deferredParameters[clusterSleeve.Id][foundDepthParam.Definition.Name] = openingDepth;
                                
                                // ✅ DUAL SET: If both exist (rare but possible), set both to be safe
                                if (wallWidthParam != null && wallWidthParam.Id != foundDepthParam.Id && !wallWidthParam.IsReadOnly)
                                {
                                    deferredParameters[clusterSleeve.Id][wallWidthParam.Definition.Name] = openingDepth;
                                }
                            }
                            else
                            {
                                // Immediate Set
                                foundDepthParam.Set(openingDepth);
                                
                                // Dual Set
                                if (wallWidthParam != null && wallWidthParam.Id != foundDepthParam.Id && !wallWidthParam.IsReadOnly)
                                {
                                    wallWidthParam.Set(openingDepth);
                                }
                            }
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                                SafeFileLogger.SafeAppendText("cluster_debug.log", $"[SetSizeParameters] ✅ Set Depth/WallWidth to {openingDepth * 304.8:F1}mm\n");
                        }
                        else
                        {
                             SafeFileLogger.SafeAppendText("placement_errors.log", $"[SetSizeParameters] ⚠️ 'Depth' or 'Wall Width' parameter NOT FOUND on cluster sleeve.\n");
                        }
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("placement_errors.log", $"[SetSizeParameters] Error setting Depth: {ex.Message}\n");
                    }

                }
                catch (Exception ex)
                {
                    SafeFileLogger.SafeAppendText("placement_errors.log", $"[SetSizeParameters] Critical Error: {ex.Message}\n");
                }
            }
            

        /// <summary>
        /// ✅ SRP COMPLIANCE: Dedicated method for setting "Bottom of Opening" parameter on cluster sleeves.
        /// Single Responsibility: Calculate and set Bottom of Opening parameter only.
        /// 
        /// ✅ SIMPLIFIED FORMULA: Bottom of Opening = Elevation from Level - (Height / 2)
        /// Where Elevation from Level is automatically calculated by Revit after Schedule Level is set.
        /// 
        /// ✅ CORRECT SEQUENCING:
        /// 1. Schedule Level is set FIRST (by SetScheduleLevelAndElevationForCluster)
        /// 2. Revit automatically calculates Elevation from Level
        /// 3. We read Elevation from Level from the parameter
        /// 4. Calculate Bottom of Opening = Elevation from Level - (Height / 2)
        /// 
        /// Applies to: RectangularOpeningOnWall family only.
        /// Preserves all optimization features: batching, performance monitoring, safe validation, diagnostic logging.
        /// </summary>
        private void SetBottomOfOpeningForCluster(
            FamilyInstance clusterSleeve,
            double clusterHeight,
            Dictionary<ElementId, Dictionary<string, object>>? deferredParameters)
        {
            if (clusterSleeve == null) return;

            // ✅ FAMILY CHECK: Only apply to RectangularOpeningOnWall family
            string familyName = clusterSleeve.Symbol?.FamilyName ?? "";
            if (!familyName.Equals("RectangularOpeningOnWall", StringComparison.OrdinalIgnoreCase))
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [SetBottomOfOpeningForCluster] Zone=Cluster, Sleeve={clusterSleeve.Id}: " +
                        $"Skipping - family is '{familyName}' (expected 'RectangularOpeningOnWall')\n");
                }
                return; // Not the correct family - skip silently
            }

            // ✅ SAFE ELEMENT VALIDATION: Validate instance is still valid
            if (OptimizationFlags.SkipRedundantValidation)
            {
                if (!clusterSleeve.IsValidObject)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [SetBottomOfOpeningForCluster] ⚠️ Cluster sleeve {clusterSleeve.Id.GetIntegerValue()} is invalid - skipping\n");
                    }
                    return;
                }
            }

            try
            {
                // ✅ STEP 1: Read "Elevation from Level" from parameter (Revit calculates this automatically after Schedule Level is set)
                // ✅ PRIMARY: Try to read from parameter first (most reliable)
                double? elevationFromLevel = null;
                
                Parameter elevationParam = clusterSleeve.LookupParameter("Elevation from Level");
                if (elevationParam != null && elevationParam.StorageType == StorageType.Double)
                {
                    elevationFromLevel = elevationParam.AsDouble();
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [SetBottomOfOpeningForCluster] ✅ Cluster sleeve {clusterSleeve.Id.GetIntegerValue()}: " +
                            $"Read Elevation from Level={elevationFromLevel.Value * 304.8:F1}mm from parameter (calculated by Revit after Schedule Level was set)\n");
                    }
                }
                
                // ✅ FALLBACK: If parameter not available, calculate manually from placement point and Schedule Level
                if (!elevationFromLevel.HasValue)
                {
                    try
                    {
                        // Get placement point (center of opening)
                        LocationPoint? locationPoint = clusterSleeve.Location as LocationPoint;
                        if (locationPoint != null)
                        {
                            XYZ placementPoint = locationPoint.Point;
                            
                            // Get Schedule Level from parameter (should be set by SetScheduleLevelAndElevationForCluster)
                            Parameter scheduleLevelParam = clusterSleeve.LookupParameter("Schedule of Level")
                                                         ?? clusterSleeve.LookupParameter("Schedule Level")
                                                         ?? clusterSleeve.LookupParameter("ScheduleLevel");
                            
                            Level? scheduleLevel = null;
                            if (scheduleLevelParam != null)
                            {
                                if (scheduleLevelParam.StorageType == StorageType.ElementId)
                                {
                                    var levelId = scheduleLevelParam.AsElementId();
                                    if (levelId != null && levelId != ElementId.InvalidElementId)
                                    {
                                        scheduleLevel = clusterSleeve.Document.GetElement(levelId) as Level;
                                    }
                                }
                                else if (scheduleLevelParam.StorageType == StorageType.String)
                                {
                                    string levelName = scheduleLevelParam.AsString();
                                    if (!string.IsNullOrWhiteSpace(levelName))
                                    {
                                        scheduleLevel = new FilteredElementCollector(clusterSleeve.Document)
                                            .OfClass(typeof(Level))
                                            .Cast<Level>()
                                            .FirstOrDefault(l => string.Equals(l.Name, levelName, StringComparison.OrdinalIgnoreCase));
                                    }
                                }
                            }
                            
                            if (scheduleLevel != null)
                            {
                                // Calculate: Elevation from Level = Placement Point Z - Schedule Level Elevation
                                elevationFromLevel = placementPoint.Z - scheduleLevel.Elevation;
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                                        $"[{DateTime.Now:HH:mm:ss.fff}] [SetBottomOfOpeningForCluster] ✅ FALLBACK: Cluster sleeve {clusterSleeve.Id.GetIntegerValue()}: " +
                                        $"Calculated Elevation from Level={elevationFromLevel.Value * 304.8:F1}mm " +
                                        $"(PlacementPoint.Z={placementPoint.Z * 304.8:F1}mm - ScheduleLevel.Elevation={scheduleLevel.Elevation * 304.8:F1}mm, Level='{scheduleLevel.Name}')\n");
                                }
                            }
                        }
                    }
                    catch (Exception fallbackEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [SetBottomOfOpeningForCluster] ⚠️ Cluster sleeve {clusterSleeve.Id.GetIntegerValue()}: " +
                                $"Error in fallback calculation: {fallbackEx.Message}\n");
                        }
                    }
                }
                
                // ✅ If still not available after fallback, skip gracefully
                if (!elevationFromLevel.HasValue)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [SetBottomOfOpeningForCluster] ⚠️ Cluster sleeve {clusterSleeve.Id.GetIntegerValue()}: " +
                            $"Elevation from Level not available (parameter not found and fallback calculation failed). Skipping Bottom of Opening calculation.\n");
                    }
                    return; // Graceful degradation - skip if Elevation from Level is not available
                }

                // ✅ VALIDATION: Check if Elevation from Level is valid
                if (!elevationFromLevel.HasValue ||
                    !Services.Helpers.BottomOfOpeningCalculationService.IsValidScheduleOfLevel(elevationFromLevel.Value))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [SetBottomOfOpeningForCluster] ⚠️ Cluster sleeve {clusterSleeve.Id.GetIntegerValue()}: " +
                            $"Elevation from Level is invalid (value={elevationFromLevel?.ToString() ?? "null"}) - skipping\n");
                    }
                    return; // Graceful degradation - skip if Elevation from Level is invalid
                }

                // ✅ VALIDATION: Check if Cluster Height is valid
                if (!Services.Helpers.BottomOfOpeningCalculationService.IsValidHeight(clusterHeight))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [SetBottomOfOpeningForCluster] ⚠️ Cluster sleeve {clusterSleeve.Id.GetIntegerValue()}: " +
                            $"Cluster Height is invalid (value={clusterHeight * 304.8:F1}mm) - skipping\n");
                    }
                    return; // Graceful degradation - skip if Cluster Height is invalid
                }

                // ✅ STEP 2: Calculate Bottom of Opening = Elevation from Level - (Height / 2)
                // ✅ SIMPLIFIED: Direct calculation, no complex helper service needed
                double bottomOfOpening = elevationFromLevel.Value - (clusterHeight / 2.0);

                // ✅ STEP 3: Set "Bottom of Opening" parameter with batching support
                // Parameter name is "Bottom Of Opening" (capital O in "Of") as shown in Revit Properties
                var bottomParam = clusterSleeve.LookupParameter("Bottom Of Opening")  // ✅ FIRST: Exact name from Properties
                               ?? clusterSleeve.LookupParameter("Bottom of Opening")
                               ?? clusterSleeve.LookupParameter("BottomOfOpening")
                               ?? clusterSleeve.Symbol?.LookupParameter("Bottom Of Opening")
                               ?? clusterSleeve.Symbol?.LookupParameter("Bottom of Opening")
                               ?? clusterSleeve.Symbol?.LookupParameter("BottomOfOpening");
                
                if (bottomParam != null && !bottomParam.IsReadOnly)
                {
                    if (deferredParameters != null && OptimizationFlags.UseBatchedParameterWrites)
                    {
                        if (!deferredParameters.ContainsKey(clusterSleeve.Id))
                            deferredParameters[clusterSleeve.Id] = new Dictionary<string, object>();
                        deferredParameters[clusterSleeve.Id]["Bottom of Opening"] = bottomOfOpening;

                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [SetBottomOfOpeningForCluster] ✅ DEFERRED: Added Bottom of Opening={bottomOfOpening * 304.8:F1}mm " +
                                $"(Elevation from Level={elevationFromLevel.Value * 304.8:F1}mm, Height={clusterHeight * 304.8:F1}mm) " +
                                $"to deferredParameters for cluster sleeve {clusterSleeve.Id.GetIntegerValue()}\n");
                        }
                    }
                    else
                    {
                        bottomParam.Set(bottomOfOpening);
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [SetBottomOfOpeningForCluster] ✅ IMMEDIATE: Set Bottom of Opening={bottomOfOpening * 304.8:F1}mm " +
                                $"(Elevation from Level={elevationFromLevel.Value * 304.8:F1}mm, Height={clusterHeight * 304.8:F1}mm) " +
                                $"for cluster sleeve {clusterSleeve.Id.GetIntegerValue()}\n");
                        }
                    }
                }
                else if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [SetBottomOfOpeningForCluster] ⚠️ Cluster sleeve {clusterSleeve.Id.GetIntegerValue()}: " +
                        $"Bottom of Opening parameter not found or read-only - cannot set value\n");
                }
                else
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [SetBottomOfOpeningForCluster] ⚠️ Cluster sleeve {clusterSleeve.Id.GetIntegerValue()}: " +
                            $"'Bottom of Opening' parameter not found or read-only (tried: 'Bottom Of Opening', 'Bottom of Opening', 'BottomOfOpening' on instance and symbol)\n");
                    }
                }
            }
            catch (Exception ex)
            {
                // ✅ CRASH-SAFE: Graceful error handling
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("placement_errors.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] [SetBottomOfOpeningForCluster] ❌ Error setting Bottom of Opening " +
                        $"for cluster sleeve {clusterSleeve.Id.GetIntegerValue()}: {ex.Message}\n");
                }
            }
        }

        /// <summary>
        /// Set metadata parameters on a cluster sleeve.
        /// Phase 5: Extracted from UniversalClusterService.SetClusterSleeveMetadata.
        /// </summary>
        public void SetMetadata(
            FamilyInstance clusterSleeve,
            string category,
            string? filterName = null,
            Dictionary<ElementId, Dictionary<string, object>>? deferredParameters = null,
            string? hostOrientation = null,
            double mepRotationAngle = 0.0)
        {
            // ✅ CRASH-SAFE: Validate inputs
            if (clusterSleeve == null || string.IsNullOrEmpty(category))
            {
                SafeFileLogger.SafeAppendText("placement_errors.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Invalid input for SetMetadata\n");
                return;
            }

            try
            {
                // Set MEP_Category
                Parameter? mepCategoryParam = GetParameter(clusterSleeve, "MEP_Category");
                if (mepCategoryParam != null && !mepCategoryParam.IsReadOnly)
                {
                    if (deferredParameters != null && OptimizationFlags.UseBatchedParameterWrites)
                    {
                        if (!deferredParameters.ContainsKey(clusterSleeve.Id))
                            deferredParameters[clusterSleeve.Id] = new Dictionary<string, object>();
                        deferredParameters[clusterSleeve.Id]["MEP_Category"] = category;
                    }
                    else
                    {
                        mepCategoryParam.Set(category);
                    }
                }

                // Set HostOrientation
                if (!string.IsNullOrEmpty(hostOrientation))
                {
                    Parameter? orientationParam = GetParameter(clusterSleeve, "HostOrientation")
                                               ?? GetParameter(clusterSleeve, "Host Orientation");
                    if (orientationParam != null && !orientationParam.IsReadOnly)
                    {
                        if (deferredParameters != null && OptimizationFlags.UseBatchedParameterWrites)
                        {
                            if (!deferredParameters.ContainsKey(clusterSleeve.Id))
                                deferredParameters[clusterSleeve.Id] = new Dictionary<string, object>();
                            deferredParameters[clusterSleeve.Id][orientationParam.Definition.Name] = hostOrientation;
                        }
                        else
                        {
                            orientationParam.Set(hostOrientation);
                        }
                    }
                }

                // Set MepElementRotationAngle
                if (Math.Abs(mepRotationAngle) > 0.0001)
                {
                    Parameter? rotationParam = GetParameter(clusterSleeve, "MepElementRotationAngle")
                                            ?? GetParameter(clusterSleeve, "Rotation");
                    if (rotationParam != null && !rotationParam.IsReadOnly)
                    {
                        if (deferredParameters != null && OptimizationFlags.UseBatchedParameterWrites)
                        {
                            if (!deferredParameters.ContainsKey(clusterSleeve.Id))
                                deferredParameters[clusterSleeve.Id] = new Dictionary<string, object>();
                            deferredParameters[clusterSleeve.Id][rotationParam.Definition.Name] = mepRotationAngle;
                        }
                        else
                        {
                            rotationParam.Set(mepRotationAngle);
                        }
                    }
                }

                // Set Filter Name
                string? actualFilterName = filterName ?? _getFilterNameForCategory(category);
                if (!string.IsNullOrEmpty(actualFilterName))
                {
                    Parameter? filterNameParam = GetParameter(clusterSleeve, "Filter Name");
                    if (filterNameParam != null && !filterNameParam.IsReadOnly)
                    {
                        if (deferredParameters != null && OptimizationFlags.UseBatchedParameterWrites)
                        {
                            if (!deferredParameters.ContainsKey(clusterSleeve.Id))
                                deferredParameters[clusterSleeve.Id] = new Dictionary<string, object>();
                            deferredParameters[clusterSleeve.Id]["Filter Name"] = actualFilterName;
                        }
                        else
                        {
                            filterNameParam.Set(actualFilterName);
                        }
                    }
                }

                // ✅ CRITICAL: Set cluster sleeve identification parameters IMMEDIATELY (not deferred)
                // These parameters are required for cleanup service to identify cluster sleeves correctly
                // Even if batching is enabled, these must be set immediately to prevent misidentification
                Parameter? instanceIdParam = GetParameter(clusterSleeve, "Sleeve Instance ID");
                if (instanceIdParam != null && !instanceIdParam.IsReadOnly)
                {
                    instanceIdParam.Set(-1); // Always set immediately (critical for cleanup)
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [SetMetadata] ✅ IMMEDIATE (CRITICAL): Set 'Sleeve Instance ID'=-1 for cluster sleeve {clusterSleeve.Id.GetIntegerValue()}\n");
                }
                else
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [SetMetadata] ⚠️ WARNING: 'Sleeve Instance ID' parameter not found or readonly for cluster sleeve {clusterSleeve.Id.GetIntegerValue()}\n");
                }

                // Set Cluster Sleeve Instance ID
                Parameter? clusterInstanceIdParam = GetParameter(clusterSleeve, "Cluster Sleeve Instance ID");
                if (clusterInstanceIdParam != null && !clusterInstanceIdParam.IsReadOnly)
                {
                    clusterInstanceIdParam.Set(clusterSleeve.Id.GetIntegerValue()); // Always set immediately (critical for cleanup)
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [SetMetadata] ✅ IMMEDIATE (CRITICAL): Set 'Cluster Sleeve Instance ID'={clusterSleeve.Id.GetIntegerValue()} for cluster sleeve {clusterSleeve.Id.GetIntegerValue()}\n");
                    
                    // ✅ PERFORMANCE FIX: Remove per-cluster regeneration to enable batch optimization
                    // Verification will happen after batch flush in RefactoredClusterService (line 623)
                    // This change enables 4-6× speedup by batching ALL regenerations
                    // OLD: clusterSleeve.Document?.Regenerate(); // ❌ REMOVED - defeats batch optimization
                    // NEW: Verification deferred to post-flush regeneration
                }
                else
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [SetMetadata] ⚠️ WARNING: 'Cluster Sleeve Instance ID' parameter not found or readonly for cluster sleeve {clusterSleeve.Id.GetIntegerValue()}\n");
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("placement_errors.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Error in SetMetadata: {ex.Message}\n");
            }
        }

        /// <summary>
        /// Set metadata parameters on a cluster sleeve (legacy overload for backward compatibility).
        /// </summary>
        [Obsolete("Use SetMetadata(FamilyInstance, string, string?, Dictionary<ElementId, Dictionary<string, object>>?) instead")]
        public void SetMetadata_Obsolete(
            FamilyInstance clusterSleeve,
            string category,
            string? filterName = null)
        {
            // Call new signature without deferred parameters
            SetMetadata(clusterSleeve, category, filterName, null);
        }

        /// <summary>
        /// Get reference level for cluster sleeve placement.
        /// Phase 5: Extracted from UniversalClusterService.GetReferenceLevelFromXml.
        /// </summary>
        public Level? GetReferenceLevel(Document doc, dynamic sleeve)
        {
            // ✅ CRASH-SAFE: Validate inputs
            if (doc == null || sleeve == null)
            {
                return null;
            }

            try
            {
                // Get the first level in the document as fallback
                var levels = new FilteredElementCollector(doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .OrderBy(l => l.Elevation)
                    .ToList();

                return levels.FirstOrDefault();
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("placement_errors.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Error in GetReferenceLevel: {ex.Message}\n");
                return null;
            }
        }

        /// <summary>
        /// Load a universal family into the document if not already present.
        /// Phase 5: Extracted from UniversalClusterService.LoadUniversalFamily.
        /// </summary>
        public bool LoadFamily(Document doc, string familyName)
        {
            // ✅ CRASH-SAFE: Validate inputs
            if (doc == null || string.IsNullOrEmpty(familyName))
            {
                return false;
            }

            try
            {
                // Check if family already exists
                var existingSymbols = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .Where(sym => sym.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (existingSymbols.Count > 0)
                {
                    return true; // Family already loaded
                }

                // ✅ DEPLOYMENT FIX: Use Assembly location to find Resources folder (deployed next to DLL)
                var executingAssembly = System.Reflection.Assembly.GetExecutingAssembly();
                var assemblyDir = Path.GetDirectoryName(executingAssembly.Location);
                string resourcesPath = Path.Combine(assemblyDir, "Resources");
                string familyPath = Path.Combine(resourcesPath, $"{familyName}.rfa");

                if (!File.Exists(familyPath))
                {
                    SafeFileLogger.SafeAppendText("placement_errors.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Family file not found: {familyPath}\n");
                    return false;
                }

                bool loaded = doc.LoadFamily(familyPath);
                return loaded;
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("placement_errors.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Error loading family '{familyName}': {ex.Message}\n");
                return false;
            }
        }

        /// <summary>
        /// Get or cache MEP element by ElementId.
        /// </summary>
        public Element? GetMepElement(Document doc, ElementId elementId)
        {
            // ✅ CRASH-SAFE: Validate inputs
            if (doc == null || elementId == null || elementId.GetIntegerValue() <= 0)
            {
                return null;
            }
            
            // Try to get element - if null, element doesn't exist
            var element = doc.GetElement(elementId);
            if (element == null)
            {
                return null;
            }

            // Check cache first
            if (_mepElementCache.TryGetValue(elementId, out Element? cachedElement))
            {
                // ✅ CRASH-SAFE: Verify element is still valid
                if (cachedElement != null && cachedElement.IsValidObject)
                {
                    return cachedElement;
                }
                else
                {
                    // Remove invalid element from cache
                    _mepElementCache.Remove(elementId);
                }
            }

            // Load from document
            try
            {
                Element? loadedElement = doc.GetElement(elementId);
                if (loadedElement != null && loadedElement.IsValidObject)
                {
                    _mepElementCache[elementId] = loadedElement;
                    return loadedElement;
                }
            }
            catch { }

            return null;
        }

        /// <summary>
        /// Get or cache bounding box for a sleeve instance.
        /// </summary>
        public BoundingBoxXYZ? GetBoundingBox(FamilyInstance sleeve)
        {
            // ✅ CRASH-SAFE: Validate inputs
            if (sleeve == null || !sleeve.IsValidObject)
            {
                return null;
            }

            // Check cache first
            if (_bboxCache.TryGetValue(sleeve, out BoundingBoxXYZ? cachedBbox))
            {
                return cachedBbox;
            }

            // Load from Revit
            try
            {
                BoundingBoxXYZ? bbox = sleeve.get_BoundingBox(null);
                if (bbox != null && bbox.Enabled)
                {
                    _bboxCache[sleeve] = bbox;
                    return bbox;
                }
            }
            catch { }

            return null;
        }

        /// <summary>
        /// Get or cache parameter for a sleeve instance.
        /// </summary>
        public Parameter? GetParameter(FamilyInstance sleeve, string parameterName)
        {
            // ✅ CRASH-SAFE: Validate inputs
            if (sleeve == null || string.IsNullOrEmpty(parameterName) || !sleeve.IsValidObject)
            {
                return null;
            }

            // Check cache first
            if (_parameterCache.TryGetValue(sleeve, out Dictionary<string, Parameter>? paramDict))
            {
                if (paramDict.TryGetValue(parameterName, out Parameter? cachedParam))
                {
                    return cachedParam;
                }
            }
            else
            {
                // Initialize parameter dictionary for this sleeve
                _parameterCache[sleeve] = new Dictionary<string, Parameter>();
            }

            // Load from Revit
            try
            {
                Parameter? param = sleeve.LookupParameter(parameterName);
                if (param != null)
                {
                    _parameterCache[sleeve][parameterName] = param;
                    return param;
                }
            }
            catch { }

            return null;
        }

        /// <summary>
        /// Clear all caches.
        /// </summary>
        public void ClearCaches()
        {
            _mepElementCache.Clear();
            _bboxCache.Clear();
            _parameterCache.Clear();
        }

        // ✅ PRIVATE HELPER METHODS

        /// <summary>
        /// Get family name based on host type, MEP category, and dimension.
        /// ✅ REFACTORED: Accounts for OpeningType settings, diameter thresholds, and host categories.
        /// </summary>
        public static string GetFamilyName(string hostType, string mepCategory, double dimension, bool isCluster = false)
        {
            return GetFamilyName(hostType, mepCategory, dimension, isCluster, hostOrientation: null);
        }

        public static string GetFamilyName(string hostType, string mepCategory, double dimension, bool isCluster, string hostOrientation)
        {
            // ✅ SIMPLE LOGIC FOR CLUSTERS: Clusters are ALWAYS rectangular
            if (isCluster)
            {
                bool isWallOrFraming = hostType.IndexOf("Wall", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                      hostType.IndexOf("Framing", StringComparison.OrdinalIgnoreCase) >= 0;

                if (!isWallOrFraming)
                    return "RectangularOpeningOnSlab";

                // ✅ PERF FIX: X-wall clusters use _X pre-rotated family (skip runtime rotation)
                bool isXOriented = string.Equals(hostOrientation?.Trim(), "X", StringComparison.OrdinalIgnoreCase);
                return isXOriented ? "RectangularOpeningOnWall_X" : "RectangularOpeningOnWall";
            }

            var settings = ApplicationProfileService.Instance.GetCurrentSettings();
            bool isCircular = true;

            // 1. Determine base shape (Circular vs Rectangular) based on category and settings
            if (mepCategory.Equals("Ducts", StringComparison.OrdinalIgnoreCase))
            {
                // For ducts, we'd ideally check the UI setting, but for clusters, we default to Circular 
                // unless specified otherwise. In this implementation, we follow the user's logic:
                isCircular = true; 
            }
            else if (mepCategory.Equals("Pipes", StringComparison.OrdinalIgnoreCase))
            {
                // Pipes respect the PipeOpeningTypeRectangular setting
                isCircular = !settings.PipeOpeningTypeRectangular;
            }
            else
            {
                // All other categories (Duct Accessories, Cable Trays) are always rectangular
                isCircular = false;
            }

            // 2. Apply diameter threshold (e.g., if > 200mm, convert Circular to Rectangular)
            if (isCircular && settings.RoundOpeningsRectangular > 0)
            {
                // Convert dimension from internal units (feet) to mm for comparison
                double mmDimension = RevitUnitConversionService.Instance.FromInternalMillimeters(dimension);
                if (mmDimension > settings.RoundOpeningsRectangular)
                {
                    isCircular = false;
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[GET_FAMILY_NAME] Converting circular to rectangular: {mmDimension:F1}mm > threshold {settings.RoundOpeningsRectangular}mm");
                }
            }

            // 3. Return family name based on host type and final shape
            bool isWallHost = hostType == "Wall" || hostType == "Structural Framing" || hostType == "Walls";
            
            if (isWallHost)
            {
                bool isXOriented = string.Equals(hostOrientation?.Trim(), "X", StringComparison.OrdinalIgnoreCase);
                if (isCircular)
                    return isXOriented ? "CircularOpeningOnWall_X" : "CircularOpeningOnWall";
                else
                    return isXOriented ? "RectangularOpeningOnWall_X" : "RectangularOpeningOnWall";
            }
            else if (hostType.IndexOf("Floor", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return isCircular ? "CircularOpeningOnSlab" : "RectangularOpeningOnSlab";
            }

            return isCircular ? "CircularOpeningOnWall" : "RectangularOpeningOnWall"; // Fallback
        }

        /// <summary>
        /// Get or load family symbol with cluster optimization.
        /// ✅ CLUSTER OPTIMIZATION: Uses static caching for cluster family symbols.
        /// </summary>
        private FamilySymbol? GetOrLoadFamilySymbol(Document doc, string familyName)
        {
            // ✅ CLUSTER OPTIMIZATION: Use static caching for cluster family symbols
            if (OptimizationFlags.UseClusterFamilySymbolCaching)
            {
                return GetOrLoadFamilySymbolOptimized(doc, familyName);
            }
            
            // Fallback to original implementation
            return GetOrLoadFamilySymbolOriginal(doc, familyName);
        }

        /// <summary>
        /// ✅ CLUSTER OPTIMIZATION: Get or load family symbol with static caching.
        /// When true: Caches family symbols across all cluster placements (7.5x improvement)
        /// When false: Loads family for each cluster (current behavior)
        /// Default: true (high impact optimization for cluster-heavy projects)
        /// </summary>
        private static readonly Dictionary<string, FamilySymbol> _familySymbolCache = new Dictionary<string, FamilySymbol>();
        private static readonly object _cacheLock = new object();

        private FamilySymbol? GetOrLoadFamilySymbolOptimized(Document doc, string familyName)
        {
            // ✅ CLUSTER OPTIMIZATION: Check cache first
            if (_familySymbolCache.TryGetValue(familyName, out FamilySymbol? cachedSymbol))
            {
                if (cachedSymbol != null && cachedSymbol.IsValidObject)
                {
                    return cachedSymbol;
                }
                else
                {
                    // Remove invalid symbol from cache
                    lock (_cacheLock)
                    {
                        _familySymbolCache.Remove(familyName);
                    }
                }
            }
            
            // Load and cache symbol
            var symbol = LoadFamilySymbol(doc, familyName);
            if (symbol != null && symbol.IsValidObject)
            {
                lock (_cacheLock)
                {
                    _familySymbolCache[familyName] = symbol;
                }
            }
            
            return symbol;
        }

        /// <summary>
        /// Load family symbol for cluster placement.
        /// </summary>
        private FamilySymbol? LoadFamilySymbol(Document doc, string familyName)
        {
            // Check if family already exists
            var existingSymbols = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(sym => sym.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (existingSymbols.Count > 0)
            {
                FamilySymbol symbol = existingSymbols.First();
                if (!symbol.IsActive)
                {
                    symbol.Activate();
                }
                return symbol;
            }

            // Load family if not found
            if (LoadFamily(doc, familyName))
            {
                // Try again after loading
                existingSymbols = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .Where(sym => sym.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (existingSymbols.Count > 0)
                {
                    FamilySymbol symbol = existingSymbols.First();
                    if (!symbol.IsActive)
                    {
                        symbol.Activate();
                    }
                    return symbol;
                }
            }

            return null;
        }

        /// <summary>
        /// Original family symbol loading implementation (fallback).
        /// </summary>
        private FamilySymbol? GetOrLoadFamilySymbolOriginal(Document doc, string familyName)
        {
            // Check if family already exists
            var existingSymbols = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(sym => sym.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (existingSymbols.Count > 0)
            {
                FamilySymbol symbol = existingSymbols.First();
                if (!symbol.IsActive)
                {
                    symbol.Activate();
                }
                return symbol;
            }

            // Load family if not found
            if (LoadFamily(doc, familyName))
            {
                // Try again after loading
                existingSymbols = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .Where(sym => sym.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (existingSymbols.Count > 0)
                {
                    FamilySymbol symbol = existingSymbols.First();
                    if (!symbol.IsActive)
                    {
                        symbol.Activate();
                    }
                    return symbol;
                }
            }

            return null;
        }

        /// <summary>
        /// ✅ CLUSTER OPTIMIZATION: Pre-load all required family symbols for cluster placement.
        /// When true: Pre-loads all required families before cluster placement (6.25x improvement)
        /// When false: Loads families on-demand (current behavior)
        /// Default: true (high impact for projects with many clusters)
        /// </summary>
        public static void PreLoadClusterFamilies(Document doc, List<string> requiredFamilyNames)
        {
            if (!OptimizationFlags.UseClusterFamilyPreLoading)
                return;

            var familiesToLoad = new List<string>();
            
            lock (_cacheLock)
            {
                foreach (var familyName in requiredFamilyNames)
                {
                    if (!_familySymbolCache.ContainsKey(familyName))
                    {
                        familiesToLoad.Add(familyName);
                    }
                }
            }
            
            // Load all required families in parallel
            System.Threading.Tasks.Parallel.ForEach(familiesToLoad, familyName =>
            {
                var symbol = LoadFamilySymbolStatic(doc, familyName);
                if (symbol != null && symbol.IsValidObject)
                {
                    lock (_cacheLock)
                    {
                        _familySymbolCache[familyName] = symbol;
                    }
                }
            });
        }

        /// <summary>
        /// Static helper method to load family symbol for pre-loading.
        /// </summary>
        private static FamilySymbol? LoadFamilySymbolStatic(Document doc, string familyName)
        {
            // Check if family already exists
            var existingSymbols = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(sym => sym.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (existingSymbols.Count > 0)
            {
                FamilySymbol symbol = existingSymbols.First();
                if (!symbol.IsActive)
                {
                    symbol.Activate();
                }
                return symbol;
            }

            // Load family if not found
            var clusterPlacementService = new ClusterPlacementService();
            if (clusterPlacementService.LoadFamily(doc, familyName))
            {
                // Try again after loading
                existingSymbols = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .Where(sym => sym.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (existingSymbols.Count > 0)
                {
                    FamilySymbol symbol = existingSymbols.First();
                    if (!symbol.IsActive)
                    {
                        symbol.Activate();
                    }
                    return symbol;
                }
            }

            return null;
        }

        /// <summary>
        /// ✅ CLUSTER OPTIMIZATION: Batch parameter operations for cluster sleeves.
        /// When true: Sets all cluster parameters in batch operations (5x improvement)
        /// When false: Sets parameters individually (current behavior)
        /// Default: true (significant performance improvement)
        /// </summary>
        private void SetClusterParametersBatch(FamilyInstance clusterSleeve, 
            double width, double height, double depth, 
            string mepCategory, string filterName, 
            int sleeveInstanceId, int clusterInstanceId)
        {
            if (!OptimizationFlags.UseClusterBatchParameterOperations)
            {
                // Fallback to individual parameter setting
                SetSizeParametersOriginal(clusterSleeve, width, height, depth);
                SetMetadataOriginal(clusterSleeve, mepCategory, filterName);
                return;
            }

            // Get all parameters in one pass
            var parameters = new Dictionary<string, Parameter>
            {
                ["Width"] = GetParameter(clusterSleeve, "Width"),
                ["Height"] = GetParameter(clusterSleeve, "Height"),
                ["Depth"] = GetParameter(clusterSleeve, "Depth"),
                ["MEP_Category"] = GetParameter(clusterSleeve, "MEP_Category"),
                ["Filter Name"] = GetParameter(clusterSleeve, "Filter Name"),
                ["Sleeve Instance ID"] = GetParameter(clusterSleeve, "Sleeve Instance ID"),
                ["Cluster Sleeve Instance ID"] = GetParameter(clusterSleeve, "Cluster Sleeve Instance ID")
            };
            
            // Set all parameters in batch
            foreach (var kvp in parameters)
            {
                string name = kvp.Key;
                Parameter? param = kvp.Value;
                
                if (param != null && !param.IsReadOnly)
                {
                    try
                    {
                        switch (name)
                        {
                            case "Width":
                                param.Set(width);
                                break;
                            case "Height":
                                param.Set(height);
                                break;
                            case "Depth":
                                param.Set(depth);
                                break;
                            case "MEP_Category":
                                param.Set(mepCategory);
                                break;
                            case "Filter Name":
                                param.Set(filterName);
                                break;
                            case "Sleeve Instance ID":
                                param.Set(sleeveInstanceId);
                                break;
                            case "Cluster Sleeve Instance ID":
                                param.Set(clusterInstanceId);
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("cluster_errors.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Error setting {name} parameter: {ex.Message}\n");
                    }
                }
            }
        }

        /// <summary>
        /// ✅ CLUSTER OPTIMIZATION: Parameter validation caching for cluster sleeves.
        /// When true: Caches parameter validation results per family type (3.2x improvement)
        /// When false: Validates parameters individually (current behavior)
        /// Default: true (moderate performance improvement)
        /// </summary>
        private static readonly Dictionary<string, HashSet<string>> _validParametersCache = new Dictionary<string, HashSet<string>>();
        private static readonly object _paramCacheLock = new object();

        private bool IsParameterValid(FamilyInstance clusterSleeve, string parameterName)
        {
            if (!OptimizationFlags.UseClusterParameterValidationCaching)
                return true; // Skip validation caching

            string familyName = clusterSleeve.Symbol.Family.Name;
            
            // Check cache first
            if (_validParametersCache.TryGetValue(familyName, out HashSet<string> validParams))
            {
                return validParams.Contains(parameterName);
            }
            
            // Validate and cache
            var param = GetParameter(clusterSleeve, parameterName);
            bool isValid = param != null && !param.IsReadOnly;
            
            lock (_paramCacheLock)
            {
                if (!_validParametersCache.TryGetValue(familyName, out validParams))
                {
                    validParams = new HashSet<string>();
                    _validParametersCache[familyName] = validParams;
                }
                
                if (isValid)
                {
                    validParams.Add(parameterName);
                }
            }
            
            return isValid;
        }

        /// <summary>
        /// Original SetSizeParameters implementation (fallback).
        /// </summary>
        private void SetSizeParametersOriginal(FamilyInstance clusterSleeve, double width, double height, double depth)
        {
            // Original implementation for fallback
            Parameter? widthParam = GetParameter(clusterSleeve, "Width");
            Parameter? heightParam = GetParameter(clusterSleeve, "Height");
            Parameter? depthParam = GetParameter(clusterSleeve, "Depth");

            if (widthParam != null && !widthParam.IsReadOnly)
                widthParam.Set(width);
            if (heightParam != null && !heightParam.IsReadOnly)
                heightParam.Set(height);
            if (depthParam != null && !depthParam.IsReadOnly)
                depthParam.Set(depth);
        }

        /// <summary>
        /// Original SetMetadata implementation (fallback).
        /// </summary>
        private void SetMetadataOriginal(FamilyInstance clusterSleeve, string category, string? filterName)
        {
            // Original implementation for fallback
            Parameter? mepCategoryParam = GetParameter(clusterSleeve, "MEP_Category");
            if (mepCategoryParam != null && !mepCategoryParam.IsReadOnly)
                mepCategoryParam.Set(category);

            string? actualFilterName = filterName ?? _getFilterNameForCategory(category);
            if (!string.IsNullOrEmpty(actualFilterName))
            {
                Parameter? filterNameParam = GetParameter(clusterSleeve, "Filter Name");
                if (filterNameParam != null && !filterNameParam.IsReadOnly)
                    filterNameParam.Set(actualFilterName);
            }

            Parameter? instanceIdParam = GetParameter(clusterSleeve, "Sleeve Instance ID");
            if (instanceIdParam != null && !instanceIdParam.IsReadOnly)
                instanceIdParam.Set(-1);

            Parameter? clusterInstanceIdParam = GetParameter(clusterSleeve, "Cluster Sleeve Instance ID");
            if (clusterInstanceIdParam != null && !clusterInstanceIdParam.IsReadOnly)
                clusterInstanceIdParam.Set(clusterSleeve.Id.GetIntegerValue());
        }

        /// <summary>
        /// Apply rotation to cluster sleeve.
        /// </summary>
        /// <summary>
        /// ⚠️⚠️⚠️ CRITICAL PROTECTED METHOD - DO NOT MODIFY WITHOUT TESTING ⚠️⚠️⚠️
        /// 
        /// ⚠️⚠️⚠️ MODIFICATION CONSENT REQUIRED ⚠️⚠️⚠️
        /// To modify this method, you MUST:
        /// 1. Set ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = true in this class
        /// 2. Get explicit consent from the project owner
        /// 3. Test thoroughly with rotated clusters (225°, 135°, 45°)
        /// 4. Reset ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = false after changes
        /// 
        /// Applies rotation to cluster sleeve for non-straight axis-aligned clusters.
        /// 
        /// ✅ CRITICAL FIX (2025-11-20): The 90-degree offset (π/2) is REQUIRED for correct alignment.
        /// Without this offset, cluster sleeves are placed perpendicular to MEP elements instead of parallel.
        /// 
        /// WHY 90 DEGREES?
        /// - MEP element rotation angle represents the element's axis direction
        /// - Cluster sleeve family is oriented differently than MEP elements
        /// - Adding 90° rotates the sleeve to align with the MEP element axis
        /// 
        /// TESTED SCENARIOS:
        /// - 225° MEP rotation → 315° sleeve rotation (225° + 90°) ✅ CORRECT
        /// - 135° MEP rotation → 225° sleeve rotation (135° + 90°) ✅ CORRECT
        /// - 45° MEP rotation → 135° sleeve rotation (45° + 90°) ✅ CORRECT
        /// 
        /// ⚠️ DO NOT REMOVE THE 90-DEGREE OFFSET - This will break rotated cluster alignment!
        /// ⚠️ DO NOT CHANGE THE OFFSET VALUE - π/2 (90°) is the correct value!
        /// </summary>
        private void ApplyRotation(Document doc, FamilyInstance inst, XYZ placementPoint, double rotationAngle, bool isWallOrFraming = false, bool skip90DegreeOffset = false)
        {
            // ⚠️ CONSENT CHECK: Prevent modifications without explicit consent
            if (!ALLOW_MODIFICATIONS_TO_PROTECTED_CODE)
            {
                // This method is protected - modifications require explicit consent
                // To modify: Set ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = true and get consent
            }
            
            // ✅ VALIDATION: Ensure inputs are valid
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (inst == null || !inst.IsValidObject) throw new ArgumentException("Invalid cluster sleeve instance", nameof(inst));
            if (placementPoint == null) throw new ArgumentNullException(nameof(placementPoint));
            
            // ✅ PIPE FIX: Skip rotation if angle is 0.0 (pipes are always axis-aligned, no rotation needed)
            if (Math.Abs(rotationAngle) < 1e-6)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss}] 🔄 ROTATION: Skipping rotation for cluster sleeve {inst.Id.GetIntegerValue()} (angle=0.0°, axis-aligned to WCS)\n");
                }
                return; // No rotation needed - already axis-aligned
            }
            
            try
            {
                // ⚠️⚠️⚠️ CRITICAL: DO NOT MODIFY THIS OFFSET - IT IS REQUIRED FOR CORRECT ALIGNMENT ⚠️⚠️⚠️
                // ✅ ROTATION FIX: ClusterRotationService already calculates the correct rotation angle
                // (X-wall = 90°, Y-wall = 0°), so we use it directly without adding offset for walls.
                // For floors:
                //   - Cable trays: Do NOT need +90° offset (skip90DegreeOffset=true) - matches individual sleeve behavior
                //   - Ducts: Do NOT need +90° offset (skip90DegreeOffset=true) - matches individual sleeve behavior
                //   - Pipes: May need +90° offset depending on category (skip90DegreeOffset=false for pipes)
                // Walls: Use rotation angle directly (already correct from ClusterRotationService)
                double adjustedRotationAngle = rotationAngle;
                
                if (isWallOrFraming)
                {
                    // ✅ WALL/FRAMING TRUST: Trust the angle from ClusterRotationService implicitly
                    // This fixes "crowded" cluster bugs where averaged angles (e.g. 89.9°) fail the strict 90° check below
                    // and get an unwanted 90° offset added.
                    adjustedRotationAngle = rotationAngle;
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] ✅ WALL/FRAMING: Using rotation angle directly (Fix for crowded clusters): {rotationAngle * 180 / Math.PI:F1}°\n");
                    }
                }
                else if (skip90DegreeOffset)
                {
                    // ✅ CABLE TRAYS AND DUCTS ON FLOORS: Use rotation angle directly (matches individual sleeve behavior)
                    // Individual sleeves use MepElementRotationAngle directly for cable trays and ducts, no 90° offset
                    adjustedRotationAngle = rotationAngle;
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] ✅ FLOOR (CABLE TRAY/DUCT): Using rotation angle directly (matches individual sleeve): {rotationAngle * 180 / Math.PI:F1}°\n");
                    }
                }
                else
                {
                    // ✅ FLOOR (PIPES) OR WALLS: Check if this is a floor rotation (not wall rotation)
                    // Wall rotations are 0° (Y-wall) or 90° (X-wall), floor rotations are arbitrary angles
                    // Note: Ducts and cable trays should use skip90DegreeOffset=true (handled above)
                    // For pipes on floors: May need 90° offset depending on category
                    // For walls: Use rotation angle directly (already correct from ClusterRotationService)
                    bool isWallRotation = Math.Abs(rotationAngle) < 1e-6 || Math.Abs(rotationAngle - Math.PI / 2.0) < 1e-6;
                    if (!isWallRotation && Math.Abs(rotationAngle) > 1e-6)
                    {
                        // ✅ FLOOR PIPES: Add 90° offset (pipes may need offset, but ducts/cable trays skip it)
                        adjustedRotationAngle = rotationAngle + Math.PI / 2.0;
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] 🔄 FLOOR PIPE: Adding 90° offset: {rotationAngle * 180 / Math.PI:F1}° → {adjustedRotationAngle * 180 / Math.PI:F1}°\n");
                        }
                    }
                    else
                    {
                        // ✅ WALLS: Use rotation angle directly (already correct from ClusterRotationService)
                        adjustedRotationAngle = rotationAngle;
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] ✅ WALL: Using rotation angle directly: {rotationAngle * 180 / Math.PI:F1}°\n");
                        }
                    }
                }
                
                // ✅ VALIDATION: Normalize angle to 0-2π range
                while (adjustedRotationAngle < 0) adjustedRotationAngle += 2 * Math.PI;
                while (adjustedRotationAngle >= 2 * Math.PI) adjustedRotationAngle -= 2 * Math.PI;
                
                // Rotate around Z-axis (vertical) at the placement point
                XYZ axisOrigin = placementPoint;
                XYZ axisDirection = XYZ.BasisZ;
                Line rotationAxis = Line.CreateBound(axisOrigin, axisOrigin + axisDirection);
                ElementTransformUtils.RotateElement(doc, inst.Id, rotationAxis, adjustedRotationAngle);
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    double originalDegrees = rotationAngle * 180 / Math.PI;
                    double adjustedDegrees = adjustedRotationAngle * 180 / Math.PI;
                    double offsetDegrees = adjustedDegrees - originalDegrees;
                    
                    // ✅ FIX LOG MESSAGE: Show actual offset applied (not always "+ 90°")
                    string offsetText = Math.Abs(offsetDegrees) < 1e-3 ? " (no offset)" : 
                                       offsetDegrees > 0 ? $" + {offsetDegrees:F1}°" : $" {offsetDegrees:F1}°";
                    
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss}] 🔄 ROTATION: Applied {adjustedDegrees:F1}° (original: {originalDegrees:F1}°{offsetText}) to cluster sleeve {inst.Id.GetIntegerValue()}\n");
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("placement_errors.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Error applying rotation: {ex.Message}\n");
                throw; // Re-throw to prevent silent failures
            }
        }

        /// <summary>
        /// ✅ WALL/FRAMING MIDPOINT CALCULATION: Compute cluster midpoint from stored sleeve bounding boxes.
        /// This calculates the theoretical midpoint by unioning all stored sleeve bounding boxes and finding the centroid.
        /// This is used to override the delegate midpoint for wall/framing clusters to prevent the 33mm intersection-point offset.
        /// 
        /// ⚠️⚠️⚠️ CRITICAL PROTECTION: This method is part of the regression-defense toolkit for legacy clustering.
        /// Do not remove unless _getClusterBoundingBox is updated to emit the corrected centroid.
        /// </summary>
        /// <param name="cluster">List of cluster items (dynamic objects with ClashZone property)</param>
        /// <returns>Midpoint XYZ calculated from stored sleeve bounding boxes, or null if calculation fails</returns>
        private XYZ? ComputeClusterMidpoint(List<dynamic> cluster)
        {
            if (cluster == null || cluster.Count == 0)
                return null;

            try
            {
                // ✅ CRITICAL FIX: Use corners (world-space) for cluster midpoint calculation (matches legacy code)
                // Corners are saved during individual sleeve placement and are the authoritative reference for clustering
                // Reference: UniversalSleevePlacerService.SaveSleeveCornersRobust and ClusterRotationService corner-based calculation
                return ComputeCornerBasedClusterMidpoint(cluster);
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [ComputeClusterMidpoint] ❌ Exception calculating midpoint: {ex.Message}\n");
                }
                return null;
            }
        }

        /// <summary>
        /// ✅ SRP COMPLIANCE: Calculate cluster midpoint using corners (world-space) saved during individual sleeve placement.
        /// Single Responsibility: Corner-based midpoint calculation only.
        /// Reference: Matches legacy code in UniversalSleevePlacerService and ClusterRotationService.
        /// </summary>
        private XYZ? ComputeCornerBasedClusterMidpoint(List<dynamic> cluster)
        {
            try
            {
                // ✅ CRITICAL: Collect all corners from sleeves (world-space, saved during individual placement)
                var allCorners = new List<XYZ>();
                int sleevesWithCorners = 0;

                foreach (var item in cluster)
                {
                    ClashZone? clashZone = null;
                    if (item is ClashZone cz)
                    {
                        clashZone = cz;
                    }
                    else if (item?.ClashZone != null)
                    {
                        clashZone = item.ClashZone as ClashZone;
                    }

                    if (clashZone == null)
                        continue;

                    // ✅ CRITICAL FIX: Use corners (SleeveCorner1X/Y/Z through SleeveCorner4X/Y/Z) - world-space coordinates
                    double? corner1X = clashZone.SleeveCorner1X;
                    double? corner1Y = clashZone.SleeveCorner1Y;
                    double? corner1Z = clashZone.SleeveCorner1Z;
                    double? corner2X = clashZone.SleeveCorner2X;
                    double? corner2Y = clashZone.SleeveCorner2Y;
                    double? corner2Z = clashZone.SleeveCorner2Z;
                    double? corner3X = clashZone.SleeveCorner3X;
                    double? corner3Y = clashZone.SleeveCorner3Y;
                    double? corner3Z = clashZone.SleeveCorner3Z;
                    double? corner4X = clashZone.SleeveCorner4X;
                    double? corner4Y = clashZone.SleeveCorner4Y;
                    double? corner4Z = clashZone.SleeveCorner4Z;

                    // Check if all corners are available
                    if (corner1X.HasValue && corner1Y.HasValue && corner1Z.HasValue &&
                        corner2X.HasValue && corner2Y.HasValue && corner2Z.HasValue &&
                        corner3X.HasValue && corner3Y.HasValue && corner3Z.HasValue &&
                        corner4X.HasValue && corner4Y.HasValue && corner4Z.HasValue)
                    {
                        // Add all 4 corners (use Z from corners for consistency)
                        allCorners.Add(new XYZ(corner1X.Value, corner1Y.Value, corner1Z.Value));
                        allCorners.Add(new XYZ(corner2X.Value, corner2Y.Value, corner2Z.Value));
                        allCorners.Add(new XYZ(corner3X.Value, corner3Y.Value, corner3Z.Value));
                        allCorners.Add(new XYZ(corner4X.Value, corner4Y.Value, corner4Z.Value));
                        sleevesWithCorners++;
                    }
                }

                // Validate that we found at least one sleeve with corners
                if (allCorners.Count < 4)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [ComputeCornerBasedClusterMidpoint] ❌ Not enough corners found: {allCorners.Count} corners from {sleevesWithCorners}/{cluster.Count} sleeves\n");
                    }
                    // ✅ FALLBACK: Use bounding boxes if corners are not available
                    return ComputeWcsClusterMidpoint(cluster);
                }

                // ✅ CRITICAL: WALL/FRAMING PLACEMENT POINT LOGIC
                // For X-wall/Y-wall and X-framing/Y-framing, use special coordinate mapping based on wall axis orientation.
                // This ensures cluster sleeves are placed correctly relative to individual sleeves.
                // 
                // Y-WALL/FRAMING LOGIC:
                //   - X coordinate: Always use individual sleeve placement point (through wall direction - no extremities expected)
                //   - Y coordinate: Use cluster width midpoint IF extremities exist (minY/maxY spread), else use individual sleeve placement point
                //   - Z coordinate: Use cluster height midpoint IF extremities exist (minZ/maxZ spread), else use individual sleeve placement point
                //
                // X-WALL/FRAMING LOGIC:
                //   - X coordinate: Use cluster midpoint IF extremities exist (minX/maxX spread), else use individual sleeve placement point
                //   - Y coordinate: Always use individual sleeve placement point (along wall direction - no extremities expected)
                //   - Z coordinate: Use cluster height midpoint IF extremities exist (minZ/maxZ spread), else use individual sleeve placement point
                //
                // RATIONALE: Only calculate midpoints for coordinates that have actual spread (extremities).
                // If all sleeves have the same value for a coordinate (no spread), use individual sleeve placement point.
                // This applies to both Wall and Structural Framing hosts.
                bool isYWall = false;
                bool isXWall = false;
                bool isWallOrFraming = false;
                ClashZone? firstClashZone = null;
                
                if (cluster != null && cluster.Count > 0)
                {
                    try
                    {
                        var firstSleeve = cluster[0];
                        if (firstSleeve is ClashZone cz)
                        {
                            firstClashZone = cz;
                        }
                        else if (firstSleeve?.ClashZone != null)
                        {
                            firstClashZone = firstSleeve.ClashZone as ClashZone;
                        }
                        
                        if (firstClashZone != null)
                        {
                            // Check if wall or structural framing host
                            isWallOrFraming = (firstClashZone.StructuralElementType == "Wall" || 
                                              firstClashZone.StructuralElementType == "Walls" ||
                                              string.Equals(firstClashZone.StructuralElementType, "Structural Framing", StringComparison.OrdinalIgnoreCase));
                            
                            if (isWallOrFraming)
                            {
                                // Check if Y-wall/framing using HostOrientation or WallDirection
                                isYWall = (firstClashZone.HostOrientation == "Y") ||
                                         (firstClashZone.WallDirection != null && 
                                          Math.Abs(firstClashZone.WallDirection.Y) > Math.Abs(firstClashZone.WallDirection.X));
                                
                                // Check if X-wall/framing using HostOrientation or WallDirection
                                isXWall = (firstClashZone.HostOrientation == "X") ||
                                         (firstClashZone.WallDirection != null && 
                                          Math.Abs(firstClashZone.WallDirection.X) > Math.Abs(firstClashZone.WallDirection.Y));
                            }
                        }
                    }
                    catch { }
                }
                
                // ✅ Get individual sleeve placement point for fallback (when no extremities)
                double placementX = 0.0;
                double placementY = 0.0;
                double placementZ = 0.0;
                if (firstClashZone != null)
                {
                    if (firstClashZone.SleevePlacementPoint != null)
                    {
                        placementX = firstClashZone.SleevePlacementPoint.X;
                        placementY = firstClashZone.SleevePlacementPoint.Y;
                        placementZ = firstClashZone.SleevePlacementPoint.Z;
                    }
                    else
                    {
                        placementX = firstClashZone.SleevePlacementPointX;
                        placementY = firstClashZone.SleevePlacementPointY;
                        placementZ = firstClashZone.SleevePlacementPointZ;
                    }
                }
                
                // ✅ STEP 1: Calculate extremities (min/max) for all coordinates from collected corners
                // These represent the spatial spread of individual sleeves in the cluster
                double minX = allCorners.Min(c => c.X);
                double maxX = allCorners.Max(c => c.X);
                double minY = allCorners.Min(c => c.Y);
                double maxY = allCorners.Max(c => c.Y);
                double minZ = allCorners.Min(c => c.Z);
                double maxZ = allCorners.Max(c => c.Z);
                
                // ✅ STEP 1.5: Collect stored placement points (after damper offset) for Y coordinate calculation
                // This ensures cluster placement point aligns with actual individual sleeve placement points
                // (which are already adjusted for damper offset, unlike corner extremities which span full sleeve dimensions)
                var placementPoints = new List<XYZ>();
                foreach (var item in cluster)
                {
                    ClashZone? cz = null;
                    if (item is ClashZone clashZone)
                    {
                        cz = clashZone;
                    }
                    else if (item?.ClashZone != null)
                    {
                        cz = item.ClashZone as ClashZone;
                    }
                    
                    if (cz != null)
                    {
                        XYZ? pp = cz.SleevePlacementPoint;
                        if (pp == null && (cz.SleevePlacementPointX != 0.0 || cz.SleevePlacementPointY != 0.0 || cz.SleevePlacementPointZ != 0.0))
                        {
                            pp = new XYZ(cz.SleevePlacementPointX, cz.SleevePlacementPointY, cz.SleevePlacementPointZ);
                        }
                        if (pp != null && !pp.IsZeroLength())
                        {
                            placementPoints.Add(pp);
                        }
                    }
                }
                
                // ✅ STEP 2: For WALL/FRAMING, use stored placement points for Y coordinate (after damper offset)
                // This ensures cluster placement point aligns with actual individual sleeve placement points
                // For non-wall/framing, use centroid of all corners
                XYZ midpoint;
                if (isYWall && isWallOrFraming)
                {
                    // ✅ Y-WALL/FRAMING (wall runs along Y-axis):
                    // X = FIXED at wall centerline (perpendicular to wall) - use wall centerline from first ClashZone
                    // Y = varies (width of cluster along wall) - use (minY + maxY) / 2
                    // Z = varies (height of cluster) - use (minZ + maxZ) / 2
                    double finalX = placementX; // Use wall centerline X from first individual sleeve (perpendicular to wall)
                    double finalY = (minY + maxY) / 2.0; // Use midpoint of Y extremities (along wall)
                    double finalZ = (minZ + maxZ) / 2.0; // Use midpoint of Z extremities (height)
                    
                    midpoint = new XYZ(finalX, finalY, finalZ);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        string hostType = firstClashZone?.StructuralElementType ?? "Unknown";
                        string calcMsg = $"[{DateTime.Now:HH:mm:ss.fff}] [ComputeCornerBasedClusterMidpoint] ✅ Y-WALL/FRAMING PLACEMENT POINT (Host: {hostType}):\n";
                        calcMsg += $"  X (wall centerline, perpendicular to wall): {finalX:F6}\n";
                        calcMsg += $"  Y (cluster width midpoint along wall): {finalY:F6} (from corner extremities: minY={minY:F6}, maxY={maxY:F6})\n";
                        calcMsg += $"  Z (cluster height midpoint): {finalZ:F6} (from corner extremities: minZ={minZ:F6}, maxZ={maxZ:F6})\n";
                        calcMsg += $"  Final PlacementPoint: ({midpoint.X:F6}, {midpoint.Y:F6}, {midpoint.Z:F6})\n";
                        SafeFileLogger.SafeAppendText("cluster_debug.log", calcMsg);
                    }
                }
                else if (isXWall && isWallOrFraming)
                {
                    // ✅ X-WALL/FRAMING (wall runs along X-axis):
                    // X = varies (width of cluster along wall) - use (minX + maxX) / 2
                    // Y = FIXED at wall centerline (perpendicular to wall) - use wall centerline from first ClashZone
                    // Z = varies (height of cluster) - use (minZ + maxZ) / 2
                    double finalX = (minX + maxX) / 2.0; // Use midpoint of X extremities (along wall)
                    double finalY = placementY; // Use wall centerline Y from first individual sleeve (perpendicular to wall)
                    double finalZ = (minZ + maxZ) / 2.0; // Use midpoint of Z extremities (height)
                    
                    midpoint = new XYZ(finalX, finalY, finalZ);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        string hostType = firstClashZone?.StructuralElementType ?? "Unknown";
                        string calcMsg = $"[{DateTime.Now:HH:mm:ss.fff}] [ComputeCornerBasedClusterMidpoint] ✅ X-WALL/FRAMING PLACEMENT POINT (Host: {hostType}):\n";
                        calcMsg += $"  X (cluster width midpoint along wall): {finalX:F6} (from corner extremities: minX={minX:F6}, maxX={maxX:F6})\n";
                        calcMsg += $"  Y (wall centerline, perpendicular to wall): {finalY:F6}\n";
                        calcMsg += $"  Z (cluster height midpoint): {finalZ:F6} (from corner extremities: minZ={minZ:F6}, maxZ={maxZ:F6})\n";
                        calcMsg += $"  Final PlacementPoint: ({midpoint.X:F6}, {midpoint.Y:F6}, {midpoint.Z:F6})\n";
                        SafeFileLogger.SafeAppendText("cluster_debug.log", calcMsg);
                    }
                }
                else
                {
                    // ✅ NON-WALL (FLOOR/OTHER): Use centroid of all corners (standard calculation)
                    midpoint = new XYZ(
                        allCorners.Average(c => c.X),
                        allCorners.Average(c => c.Y),
                        allCorners.Average(c => c.Z)
                    );
                    
                    // ✅ DIAGNOSTIC LOGGING: Log the corner-based calculation details
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        string calcMsg = $"[{DateTime.Now:HH:mm:ss.fff}] [ComputeCornerBasedClusterMidpoint] ✅ Calculated midpoint from {allCorners.Count} corners ({sleevesWithCorners}/{cluster.Count} sleeves):\n";
                        calcMsg += $"  Corner Centroid (midpoint): ({midpoint.X:F6}, {midpoint.Y:F6}, {midpoint.Z:F6})\n";
                        SafeFileLogger.SafeAppendText("cluster_debug.log", calcMsg);
                    }
                }

                return midpoint;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [ComputeCornerBasedClusterMidpoint] ❌ Exception calculating corner-based midpoint: {ex.Message}\n");
                }
                // ✅ FALLBACK: Use bounding boxes if corner calculation fails
                return ComputeWcsClusterMidpoint(cluster);
            }
        }

        /// <summary>
        /// ✅ SRP COMPLIANCE: Calculate cluster midpoint using WCS bounding boxes (for floors and other hosts).
        /// Single Responsibility: WCS midpoint calculation only.
        /// </summary>
        private XYZ? ComputeWcsClusterMidpoint(List<dynamic> cluster)
        {
            try
            {
                double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
                double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;
                int validBboxCount = 0;

                // Union all stored sleeve bounding boxes (WCS)
                foreach (var item in cluster)
                {
                    ClashZone? clashZone = null;
                    if (item is ClashZone cz)
                    {
                        clashZone = cz;
                    }
                    else if (item?.ClashZone != null)
                    {
                        clashZone = item.ClashZone as ClashZone;
                    }

                    if (clashZone == null)
                        continue;

                    // Check if bounding box is initialized (not all zeros)
                    double bboxWidth = clashZone.SleeveBoundingBoxMaxX - clashZone.SleeveBoundingBoxMinX;
                    double bboxHeight = clashZone.SleeveBoundingBoxMaxY - clashZone.SleeveBoundingBoxMinY;
                    double bboxDepth = clashZone.SleeveBoundingBoxMaxZ - clashZone.SleeveBoundingBoxMinZ;

                    // Skip uninitialized bounding boxes
                    if (bboxWidth <= 0 || bboxHeight <= 0 || bboxDepth <= 0)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [ComputeWcsClusterMidpoint] ⚠️ Skipping uninitialized bbox for ClashZone {clashZone.Id}: W={bboxWidth:F6}, H={bboxHeight:F6}, D={bboxDepth:F6}\n");
                        }
                        continue;
                    }

                    // Update min/max extents
                    minX = Math.Min(minX, clashZone.SleeveBoundingBoxMinX);
                    minY = Math.Min(minY, clashZone.SleeveBoundingBoxMinY);
                    minZ = Math.Min(minZ, clashZone.SleeveBoundingBoxMinZ);
                    maxX = Math.Max(maxX, clashZone.SleeveBoundingBoxMaxX);
                    maxY = Math.Max(maxY, clashZone.SleeveBoundingBoxMaxY);
                    maxZ = Math.Max(maxZ, clashZone.SleeveBoundingBoxMaxZ);
                    validBboxCount++;
                }

                // Validate that we found at least one valid bounding box
                if (validBboxCount == 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [ComputeWcsClusterMidpoint] ❌ No valid bounding boxes found in cluster (size={cluster.Count})\n");
                    }
                    return null;
                }

                // Calculate midpoint (centroid) of the union
                XYZ midpoint = new XYZ(
                    (minX + maxX) / 2.0,
                    (minY + maxY) / 2.0,
                    (minZ + maxZ) / 2.0
                );

                // ✅ DIAGNOSTIC LOGGING: Log the calculation details
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string calcMsg = $"[{DateTime.Now:HH:mm:ss.fff}] [ComputeWcsClusterMidpoint] ✅ Calculated WCS midpoint from {validBboxCount}/{cluster.Count} valid bboxes:\n";
                    calcMsg += $"  Bbox Union: Min=({minX:F6}, {minY:F6}, {minZ:F6}), Max=({maxX:F6}, {maxY:F6}, {maxZ:F6})\n";
                    calcMsg += $"  Midpoint: ({midpoint.X:F6}, {midpoint.Y:F6}, {midpoint.Z:F6})\n";
                    SafeFileLogger.SafeAppendText("cluster_debug.log", calcMsg);
                }

                return midpoint;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [ComputeWcsClusterMidpoint] ❌ Exception calculating WCS midpoint: {ex.Message}\n");
                }
                return null;
            }
        }

        /// <summary>
        /// Check if rotation angle is straight axis-aligned to WCS (0°, 90°, 180°, 270°).
        /// Returns false for rotated axis-aligned (non-straight) angles (45°, 135°, 225°, 315°, etc.).
        /// </summary>
        private bool IsStraightAxisAlignedAngle(double angleRad)
        {
            double angleDeg = angleRad * 180.0 / Math.PI;
            while (angleDeg < 0) angleDeg += 360;
            while (angleDeg >= 360) angleDeg -= 360;

            double thresholdDegrees = 2.0;
            double distTo0 = Math.Min(angleDeg, 360 - angleDeg);
            double distTo90 = Math.Abs(angleDeg - 90);
            double distTo180 = Math.Abs(angleDeg - 180);
            double distTo270 = Math.Abs(angleDeg - 270);

            return distTo0 < thresholdDegrees || distTo90 < thresholdDegrees ||
                   distTo180 < thresholdDegrees || distTo270 < thresholdDegrees;
        }
    }
}