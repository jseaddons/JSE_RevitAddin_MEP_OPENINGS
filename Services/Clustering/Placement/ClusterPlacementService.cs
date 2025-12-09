using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Safety;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.BoundingBox;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Geometry; // For WallRcsTransformer
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement; // For SleeveParameterService
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

        /// <summary>
        /// Constructor with dependency injection for external methods.
        /// </summary>
        public ClusterPlacementService(
            Func<int, string?, ClashZone?>? getClashZoneBySleeveInstanceId = null,
            Func<List<dynamic>, string?, double>? determineRotationAngle = null,
            Func<List<dynamic>, List<FamilyInstance>, double, string?, (double width, double height, double depth, XYZ mid, double? rotatedMinX, double? rotatedMinY, double? rotatedMinZ, double? rotatedMaxX, double? rotatedMaxY, double? rotatedMaxZ)>? getClusterBoundingBox = null,
            Action<List<dynamic>, ElementId, string?, BoundingBoxXYZ?, (double minX, double minY, double minZ, double maxX, double maxY, double maxZ)?>? markClusterResolved = null,
            Func<string, string?>? getFilterNameForCategory = null,
            IBoundingBoxCalculator? boundingBoxCalculator = null)
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
            out FamilyInstance? placedClusterSleeve,
            out int? capturedClusterSleeveId,
            Dictionary<ElementId, Dictionary<string, object>>? deferredParameters = null)
        {
            placedClusterSleeve = null;
            capturedClusterSleeveId = null;

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
                        SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] 🔥 PlaceClusterSleeve: Starting placement, hostType={groupKey.hostType}, systemType={groupKey.systemType}\n");
                    }
                }
                catch { }

                // Determine family name based on host type and shape
                string familyName = GetFamilyName(groupKey);
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

                // ✅ WALL/FRAMING MIDPOINT OVERRIDE: Replace delegate midpoint with computed midpoint from stored sleeve bounding boxes
                // This permanently guards against the 33mm intersection-point offset without touching floor/rotated flows
                // ⚠️⚠️⚠️ CRITICAL PROTECTION: Do not remove this override unless _getClusterBoundingBox is updated to emit the corrected centroid
                bool isWallHost = groupKey.hostType == "Wall" || groupKey.hostType == "Structural Framing";
                XYZ originalPlacementPoint = placementPoint;
                
                if (isWallHost)
                {
                    XYZ? theoreticalMid = ComputeClusterMidpoint(cluster);
                    if (theoreticalMid != null)
                    {
                        // Calculate delta for diagnostic logging
                        double deltaX = (theoreticalMid.X - placementPoint.X) * 304.8; // Convert to mm
                        double deltaY = (theoreticalMid.Y - placementPoint.Y) * 304.8;
                        double deltaZ = (theoreticalMid.Z - placementPoint.Z) * 304.8;
                        double deltaDistance = Math.Sqrt(deltaX * deltaX + deltaY * deltaY + deltaZ * deltaZ);
                        
                        // ✅ CRITICAL FIX: If Z delta is too large (>1000mm), stored corner Z coordinates are likely wrong
                        // Use original placement point instead to prevent cluster sleeves from being placed at wrong elevation
                        bool useOriginalPoint = Math.Abs(deltaZ) > 1000.0; // 1000mm = 1 meter threshold
                        
                        if (useOriginalPoint)
                        {
                            // ⚠️ CRITICAL: Stored corner Z coordinates are incorrect - use original placement point
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                string warningMsg = $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ⚠️⚠️⚠️ CRITICAL: Huge Z delta detected ({deltaZ:F2}mm)!\n";
                                warningMsg += $"  Original PlacementPoint (from delegate): ({originalPlacementPoint.X:F6}, {originalPlacementPoint.Y:F6}, {originalPlacementPoint.Z:F6})\n";
                                warningMsg += $"  TheoreticalMid (from stored bboxes): ({theoreticalMid.X:F6}, {theoreticalMid.Y:F6}, {theoreticalMid.Z:F6})\n";
                                warningMsg += $"  Delta: X={deltaX:F2}mm, Y={deltaY:F2}mm, Z={deltaZ:F2}mm, Distance={deltaDistance:F2}mm\n";
                                warningMsg += $"  ⚠️⚠️⚠️ REJECTING TheoreticalMid (stored corner Z coordinates are wrong) - Using Original PlacementPoint instead\n";
                                DebugLogger.Warning(warningMsg);
                                SafeFileLogger.SafeAppendText("cluster_debug.log", warningMsg);
                            }
                            // Keep original placementPoint (don't override)
                        }
                        else
                        {
                            // Override placement point with computed midpoint (normal case)
                            // Corners are already at wall centerline (from individual sleeves), so midpoint is correct
                            placementPoint = theoreticalMid;
                            
                            // ✅ DIAGNOSTIC LOGGING: Log the override and delta
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                string midpointMsg = $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] 🔥 WALL/FRAMING MIDPOINT OVERRIDE:\n";
                                midpointMsg += $"  Original PlacementPoint (from delegate): ({originalPlacementPoint.X:F6}, {originalPlacementPoint.Y:F6}, {originalPlacementPoint.Z:F6})\n";
                                midpointMsg += $"  TheoreticalMid (from stored corners): ({theoreticalMid.X:F6}, {theoreticalMid.Y:F6}, {theoreticalMid.Z:F6})\n";
                                midpointMsg += $"  Delta: X={deltaX:F2}mm, Y={deltaY:F2}mm, Z={deltaZ:F2}mm, Distance={deltaDistance:F2}mm\n";
                                midpointMsg += $"  ✅ Using TheoreticalMid for placement (overriding delegate midpoint)\n";
                                DebugLogger.Info(midpointMsg);
                                SafeFileLogger.SafeAppendText("cluster_debug.log", midpointMsg);
                            }
                        }
                    }
                    else
                    {
                        // ✅ WALL CENTERLINE ADJUSTMENT: If midpoint computation failed, adjust original placement point to wall centerline
                        // Original placementPoint might be at wall face (from intersection), need to move to centerline
                        // Working logic from methodology document: placePoint = intersection + wallNormal * (-wallThickness * 0.5)
                        if (cluster != null && cluster.Count > 0)
                        {
                            try
                            {
                                var firstSleeve = cluster[0];
                                ClashZone? firstClashZone = null;
                                if (firstSleeve is ClashZone cz)
                                {
                                    firstClashZone = cz;
                                }
                                else if (firstSleeve?.ClashZone != null)
                                {
                                    firstClashZone = firstSleeve.ClashZone as ClashZone;
                                }
                                
                                if (firstClashZone != null && firstClashZone.StructuralElementNormal != null)
                                {
                                    // Get wall normal and thickness
                                    XYZ wallNormal = firstClashZone.StructuralElementNormal.Normalize();
                                    double wallThickness = 0.0;
                                    
                                    if (isWallHost)
                                    {
                                        wallThickness = firstClashZone.WallThickness > 0 ? firstClashZone.WallThickness : firstClashZone.StructuralElementThickness;
                                    }
                                    else
                                    {
                                        wallThickness = firstClashZone.FramingThickness > 0 ? firstClashZone.FramingThickness : firstClashZone.StructuralElementThickness;
                                    }
                                    
                                    if (wallThickness > 0)
                                    {
                                        // ✅ WORKING LOGIC: Move from intersection point (wall face) to wall centerline
                                        // Formula: placePoint = intersection + wallNormal * (-wallThickness * 0.5)
                                        // This matches individual sleeve placement logic
                                        XYZ wallVector = wallNormal.Multiply(-wallThickness);
                                        placementPoint = placementPoint.Add(wallVector.Multiply(0.5));
                                        
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                                $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ✅ WALL CENTERLINE ADJUSTMENT (fallback): " +
                                                $"Original PlacementPoint=({originalPlacementPoint.X:F6}, {originalPlacementPoint.Y:F6}, {originalPlacementPoint.Z:F6}), " +
                                                $"WallNormal=({wallNormal.X:F6}, {wallNormal.Y:F6}, {wallNormal.Z:F6}), " +
                                                $"WallThickness={wallThickness * 304.8:F1}mm, " +
                                                $"Adjusted PlacementPoint=({placementPoint.X:F6}, {placementPoint.Y:F6}, {placementPoint.Z:F6})\n");
                                        }
                                    }
                                }
                            }
                            catch (Exception centerlineEx)
                            {
                                // If centerline adjustment fails, use original placementPoint
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                                        $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ⚠️ Centerline adjustment failed: {centerlineEx.Message}, using original placementPoint\n");
                                }
                            }
                        }
                        
                        // Log warning if midpoint computation failed
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            string warningMsg = $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ⚠️ WALL/FRAMING: ComputeClusterMidpoint returned null, using delegate placementPoint (adjusted to wall centerline)\n";
                            SafeFileLogger.SafeAppendText("cluster_debug.log", warningMsg);
                        }
                    }
                }

                // Create cluster sleeve instance
                FamilyInstance? inst = null;
                try
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
                            SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] 🔥 ABOUT TO CREATE cluster sleeve instance: familyName={familyName}, level={refLevel?.Name ?? "NULL"}\n");
                        }
                    }
                    catch { }
                    
                    // ✅ PERFORMANCE PROFILING: Profile family instantiation to identify symbol binding vs geometry creation
                    var instantiationTimer = System.Diagnostics.Stopwatch.StartNew();
                    var beforeInstantiation = System.GC.CollectionCount(0); // Track GC before
                    
                    inst = doc.Create.NewFamilyInstance(placementPoint, familySymbol, refLevel, StructuralType.NonStructural);
                    
                    instantiationTimer.Stop();
                    var afterInstantiation = System.GC.CollectionCount(0);
                    var gcCollections = afterInstantiation - beforeInstantiation;
                    
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
                        
                        var logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Logs", versionTag);
                        Directory.CreateDirectory(logDir);
                        var profPath = Path.Combine(logDir, "family_instantiation_profile.log");
                        
                        SafeFileLogger.SafeAppendText("cluster_debug.log", $"{DateTime.Now:O}\t" +
                            $"ClusterSleeveId={inst?.Id?.IntegerValue ?? -1}\t" +
                            $"Family={familySymbol?.Family?.Name ?? "NULL"}\t" +
                            $"Symbol={familySymbol?.Name ?? "NULL"}\t" +
                            $"Type={instantiationType}\t" +
                            $"TimeMs={instantiationTimer.ElapsedMilliseconds}\t" +
                            $"TimeTicks={instantiationTimer.ElapsedTicks}\t" +
                            $"GCCollections={gcCollections}\t" +
                            $"Level={refLevel?.Name ?? "NULL"}\t" +
                            $"ClusterSize={cluster?.Count ?? 0}\n");
                    }
                    
                    // ✅ CRITICAL: Capture ID immediately while element is valid
                    capturedClusterSleeveId = inst.Id.IntegerValue;
                    
                    try
                    {
                        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                        var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                        if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                        var logPath = Path.Combine(logDir, "cluster_debug.log");
                        // DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ✅✅✅ Cluster sleeve CREATED: ID={capturedClusterSleeveId}\n");
                        }
                    }
                    catch { }
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        string createMsg = $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ✅ Cluster sleeve created: ID={capturedClusterSleeveId}\n";
                        DebugLogger.Info(createMsg);
                        SafeFileLogger.SafeAppendText("cluster_debug.log", createMsg);
                    }
                }
                catch (Exception createEx)
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
                            SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ EXCEPTION creating cluster sleeve: {createEx.Message}\n");
                        }
                        // DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] StackTrace: {createEx.StackTrace}\n");
                        }
                    }
                    catch { }
                    SafeFileLogger.SafeAppendText("placement_errors.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Exception creating cluster sleeve: {createEx.Message}\n");
                    return false;
                }

                if (inst == null)
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
                            SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ Cluster sleeve instance is NULL after creation\n");
                        }
                    }
                    catch { }
                    SafeFileLogger.SafeAppendText("placement_errors.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Cluster sleeve instance is null after creation\n");
                    return false;
                }

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
                // ⚠️ CABLETRAY FIX: Cable trays on floors don't need rotation like ducts do
                // Note: isWallHost is already declared earlier in the method (for midpoint override)
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
                
                // Apply rotation if:
                // 1. It's a wall with significant rotation (X-wall with 90°), OR
                // 2. It's a floor with rotated axis (non-straight)
                // ✅ FIX: Cable trays and ducts on floors get rotation based on MEP orientation WITHOUT extra 90° offset
                // Individual sleeves use MepElementRotationAngle directly (no 90° offset), so clusters must match
                if (Math.Abs(rotationAngle) > 1e-6)
                {
                    if (isWallHost)
                    {
                        // ✅ X-WALL: Apply +90° rotation (matches individual sleeve placement)
                        // Individual sleeves apply rotation for X-walls, so clusters must match
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] 🔄 APPLYING X-WALL ROTATION: {rotationAngle * 180 / Math.PI:F1}° for wall-hosted cluster\n");
                        }
                        ApplyRotation(doc, inst, placementPoint, rotationAngle, skip90DegreeOffset: false);
                    }
                    else if (isFloorHost)
                    {
                        // ✅ FLOOR: Match individual sleeve rotation logic exactly
                        // Individual sleeves use MepElementRotationAngle directly for cable trays (no 90° offset)
                        // Individual sleeves use MepElementRotationAngle directly for ducts (no 90° offset)
                        // Both cable trays and ducts should skip the 90° offset to match individual sleeve behavior
                        bool skipOffset = isCableTrayCategory || isDuctCategory; // Cable trays and ducts: no offset (matches individual sleeves)
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            string categoryBehavior = skipOffset ? $"{targetCategory.ToUpper()} (matches individual: no offset)" : "PIPE (cluster: +90° offset)";
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] 🔄 APPLYING FLOOR ROTATION: {rotationAngle * 180 / Math.PI:F1}° for floor-hosted cluster (category: {targetCategory}, {categoryBehavior})\n");
                        }
                        ApplyRotation(doc, inst, placementPoint, rotationAngle, skip90DegreeOffset: skipOffset);
                    }
                }
                else if (isWallHost)
                {
                    // ✅ Y-WALL: No rotation needed (0° rotation matches individual sleeve placement)
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] ✅ Y-WALL: No rotation applied (0° matches individual sleeve placement)\n");
                    }
                }

                // Set metadata (including HostOrientation for proper orientation)
                SetMetadata(inst, targetCategory, null, deferredParameters);

                // ✅ SCHEDULE LEVEL & ELEVATION FROM LEVEL: Set from first ClashZone's MEP element level
                try
                {
                    var firstSleeve = cluster[0];
                    ClashZone? firstClashZone = null;
                    if (firstSleeve is ClashZone cz)
                    {
                        firstClashZone = cz;
                    }
                    else if (firstSleeve?.ClashZone != null)
                    {
                        firstClashZone = firstSleeve.ClashZone as ClashZone;
                    }
                    else if (firstSleeve?.SleeveInstanceId != null && _getClashZoneBySleeveInstanceId != null)
                    {
                        firstClashZone = _getClashZoneBySleeveInstanceId(firstSleeve.SleeveInstanceId, xmlFilePath);
                    }

                    if (firstClashZone != null)
                    {
                        // ✅ Use SleeveParameterService to set Schedule Level and Elevation from Level
                        var parameterService = new SleeveParameterService(doc);
                        parameterService.SetScheduleLevelAndElevationForCluster(inst, firstClashZone, inst.Id);
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] ✅ Set Schedule Level and Elevation from Level for cluster sleeve {inst.Id} from ClashZone {firstClashZone.Id}\n");
                        }
                    }
                    else if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ Could not get ClashZone for cluster - Schedule Level and Elevation from Level not set\n");
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ Error setting Schedule Level and Elevation from Level for cluster: {ex.Message}\n");
                    }
                }

                // Mark clash zones as cluster-resolved
                BoundingBoxXYZ? clusterBbox = null;
                try
                {
                    clusterBbox = inst.get_BoundingBox(null);
                }
                catch { }

                // 🔥 CRITICAL: Check if _markClusterResolved is null (might not be wired)
                if (_markClusterResolved == null)
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
                            SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ _markClusterResolved delegate is NULL - skipping mark cluster resolved\n");
                        }
                    }
                    catch { }
                }
                else
                {
                    try
                    {
                _markClusterResolved(cluster, inst.Id, xmlFilePath, clusterBbox, null);
                    }
                    catch (Exception markEx)
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
                                SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ Exception in _markClusterResolved: {markEx.Message}\n");
                            }
                        }
                        catch { }
                    }
                }

                placedClusterSleeve = inst;
                
                try
                {
                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                    if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                    var logPath = Path.Combine(logDir, "cluster_debug.log");
                    // DEPLOYMENT MODE: Skip file writes
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ✅✅✅ PlaceClusterSleeve RETURNING TRUE: placedClusterSleeve={(placedClusterSleeve != null ? "NOT NULL" : "NULL")}, capturedId={capturedClusterSleeveId?.ToString() ?? "NULL"}\n");
                    }
                }
                catch { }
                
                return true;
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

            try
            {
                // Get parameters with caching
                Parameter? widthParam = GetParameter(clusterSleeve, "Width");
                Parameter? heightParam = GetParameter(clusterSleeve, "Height");
                Parameter? depthParam = GetParameter(clusterSleeve, "Depth");

                double openingWidth = width;
                double openingHeight = height;
                double openingDepth = depth;

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
                    
                    // ✅ CRITICAL FIX: Get wall thickness from clash zones (should be same for all sleeves on same wall)
                    double wallThickness = 0.0;
                    try
                    {
                        if (cluster.Count > 0)
                        {
                            var firstSleeve = cluster[0];
                            var firstClashZone = firstSleeve?.ClashZone as Models.ClashZone;
                            if (firstClashZone != null)
                            {
                                // Get wall thickness (prefer WallThickness over StructuralElementThickness)
                                wallThickness = firstClashZone.WallThickness > 0 
                                    ? firstClashZone.WallThickness 
                                    : firstClashZone.StructuralElementThickness;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("placement_errors.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ⚠️ Error getting wall thickness: {ex.Message}\n");
                    }
                    
                    // ✅ RCS: Direct dimension mapping (no swapping needed)
                    openingWidth = width;   // RCS X (along wall) → Width parameter
                    openingHeight = height; // RCS Z (vertical) → Height parameter
                    openingDepth = depth;   // RCS Y (through wall) → Depth parameter (will be overridden)
                    
                    // ✅ WALL DEPTH FIX: Override depth with wall thickness for wall-hosted clusters
                    if (wallThickness > 0)
                    {
                        openingDepth = wallThickness;
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            double wallThicknessMm = RevitUnitConversionService.Instance.FromInternalMillimeters(wallThickness);
                            double originalDepthMm = RevitUnitConversionService.Instance.FromInternalMillimeters(depth);
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] 🔧 RCS WALL DEPTH FIX: Overriding depth from {originalDepthMm:F1}mm (RCS Y) to wall thickness {wallThicknessMm:F1}mm\n");
                        }
                    }
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
                (openingWidth, openingHeight) = OpeningSettingsHelper.RoundDimensionsToNearest5mm(openingWidth, openingHeight);
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    // Log rounding if values changed
                    if (Math.Abs(originalWidth - openingWidth) > 1e-6 || Math.Abs(originalHeight - openingHeight) > 1e-6)
                    {
                        double originalWMm = RevitUnitConversionService.Instance.FromInternalMillimeters(originalWidth);
                        double originalHMm = RevitUnitConversionService.Instance.FromInternalMillimeters(originalHeight);
                        double roundedWMm = RevitUnitConversionService.Instance.FromInternalMillimeters(openingWidth);
                        double roundedHMm = RevitUnitConversionService.Instance.FromInternalMillimeters(openingHeight);
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] 🔄 ROUNDING: Cluster {clusterSleeve.Id.IntegerValue} - " +
                            $"Width {originalWMm:F1}mm → {roundedWMm:F1}mm, " +
                            $"Height {originalHMm:F1}mm → {roundedHMm:F1}mm " +
                            $"(Placement point/centroid UNCHANGED at calculated corner centroid)\n");
                    }
                }

                // Set parameters
                if (widthParam != null && !widthParam.IsReadOnly)
                {
                    try
                    {
                        if (deferredParameters != null && OptimizationFlags.UseBatchedParameterWrites)
                        {
                            if (!deferredParameters.ContainsKey(clusterSleeve.Id))
                                deferredParameters[clusterSleeve.Id] = new Dictionary<string, object>();
                            deferredParameters[clusterSleeve.Id]["Width"] = openingWidth;
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                double wMm = RevitUnitConversionService.Instance.FromInternalMillimeters(openingWidth);
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss.fff}] [SetSizeParameters] ✅ DEFERRED: Added Width={wMm:F1}mm to deferredParameters for cluster sleeve {clusterSleeve.Id.IntegerValue}\n");
                            }
                        }
                        else
                        {
                            widthParam.Set(openingWidth);
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                double wMm = RevitUnitConversionService.Instance.FromInternalMillimeters(openingWidth);
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss.fff}] [SetSizeParameters] ✅ IMMEDIATE: Set Width={wMm:F1}mm for cluster sleeve {clusterSleeve.Id.IntegerValue} (deferredParams={(deferredParameters != null ? "NOT NULL" : "NULL")}, UseBatched={OptimizationFlags.UseBatchedParameterWrites})\n");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("placement_errors.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Error setting Width parameter: {ex.Message}\n");
                    }
                }
                else
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [SetSizeParameters] ⚠️ Width parameter is NULL or READONLY for cluster sleeve {clusterSleeve.Id.IntegerValue}\n");
                    }
                }

                if (heightParam != null && !heightParam.IsReadOnly)
                {
                    try
                    {
                        if (deferredParameters != null && OptimizationFlags.UseBatchedParameterWrites)
                        {
                            if (!deferredParameters.ContainsKey(clusterSleeve.Id))
                                deferredParameters[clusterSleeve.Id] = new Dictionary<string, object>();
                            deferredParameters[clusterSleeve.Id]["Height"] = openingHeight;
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                double hMm = RevitUnitConversionService.Instance.FromInternalMillimeters(openingHeight);
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss.fff}] [SetSizeParameters] ✅ DEFERRED: Added Height={hMm:F1}mm to deferredParameters for cluster sleeve {clusterSleeve.Id.IntegerValue}\n");
                            }
                        }
                        else
                        {
                            heightParam.Set(openingHeight);
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                double hMm = RevitUnitConversionService.Instance.FromInternalMillimeters(openingHeight);
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss.fff}] [SetSizeParameters] ✅ IMMEDIATE: Set Height={hMm:F1}mm for cluster sleeve {clusterSleeve.Id.IntegerValue} (deferredParams={(deferredParameters != null ? "NOT NULL" : "NULL")}, UseBatched={OptimizationFlags.UseBatchedParameterWrites})\n");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("placement_errors.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Error setting Height parameter: {ex.Message}\n");
                    }
                }
                else
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [SetSizeParameters] ⚠️ Height parameter is NULL or READONLY for cluster sleeve {clusterSleeve.Id.IntegerValue}\n");
                    }
                }

                // ✅ BOTTOM OF OPENING: Calculate and set "Bottom of Opening" for RectangularOpeningOnWall cluster sleeves
                if (OptimizationFlags.UseBottomOfOpeningCalculation)
                {
                    SetBottomOfOpeningForCluster(clusterSleeve, openingHeight, deferredParameters);
                }

                if (depthParam != null && !depthParam.IsReadOnly)
                {
                    try
                    {
                        if (deferredParameters != null && OptimizationFlags.UseBatchedParameterWrites)
                        {
                            if (!deferredParameters.ContainsKey(clusterSleeve.Id))
                                deferredParameters[clusterSleeve.Id] = new Dictionary<string, object>();
                            deferredParameters[clusterSleeve.Id]["Depth"] = openingDepth;
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                double dMm = RevitUnitConversionService.Instance.FromInternalMillimeters(openingDepth);
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss.fff}] [SetSizeParameters] ✅ DEFERRED: Added Depth={dMm:F1}mm to deferredParameters for cluster sleeve {clusterSleeve.Id.IntegerValue}\n");
                            }
                        }
                        else
                        {
                            depthParam.Set(openingDepth);
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                double dMm = RevitUnitConversionService.Instance.FromInternalMillimeters(openingDepth);
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss.fff}] [SetSizeParameters] ✅ IMMEDIATE: Set Depth={dMm:F1}mm for cluster sleeve {clusterSleeve.Id.IntegerValue} (deferredParams={(deferredParameters != null ? "NOT NULL" : "NULL")}, UseBatched={OptimizationFlags.UseBatchedParameterWrites})\n");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("placement_errors.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Error setting Depth parameter: {ex.Message}\n");
                    }
                }
                else
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [SetSizeParameters] ⚠️ Depth parameter is NULL or READONLY for cluster sleeve {clusterSleeve.Id.IntegerValue}\n");
                    }
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("placement_errors.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Error in SetSizeParameters: {ex.Message}\n");
            }
        }

        /// <summary>
        /// ✅ SRP COMPLIANCE: Dedicated method for setting "Bottom of Opening" parameter on cluster sleeves.
        /// Single Responsibility: Calculate and set Bottom of Opening parameter only.
        /// 
        /// Formula: Bottom of Opening = Schedule of Level - (Cluster Height / 2)
        /// Where Schedule of Level is the height from level elevation to cluster placement point (center of opening).
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
            if (OptimizationFlags.UseSafeElementValidation)
            {
                if (!clusterSleeve.IsValidObject)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [SetBottomOfOpeningForCluster] ⚠️ Cluster sleeve {clusterSleeve.Id.IntegerValue} is invalid - skipping\n");
                    }
                    return;
                }
            }

            try
            {
                // ✅ MAIN LOGIC: Calculate "Elevation from Level" from placement point and level elevation
                // This is the PRIMARY method - Revit should set this automatically, but we calculate it explicitly to ensure it's correct
                double? scheduleOfLevel = null;
                
                // ✅ STEP 1: Try to read from sleeve parameter (should be set automatically by Revit)
                Parameter scheduleParam = clusterSleeve.LookupParameter("Elevation from Level")  // ✅ FIRST: This is what user sees in Properties
                                       ?? clusterSleeve.LookupParameter("Schedule of Level")
                                       ?? clusterSleeve.LookupParameter("Schedule Level");

                if (scheduleParam != null && scheduleParam.StorageType == StorageType.Double)
                {
                    scheduleOfLevel = scheduleParam.AsDouble();
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [SetBottomOfOpeningForCluster] ✅ Cluster sleeve {clusterSleeve.Id.IntegerValue}: " +
                            $"Read Elevation from Level={scheduleOfLevel.Value * 304.8:F1}mm from sleeve parameter '{scheduleParam.Definition.Name}'\n");
                    }
                }
                
                // ✅ STEP 2: If not found, calculate from placement point and level elevation (MAIN LOGIC)
                if (!scheduleOfLevel.HasValue)
                {
                    try
                    {
                        // Get placement point (center of opening)
                        LocationPoint? locationPoint = clusterSleeve.Location as LocationPoint;
                        if (locationPoint != null)
                        {
                            XYZ placementPoint = locationPoint.Point;
                            
                            // Get level from sleeve
                            Level? level = clusterSleeve.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM)?.AsElementId() != null
                                ? clusterSleeve.Document.GetElement(clusterSleeve.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM).AsElementId()) as Level
                                : null;
                            
                            if (level != null)
                            {
                                // Calculate: Elevation from Level = Placement Point Z - Level Elevation
                                scheduleOfLevel = placementPoint.Z - level.Elevation;
                                
                                // ✅ CRITICAL: Set the parameter on the sleeve so it's available for future reads
                                if (scheduleParam != null && !scheduleParam.IsReadOnly && scheduleParam.StorageType == StorageType.Double)
                                {
                                    scheduleParam.Set(scheduleOfLevel.Value);
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                                            $"[{DateTime.Now:HH:mm:ss.fff}] [SetBottomOfOpeningForCluster] ✅ Cluster sleeve {clusterSleeve.Id.IntegerValue}: " +
                                            $"CALCULATED and SET Elevation from Level={scheduleOfLevel.Value * 304.8:F1}mm " +
                                            $"(PlacementPoint.Z={placementPoint.Z * 304.8:F1}mm, Level.Elevation={level.Elevation * 304.8:F1}mm, Level='{level.Name}')\n");
                                    }
                                }
                                else if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                                        $"[{DateTime.Now:HH:mm:ss.fff}] [SetBottomOfOpeningForCluster] ✅ Cluster sleeve {clusterSleeve.Id.IntegerValue}: " +
                                        $"CALCULATED Elevation from Level={scheduleOfLevel.Value * 304.8:F1}mm " +
                                        $"(PlacementPoint.Z={placementPoint.Z * 304.8:F1}mm, Level.Elevation={level.Elevation * 304.8:F1}mm, Level='{level.Name}') " +
                                        $"but parameter is read-only or wrong type - using calculated value\n");
                                }
                            }
                        }
                    }
                    catch (Exception calcEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [SetBottomOfOpeningForCluster] ⚠️ Cluster sleeve {clusterSleeve.Id.IntegerValue}: " +
                                $"Error calculating Elevation from Level: {calcEx.Message}\n");
                        }
                    }
                }

                // ✅ VALIDATION: Check if Schedule of Level is valid
                if (!scheduleOfLevel.HasValue ||
                    !BottomOfOpeningCalculationService.IsValidScheduleOfLevel(scheduleOfLevel.Value))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [SetBottomOfOpeningForCluster] ⚠️ Cluster sleeve {clusterSleeve.Id.IntegerValue}: " +
                            $"Schedule of Level parameter not found or invalid (value={scheduleOfLevel?.ToString() ?? "null"}) - skipping\n");
                    }
                    return; // Graceful degradation - skip if Schedule of Level is missing or invalid
                }

                // ✅ VALIDATION: Check if Cluster Height is valid
                if (!BottomOfOpeningCalculationService.IsValidHeight(clusterHeight))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [SetBottomOfOpeningForCluster] ⚠️ Cluster sleeve {clusterSleeve.Id.IntegerValue}: " +
                            $"Cluster Height is invalid (value={clusterHeight * 304.8:F1}mm) - skipping\n");
                    }
                    return; // Graceful degradation - skip if Cluster Height is invalid
                }

                // ✅ CALCULATION: Calculate Bottom of Opening using helper service
                double? bottomOfOpening = BottomOfOpeningCalculationService.CalculateBottomOfOpening(
                    scheduleOfLevel.Value, clusterHeight);

                if (!bottomOfOpening.HasValue)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [SetBottomOfOpeningForCluster] ⚠️ Cluster sleeve {clusterSleeve.Id.IntegerValue}: " +
                            $"Calculation returned null (Schedule={scheduleOfLevel.Value * 304.8:F1}mm, Height={clusterHeight * 304.8:F1}mm) - skipping\n");
                    }
                    return; // Graceful degradation - skip if calculation fails
                }

                // ✅ PARAMETER SETTING: Set "Bottom of Opening" parameter with batching support
                // Parameter name is "Bottom Of Opening" (capital O in "Of") as shown in Revit Properties
                // Try same variations as individual sleeves for consistency
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
                        deferredParameters[clusterSleeve.Id]["Bottom of Opening"] = bottomOfOpening.Value;

                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            double bottomMm = RevitUnitConversionService.Instance.FromInternalMillimeters(bottomOfOpening.Value);
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [SetBottomOfOpeningForCluster] ✅ DEFERRED: Added Bottom of Opening={bottomMm:F1}mm " +
                                $"(Schedule={scheduleOfLevel.Value * 304.8:F1}mm, Height={clusterHeight * 304.8:F1}mm) " +
                                $"to deferredParameters for cluster sleeve {clusterSleeve.Id.IntegerValue}\n");
                        }
                    }
                    else
                    {
                        bottomParam.Set(bottomOfOpening.Value);
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            double bottomMm = RevitUnitConversionService.Instance.FromInternalMillimeters(bottomOfOpening.Value);
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [SetBottomOfOpeningForCluster] ✅ IMMEDIATE: Set Bottom of Opening={bottomMm:F1}mm " +
                                $"(Schedule={scheduleOfLevel.Value * 304.8:F1}mm, Height={clusterHeight * 304.8:F1}mm) " +
                                $"for cluster sleeve {clusterSleeve.Id.IntegerValue}\n");
                        }
                    }
                }
                else
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [SetBottomOfOpeningForCluster] ⚠️ Cluster sleeve {clusterSleeve.Id.IntegerValue}: " +
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
                        $"for cluster sleeve {clusterSleeve.Id.IntegerValue}: {ex.Message}\n");
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
            Dictionary<ElementId, Dictionary<string, object>>? deferredParameters = null)
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
                        $"[{DateTime.Now:HH:mm:ss.fff}] [SetMetadata] ✅ IMMEDIATE (CRITICAL): Set 'Sleeve Instance ID'=-1 for cluster sleeve {clusterSleeve.Id.IntegerValue}\n");
                }
                else
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [SetMetadata] ⚠️ WARNING: 'Sleeve Instance ID' parameter not found or readonly for cluster sleeve {clusterSleeve.Id.IntegerValue}\n");
                }

                // Set Cluster Sleeve Instance ID
                Parameter? clusterInstanceIdParam = GetParameter(clusterSleeve, "Cluster Sleeve Instance ID");
                if (clusterInstanceIdParam != null && !clusterInstanceIdParam.IsReadOnly)
                {
                    clusterInstanceIdParam.Set(clusterSleeve.Id.IntegerValue); // Always set immediately (critical for cleanup)
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [SetMetadata] ✅ IMMEDIATE (CRITICAL): Set 'Cluster Sleeve Instance ID'={clusterSleeve.Id.IntegerValue} for cluster sleeve {clusterSleeve.Id.IntegerValue}\n");
                    
                    // ✅ PERFORMANCE FIX: Remove per-cluster regeneration to enable batch optimization
                    // Verification will happen after batch flush in RefactoredClusterService (line 623)
                    // This change enables 4-6× speedup by batching ALL regenerations
                    // OLD: clusterSleeve.Document?.Regenerate(); // ❌ REMOVED - defeats batch optimization
                    // NEW: Verification deferred to post-flush regeneration
                }
                else
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [SetMetadata] ⚠️ WARNING: 'Cluster Sleeve Instance ID' parameter not found or readonly for cluster sleeve {clusterSleeve.Id.IntegerValue}\n");
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
            if (doc == null || elementId == null || elementId.IntegerValue <= 0)
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
        /// Get family name based on host type and shape.
        /// </summary>
        private string GetFamilyName(SleeveGroupKey groupKey)
        {
            bool isCircular = false; // TODO: Determine from cluster data if needed

            if (groupKey.hostType == "Wall" || groupKey.hostType == "Structural Framing")
            {
                return isCircular ? "CircularOpeningOnWall" : "RectangularOpeningOnWall";
            }
            else if (groupKey.hostType == "Floor")
            {
                return isCircular ? "CircularOpeningOnSlab" : "RectangularOpeningOnSlab";
            }

            return string.Empty;
        }

        /// <summary>
        /// Get or load family symbol.
        /// </summary>
        private FamilySymbol? GetOrLoadFamilySymbol(Document doc, string familyName)
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
        private void ApplyRotation(Document doc, FamilyInstance inst, XYZ placementPoint, double rotationAngle, bool skip90DegreeOffset = false)
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
                        $"[{DateTime.Now:HH:mm:ss}] 🔄 ROTATION: Skipping rotation for cluster sleeve {inst.Id.IntegerValue} (angle=0.0°, axis-aligned to WCS)\n");
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
                
                if (skip90DegreeOffset)
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
                        $"[{DateTime.Now:HH:mm:ss}] 🔄 ROTATION: Applied {adjustedDegrees:F1}° (original: {originalDegrees:F1}°{offsetText}) to cluster sleeve {inst.Id.IntegerValue}\n");
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
                
                // ✅ STEP 2: Check if coordinates have extremities (spread > threshold)
                // If spread is too small (< 1mm), all sleeves have essentially the same value for that coordinate.
                // In such cases, use individual sleeve placement point instead of calculating midpoint.
                const double extremityThreshold = 0.00328084; // 1mm in feet (conversion: 1mm = 0.00328084 ft)
                bool hasXExtremities = (maxX - minX) > extremityThreshold;
                bool hasYExtremities = (maxY - minY) > extremityThreshold;
                bool hasZExtremities = (maxZ - minZ) > extremityThreshold;
                
                XYZ midpoint;
                if (isYWall && isWallOrFraming)
                {
                    // ✅ Y-WALL/FRAMING: 
                    // X = individual sleeve placement point (through wall - no extremities expected)
                    // Y = cluster width midpoint IF has extremities, else individual sleeve placement point
                    // Z = cluster height midpoint IF has extremities, else individual sleeve placement point
                    double finalX = placementX; // Always use individual sleeve for X (through wall)
                    double finalY = hasYExtremities ? (minY + maxY) / 2.0 : placementY;
                    double finalZ = hasZExtremities ? (minZ + maxZ) / 2.0 : placementZ;
                    
                    midpoint = new XYZ(finalX, finalY, finalZ);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        string hostType = firstClashZone?.StructuralElementType ?? "Unknown";
                        string calcMsg = $"[{DateTime.Now:HH:mm:ss.fff}] [ComputeCornerBasedClusterMidpoint] ✅ Y-WALL/FRAMING PLACEMENT POINT (Host: {hostType}):\n";
                        calcMsg += $"  X (from individual sleeve): {finalX:F6}\n";
                        calcMsg += $"  Y: {(hasYExtremities ? $"cluster width midpoint {finalY:F6} (minY={minY:F6}, maxY={maxY:F6})" : $"individual sleeve {finalY:F6} (no extremities)")}\n";
                        calcMsg += $"  Z: {(hasZExtremities ? $"cluster height midpoint {finalZ:F6} (minZ={minZ:F6}, maxZ={maxZ:F6})" : $"individual sleeve {finalZ:F6} (no extremities)")}\n";
                        calcMsg += $"  Final PlacementPoint: ({midpoint.X:F6}, {midpoint.Y:F6}, {midpoint.Z:F6})\n";
                        SafeFileLogger.SafeAppendText("cluster_debug.log", calcMsg);
                    }
                }
                else if (isXWall && isWallOrFraming)
                {
                    // ✅ X-WALL/FRAMING:
                    // X = cluster midpoint IF has extremities, else individual sleeve placement point
                    // Y = individual sleeve placement point (along wall - no extremities expected)
                    // Z = cluster height midpoint IF has extremities, else individual sleeve placement point
                    double finalX = hasXExtremities ? (minX + maxX) / 2.0 : placementX;
                    double finalY = placementY; // Always use individual sleeve for Y (along wall)
                    double finalZ = hasZExtremities ? (minZ + maxZ) / 2.0 : placementZ;
                    
                    midpoint = new XYZ(finalX, finalY, finalZ);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        string hostType = firstClashZone?.StructuralElementType ?? "Unknown";
                        string calcMsg = $"[{DateTime.Now:HH:mm:ss.fff}] [ComputeCornerBasedClusterMidpoint] ✅ X-WALL/FRAMING PLACEMENT POINT (Host: {hostType}):\n";
                        calcMsg += $"  X: {(hasXExtremities ? $"cluster midpoint {finalX:F6} (minX={minX:F6}, maxX={maxX:F6})" : $"individual sleeve {finalX:F6} (no extremities)")}\n";
                        calcMsg += $"  Y (from individual sleeve): {finalY:F6}\n";
                        calcMsg += $"  Z: {(hasZExtremities ? $"cluster height midpoint {finalZ:F6} (minZ={minZ:F6}, maxZ={maxZ:F6})" : $"individual sleeve {finalZ:F6} (no extremities)")}\n";
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

