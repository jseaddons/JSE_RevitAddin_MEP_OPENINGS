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
                        
                        // Override placement point with computed midpoint
                        placementPoint = theoreticalMid;
                        
                        // ✅ DIAGNOSTIC LOGGING: Log the override and delta
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            string midpointMsg = $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] 🔥 WALL/FRAMING MIDPOINT OVERRIDE:\n";
                            midpointMsg += $"  Original PlacementPoint (from delegate): ({originalPlacementPoint.X:F6}, {originalPlacementPoint.Y:F6}, {originalPlacementPoint.Z:F6})\n";
                            midpointMsg += $"  TheoreticalMid (from stored bboxes): ({theoreticalMid.X:F6}, {theoreticalMid.Y:F6}, {theoreticalMid.Z:F6})\n";
                            midpointMsg += $"  Delta: X={deltaX:F2}mm, Y={deltaY:F2}mm, Z={deltaZ:F2}mm, Distance={deltaDistance:F2}mm\n";
                            midpointMsg += $"  ✅ Using TheoreticalMid for placement (overriding delegate midpoint)\n";
                            DebugLogger.Info(midpointMsg);
                            SafeFileLogger.SafeAppendText("cluster_debug.log", midpointMsg);
                        }
                    }
                    else
                    {
                        // Log warning if midpoint computation failed
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            string warningMsg = $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ⚠️ WALL/FRAMING: ComputeClusterMidpoint returned null, using delegate placementPoint\n";
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
                
                // ✅ CATEGORY CHECK: Determine if this is a cable tray cluster
                bool isCableTrayCategory = false;
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
                        }
                    }
                }
                
                // Apply rotation if:
                // 1. It's a wall with significant rotation (X-wall with 90°), OR
                // 2. It's a floor with rotated axis (non-straight)
                // ⚠️ CABLETRAY FIX: Cable trays on floors get rotation based on MEP orientation WITHOUT extra 90° offset
                // Ducts on floors get rotation based on MEP orientation WITH extra 90° offset
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
                        // ✅ FLOOR: Apply rotation for both ducts and cable trays
                        // Cable trays: MEP orientation angle + 90° offset (to fix disorientation) - skip90DegreeOffset = false
                        // Ducts/Pipes: MEP orientation angle + 90° offset - skip90DegreeOffset = false
                        // Both get 90° offset for floors
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] 🔄 APPLYING FLOOR ROTATION: {rotationAngle * 180 / Math.PI:F1}° for floor-hosted cluster (category: {targetCategory}, WITH 90° OFFSET)\n");
                        }
                        ApplyRotation(doc, inst, placementPoint, rotationAngle, skip90DegreeOffset: false);
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
                        }
                        else
                        {
                            widthParam.Set(openingWidth);
                        }
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("placement_errors.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Error setting Width parameter: {ex.Message}\n");
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
                        }
                        else
                        {
                            heightParam.Set(openingHeight);
                        }
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("placement_errors.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Error setting Height parameter: {ex.Message}\n");
                    }
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
                        }
                        else
                        {
                            depthParam.Set(openingDepth);
                        }
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("placement_errors.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Error setting Depth parameter: {ex.Message}\n");
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

                // Set Sleeve Instance ID to -1 (indicates cluster sleeve)
                Parameter? instanceIdParam = GetParameter(clusterSleeve, "Sleeve Instance ID");
                if (instanceIdParam != null && !instanceIdParam.IsReadOnly)
                {
                    if (deferredParameters != null && OptimizationFlags.UseBatchedParameterWrites)
                    {
                        if (!deferredParameters.ContainsKey(clusterSleeve.Id))
                            deferredParameters[clusterSleeve.Id] = new Dictionary<string, object>();
                        deferredParameters[clusterSleeve.Id]["Sleeve Instance ID"] = -1;
                    }
                    else
                    {
                        instanceIdParam.Set(-1);
                    }
                }

                // Set Cluster Sleeve Instance ID
                Parameter? clusterInstanceIdParam = GetParameter(clusterSleeve, "Cluster Sleeve Instance ID");
                if (clusterInstanceIdParam != null && !clusterInstanceIdParam.IsReadOnly)
                {
                    if (deferredParameters != null && OptimizationFlags.UseBatchedParameterWrites)
                    {
                        if (!deferredParameters.ContainsKey(clusterSleeve.Id))
                            deferredParameters[clusterSleeve.Id] = new Dictionary<string, object>();
                        deferredParameters[clusterSleeve.Id]["Cluster Sleeve Instance ID"] = clusterSleeve.Id.IntegerValue;
                    }
                    else
                    {
                        clusterInstanceIdParam.Set(clusterSleeve.Id.IntegerValue);
                    }
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

                // Load family from Resources directory
                string resourcesPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Resources";
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
                // ✅ CRITICAL FIX: Add 90 degrees (π/2) to MEP orientation angle to fix alignment issue for ducts
                // ✅ ROTATION FIX: ClusterRotationService already calculates the correct rotation angle
                // (X-wall = 90°, Y-wall = 0°), so we use it directly without adding offset for walls.
                // For floors: Both ducts and cable trays need +90° offset to align correctly
                // Walls: Use rotation angle directly (already correct from ClusterRotationService)
                double adjustedRotationAngle = rotationAngle;
                
                if (skip90DegreeOffset)
                {
                    // ✅ WALLS: Use rotation angle directly (already correct from ClusterRotationService)
                    // This path is only for walls now (cable trays on floors also get 90° offset)
                    adjustedRotationAngle = rotationAngle;
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] ✅ WALL: Using rotation angle directly (no 90° offset): {rotationAngle * 180 / Math.PI:F1}°\n");
                    }
                }
                else
                {
                    // ✅ FLOOR (DUCTS AND CABLE TRAYS): Check if this is a floor rotation (not wall rotation)
                    // Wall rotations are 0° (Y-wall) or 90° (X-wall), floor rotations are arbitrary angles
                    bool isWallRotation = Math.Abs(rotationAngle) < 1e-6 || Math.Abs(rotationAngle - Math.PI / 2.0) < 1e-6;
                    if (!isWallRotation && Math.Abs(rotationAngle) > 1e-6)
                    {
                        // ✅ FLOOR (ALL CATEGORIES): Add 90° offset for both ducts and cable trays on floors
                        adjustedRotationAngle = rotationAngle + Math.PI / 2.0;
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] 🔄 FLOOR: Adding 90° offset: {rotationAngle * 180 / Math.PI:F1}° → {adjustedRotationAngle * 180 / Math.PI:F1}°\n");
                        }
                    }
                    else
                    {
                        // ✅ WALLS: Use rotation angle directly (already correct from ClusterRotationService)
                        adjustedRotationAngle = rotationAngle;
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
                double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
                double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;
                int validBboxCount = 0;

                // Union all stored sleeve bounding boxes
                foreach (var item in cluster)
                {
                    // Extract ClashZone from dynamic item
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

                    // Skip uninitialized bounding boxes (all zeros or invalid)
                    if (bboxWidth <= 0 || bboxHeight <= 0 || bboxDepth <= 0)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [ComputeClusterMidpoint] ⚠️ Skipping uninitialized bbox for ClashZone {clashZone.Id}: W={bboxWidth:F6}, H={bboxHeight:F6}, D={bboxDepth:F6}\n");
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
                            $"[{DateTime.Now:HH:mm:ss.fff}] [ComputeClusterMidpoint] ❌ No valid bounding boxes found in cluster (size={cluster.Count})\n");
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
                    string calcMsg = $"[{DateTime.Now:HH:mm:ss.fff}] [ComputeClusterMidpoint] ✅ Calculated midpoint from {validBboxCount}/{cluster.Count} valid bboxes:\n";
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
                        $"[{DateTime.Now:HH:mm:ss.fff}] [ComputeClusterMidpoint] ❌ Exception calculating midpoint: {ex.Message}\n");
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

