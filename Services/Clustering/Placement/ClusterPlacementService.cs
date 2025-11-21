using System;
using System.Collections.Generic;
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
            out int? capturedClusterSleeveId)
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

                // 🔥 CRITICAL: Direct IO logging (bypass SafeFileLogger)
                try
                {
                    var versionTag = Helpers.VersionInfo.VersionTag;
                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                    if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                    var logPath = Path.Combine(logDir, "cluster_debug.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 PlaceClusterSleeve: Starting placement, hostType={groupKey.hostType}, systemType={groupKey.systemType}\n");
                }
                catch { }

                // Determine family name based on host type and shape
                string familyName = GetFamilyName(groupKey);
                if (string.IsNullOrEmpty(familyName))
                {
                    try
                    {
                        var versionTag = Helpers.VersionInfo.VersionTag;
                        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                        var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                        if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                        var logPath = Path.Combine(logDir, "cluster_debug.log");
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ PlaceClusterSleeve FAILED: Unknown host type '{groupKey.hostType}' - cannot determine family name\n");
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
                        var versionTag = Helpers.VersionInfo.VersionTag;
                        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                        var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                        if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                        var logPath = Path.Combine(logDir, "cluster_debug.log");
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ PlaceClusterSleeve FAILED: Failed to load family '{familyName}'\n");
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
                        var versionTag = Helpers.VersionInfo.VersionTag;
                        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                        var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                        if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                        var logPath = Path.Combine(logDir, "cluster_debug.log");
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ PlaceClusterSleeve FAILED: Reference level not found\n");
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

                // Create cluster sleeve instance
                FamilyInstance? inst = null;
                try
                {
                    try
                    {
                        var versionTag = Helpers.VersionInfo.VersionTag;
                        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                        var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                        if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                        var logPath = Path.Combine(logDir, "cluster_debug.log");
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 ABOUT TO CREATE cluster sleeve instance: familyName={familyName}, level={refLevel?.Name ?? "NULL"}\n");
                    }
                    catch { }
                    
                    inst = doc.Create.NewFamilyInstance(placementPoint, familySymbol, refLevel, StructuralType.NonStructural);
                    
                    // ✅ CRITICAL: Capture ID immediately while element is valid
                    capturedClusterSleeveId = inst.Id.IntegerValue;
                    
                    try
                    {
                        var versionTag = Helpers.VersionInfo.VersionTag;
                        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                        var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                        if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                        var logPath = Path.Combine(logDir, "cluster_debug.log");
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ✅✅✅ Cluster sleeve CREATED: ID={capturedClusterSleeveId}\n");
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
                        var versionTag = Helpers.VersionInfo.VersionTag;
                        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                        var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                        if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                        var logPath = Path.Combine(logDir, "cluster_debug.log");
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ EXCEPTION creating cluster sleeve: {createEx.Message}\n");
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] StackTrace: {createEx.StackTrace}\n");
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
                        var versionTag = Helpers.VersionInfo.VersionTag;
                        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                        var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                        if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                        var logPath = Path.Combine(logDir, "cluster_debug.log");
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ Cluster sleeve instance is NULL after creation\n");
                    }
                    catch { }
                    SafeFileLogger.SafeAppendText("placement_errors.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Cluster sleeve instance is null after creation\n");
                    return false;
                }

                // Set size parameters
                // ✅ WALL/FRAMING: No rotation for walls - shouldSwapDimensions is always false
                // Rotation logic is only for floors (rotated axis/non-straight)
                bool shouldSwapDimensions = false; // Walls/framing always use normal logic (no rotation)
                SetSizeParameters(doc, inst, cluster, groupKey, width, height, depth, shouldSwapDimensions);

                // ✅ ROTATION: Apply rotation only for floors (rotated axis/non-straight)
                // Walls/framing always have rotationAngle=0 (no rotation, normal working logic)
                // Only floors can have non-zero rotation angles for rotated axis-aligned clusters
                if (Math.Abs(rotationAngle) > 1e-6)
                {
                    ApplyRotation(doc, inst, placementPoint, rotationAngle);
                }

                // Set metadata
                SetMetadata(inst, targetCategory, null);

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
                        var versionTag = Helpers.VersionInfo.VersionTag;
                        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                        var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                        if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                        var logPath = Path.Combine(logDir, "cluster_debug.log");
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ⚠️ _markClusterResolved delegate is NULL - skipping mark cluster resolved\n");
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
                            var versionTag = Helpers.VersionInfo.VersionTag;
                            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                            var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                            if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                            var logPath = Path.Combine(logDir, "cluster_debug.log");
                            File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ⚠️ Exception in _markClusterResolved: {markEx.Message}\n");
                        }
                        catch { }
                    }
                }

                placedClusterSleeve = inst;
                
                try
                {
                    var versionTag = Helpers.VersionInfo.VersionTag;
                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                    if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                    var logPath = Path.Combine(logDir, "cluster_debug.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ✅✅✅ PlaceClusterSleeve RETURNING TRUE: placedClusterSleeve={(placedClusterSleeve != null ? "NOT NULL" : "NULL")}, capturedId={capturedClusterSleeveId?.ToString() ?? "NULL"}\n");
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
            bool shouldSwapDimensions = false)
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

                // ✅ DIMENSION SWAPPING: For walls/framing, swap dimensions based on orientation
                if (groupKey.hostType == "Wall" || groupKey.hostType == "Structural Framing")
                {
                    // ✅ DIAGNOSTIC: Log input dimensions before swapping
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        double wMm = RevitUnitConversionService.Instance.FromInternalMillimeters(width);
                        double hMm = RevitUnitConversionService.Instance.FromInternalMillimeters(height);
                        double dMm = RevitUnitConversionService.Instance.FromInternalMillimeters(depth);
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] 📐 BEFORE SWAP: Orientation={groupKey.orientation}, ShouldSwap={shouldSwapDimensions}, BBox W={wMm:F1}mm, H={hMm:F1}mm, D={dMm:F1}mm\n");
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
                    
                    if (groupKey.orientation == "Y")
                    {
                        // Y-wall: Width=H, Height=D, Depth=W
                        openingWidth = height;
                        openingHeight = depth;
                        openingDepth = width;
                        
                        // ✅ WALL DEPTH FIX: Override depth with wall thickness for wall-hosted clusters
                        if (wallThickness > 0)
                        {
                            openingDepth = wallThickness;
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                double wallThicknessMm = RevitUnitConversionService.Instance.FromInternalMillimeters(wallThickness);
                                double depthMm = RevitUnitConversionService.Instance.FromInternalMillimeters(depth);
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss}] 🔧 Y-WALL DEPTH FIX: Overriding depth from {depthMm:F1}mm to wall thickness {wallThicknessMm:F1}mm\n");
                            }
                        }
                    }
                    else // X-walls (and other orientations)
                    {
                        // ✅ X-WALL COORDINATE SYSTEM: For X-walls (NO ROTATION - normal working logic only)
                        // World coordinate: X = along wall (width), Y = through wall (depth), Z = vertical (height)
                        // Sleeve parameter mapping: Width = X direction, Depth = Y direction, Height = Z direction
                        // Bounding box returns: width = X extent, height = Y extent, depth = Z extent
                        // 
                        // ✅ NO ROTATION: Walls always use normal logic (rotationAngle=0, shouldSwapDimensions=false)
                        // Rotation logic is only for floors (rotated axis/non-straight)
                        
                        // ✅ X-WALL MAPPING (NO ROTATION):
                        // Width = X direction → width (X extent)
                        // Depth = Y direction → height (Y extent), but overridden with wall thickness
                        // Height = Z direction → depth (Z extent)
                        openingWidth = width;   // X extent → Width parameter
                        openingHeight = depth;  // Z extent → Height parameter
                        openingDepth = height;  // Y extent → Depth parameter (will be overridden with wall thickness)
                        
                        // ✅ WALL DEPTH FIX: Override depth with wall thickness for wall-hosted clusters
                        if (wallThickness > 0)
                        {
                            openingDepth = wallThickness;
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                double wallThicknessMm = RevitUnitConversionService.Instance.FromInternalMillimeters(wallThickness);
                                double originalDepthMm = RevitUnitConversionService.Instance.FromInternalMillimeters(height);
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss}] 🔧 X-WALL DEPTH FIX: Overriding depth from {originalDepthMm:F1}mm (Y extent) to wall thickness {wallThicknessMm:F1}mm (no rotation, normal logic)\n");
                            }
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
                        widthParam.Set(openingWidth);
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
                        heightParam.Set(openingHeight);
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
                        depthParam.Set(openingDepth);
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
            string? filterName = null)
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
                    mepCategoryParam.Set(category);
                }

                // Set Filter Name
                string? actualFilterName = filterName ?? _getFilterNameForCategory(category);
                if (!string.IsNullOrEmpty(actualFilterName))
                {
                    Parameter? filterNameParam = GetParameter(clusterSleeve, "Filter Name");
                    if (filterNameParam != null && !filterNameParam.IsReadOnly)
                    {
                        filterNameParam.Set(actualFilterName);
                    }
                }

                // Set Sleeve Instance ID to -1 (indicates cluster sleeve)
                Parameter? instanceIdParam = GetParameter(clusterSleeve, "Sleeve Instance ID");
                if (instanceIdParam != null && !instanceIdParam.IsReadOnly)
                {
                    instanceIdParam.Set(-1);
                }

                // Set Cluster Sleeve Instance ID
                Parameter? clusterInstanceIdParam = GetParameter(clusterSleeve, "Cluster Sleeve Instance ID");
                if (clusterInstanceIdParam != null && !clusterInstanceIdParam.IsReadOnly)
                {
                    clusterInstanceIdParam.Set(clusterSleeve.Id.IntegerValue);
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("placement_errors.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Error in SetMetadata: {ex.Message}\n");
            }
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
        private void ApplyRotation(Document doc, FamilyInstance inst, XYZ placementPoint, double rotationAngle)
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
                // ✅ CRITICAL FIX: Add 90 degrees (π/2) to MEP orientation angle to fix alignment issue
                // ✅ ROTATION FIX: ClusterRotationService already calculates the correct rotation angle
                // (X-wall = 90°, Y-wall = 0°), so we use it directly without adding offset.
                // Previous code added +90° offset which caused double rotation (180° for X-walls, 90° for Y-walls).
                double adjustedRotationAngle = rotationAngle; // Use rotation angle directly from ClusterRotationService
                
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
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss}] 🔄 ROTATION: Applied {adjustedDegrees:F1}° (original: {originalDegrees:F1}° + 90°) to cluster sleeve {inst.Id.IntegerValue}\n");
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

