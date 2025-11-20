using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
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

            // Store dependencies (will throw ArgumentNullException if critical dependencies are null)
            _getClashZoneBySleeveInstanceId = getClashZoneBySleeveInstanceId ?? throw new ArgumentNullException(nameof(getClashZoneBySleeveInstanceId));
            _determineRotationAngle = determineRotationAngle ?? throw new ArgumentNullException(nameof(determineRotationAngle));
            _getClusterBoundingBox = getClusterBoundingBox ?? throw new ArgumentNullException(nameof(getClusterBoundingBox));
            _markClusterResolved = markClusterResolved ?? throw new ArgumentNullException(nameof(markClusterResolved));
            _getFilterNameForCategory = getFilterNameForCategory ?? throw new ArgumentNullException(nameof(getFilterNameForCategory));
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

                // Determine family name based on host type and shape
                string familyName = GetFamilyName(groupKey);
                if (string.IsNullOrEmpty(familyName))
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Unknown host type '{groupKey.hostType}' - cannot determine family name\n");
                    return false;
                }

                // Load family if needed
                FamilySymbol? familySymbol = GetOrLoadFamilySymbol(doc, familyName);
                if (familySymbol == null)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Failed to load family '{familyName}'\n");
                    return false;
                }

                // Get reference level
                Level? refLevel = GetReferenceLevel(doc, cluster[0]);
                if (refLevel == null)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Reference level not found\n");
                    return false;
                }

                // ✅ TRANSACTION SAFETY: Ensure document is modifiable before creating instance
                if (!doc.IsModifiable)
                {
                    SafeFileLogger.SafeAppendText("placement_errors.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Document is not modifiable - cannot create cluster sleeve\n");
                    return false;
                }

                // Create cluster sleeve instance
                FamilyInstance? inst = null;
                try
                {
                    inst = doc.Create.NewFamilyInstance(placementPoint, familySymbol, refLevel, StructuralType.NonStructural);
                    
                    // ✅ CRITICAL: Capture ID immediately while element is valid
                    capturedClusterSleeveId = inst.Id.IntegerValue;
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        string createMsg = $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ✅ Cluster sleeve created: ID={capturedClusterSleeveId}\n";
                        DebugLogger.Info(createMsg);
                        SafeFileLogger.SafeAppendText("cluster_debug.log", createMsg);
                    }
                }
                catch (Exception createEx)
                {
                    SafeFileLogger.SafeAppendText("placement_errors.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Exception creating cluster sleeve: {createEx.Message}\n");
                    return false;
                }

                if (inst == null)
                {
                    SafeFileLogger.SafeAppendText("placement_errors.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Cluster sleeve instance is null after creation\n");
                    return false;
                }

                // Set size parameters
                bool shouldSwapDimensions = (groupKey.hostType == "Wall" || groupKey.hostType == "Structural Framing") && Math.Abs(rotationAngle) > 1e-6;
                SetSizeParameters(doc, inst, cluster, groupKey, width, height, depth, shouldSwapDimensions);

                // Apply rotation if needed (non-axis-aligned)
                if (Math.Abs(rotationAngle) > 1e-6 && !IsAxisAlignedAngle(rotationAngle))
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

                _markClusterResolved(cluster, inst.Id, xmlFilePath, clusterBbox, null);

                placedClusterSleeve = inst;
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
                    if (groupKey.orientation == "Y")
                    {
                        // Y-wall: Width=H, Height=D, Depth=W
                        openingWidth = height;
                        openingHeight = depth;
                        openingDepth = width;
                    }
                    else if (shouldSwapDimensions) // X-walls
                    {
                        // X-wall: Width=W, Height=D, Depth=H
                        openingWidth = width;
                        openingHeight = depth;
                        openingDepth = height;
                    }
                    // For floors, no swap needed - use bounding box values directly
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
            if (doc == null || elementId == null || !elementId.IsValid())
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
                Element? element = doc.GetElement(elementId);
                if (element != null && element.IsValidObject)
                {
                    _mepElementCache[elementId] = element;
                    return element;
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
        private void ApplyRotation(Document doc, FamilyInstance inst, XYZ placementPoint, double rotationAngle)
        {
            try
            {
                // Rotate around Z-axis (vertical) at the placement point
                XYZ axisOrigin = placementPoint;
                XYZ axisDirection = XYZ.BasisZ;
                Line rotationAxis = Line.CreateBound(axisOrigin, axisOrigin + axisDirection);
                ElementTransformUtils.RotateElement(doc, inst.Id, rotationAxis, rotationAngle);
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("placement_errors.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [ClusterPlacementService] ❌ Error applying rotation: {ex.Message}\n");
            }
        }

        /// <summary>
        /// Check if rotation angle is axis-aligned (0°, 90°, 180°, 270°).
        /// </summary>
        private bool IsAxisAlignedAngle(double angleRad)
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

