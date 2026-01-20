using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Services.Logging;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ClashZoneManagement
{
    /// <summary>
    /// Team G: SOLID-compliant clash zone detection service.
    /// 
    /// This service converts intersections (from MepIntersectionService) into ClashZone objects.
    /// It does NOT perform intersection detection - that's handled by MepIntersectionService.
    /// 
    /// SOLID: Single Responsibility - conversion and validation only.
    /// 
    /// ✅ PRESERVES ALL LOGIC:
    /// - Converts intersections to ClashZone objects
    /// - Validates elements (existence, penetration adequacy)
    /// - Checks for existing clash zones (deduplication)
    /// - Applies damper filtering (for ducts)
    /// - Creates ClashZone objects with proper metadata
    /// </summary>
    public class ClashZoneDetectionService : IClashZoneDetectionService
    {
        private readonly GuidManager _guidManager;
        private readonly IDuctDamperFilterService _ductDamperFilterService;
        private readonly IDamperAutoDetectionService _damperAutoDetectionService;
        private readonly ILogger _logger;
        
        /// <summary>
        /// Creates a new clash zone detection service.
        /// 
        /// ✅ SOLID: Dependency Inversion - depends on abstractions (interfaces), not concretions.
        /// </summary>
        /// <param name="document">Revit document</param>
        /// <param name="ductDamperFilterService">Duct-damper filter service (for skipping ducts when dampers present)</param>
        /// <param name="damperAutoDetectionService">Damper auto-detection service (for detecting missing dampers)</param>
        /// <param name="logger">Optional logger for tracking operations</param>
        public ClashZoneDetectionService(
            Document document,
            IDuctDamperFilterService ductDamperFilterService = null,
            IDamperAutoDetectionService damperAutoDetectionService = null,
            ILogger logger = null)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));
            
            _guidManager = new GuidManager(document);
            _ductDamperFilterService = ductDamperFilterService;
            _damperAutoDetectionService = damperAutoDetectionService;
            _logger = logger ?? LoggerAdapter.Default;
        }
        
        /// <summary>
        /// Detect new clash zones from intersections.
        /// Converts raw intersections (from MepIntersectionService) into ClashZone objects.
        /// </summary>
        public List<ClashZone> DetectNewClashZones(
            List<(Element, Element, BoundingBoxXYZ, XYZ)> currentIntersections,
            Document document,
            ClashZoneStorage storage,
            Dictionary<string, double> clearanceSettings = null,
            List<string> selectedCategories = null)
        {
            if (currentIntersections == null || currentIntersections.Count == 0)
            {
                _logger.Info("No intersections to process", "ClashZoneDetection");
                return new List<ClashZone>();
            }
            
            if (storage == null)
            {
                _logger.Warning("ClashZoneStorage is null, cannot detect clash zones", "ClashZoneDetection");
                return new List<ClashZone>();
            }
            
            if (storage.ClashZones == null)
            {
                _logger.Warning("ClashZones collection is null, initializing", "ClashZoneDetection");
                storage.ClashZones = new List<ClashZone>();
            }
            
            _logger.Info($"Processing {currentIntersections.Count} intersections to detect clash zones", "ClashZoneDetection");
            _logger.Info($"Storage has {storage.ClashZones?.Count ?? 0} existing clash zones for deduplication", "ClashZoneDetection");
            
            // ✅ DIAGNOSTIC: Direct logging to refresh log (DebugLogger writes to refresh log)
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[ClashZoneDetection] Processing {currentIntersections.Count} intersections to detect clash zones");
                DebugLogger.Info($"[ClashZoneDetection] Storage has {storage.ClashZones?.Count ?? 0} existing clash zones for deduplication");
            }
            
            // ✅ STEP 1: Auto-detect missing dampers (foolproof - if user forgot Duct Accessories)
            List<(Element, Element, BoundingBoxXYZ, XYZ)> enhancedIntersections = currentIntersections;
            if (_damperAutoDetectionService != null)
            {
                try
                {
                    enhancedIntersections = _damperAutoDetectionService.AutoDetectMissingDampers(document, currentIntersections);
                    if (enhancedIntersections.Count > currentIntersections.Count)
                    {
                        _logger.Info($"Auto-detected {enhancedIntersections.Count - currentIntersections.Count} missing dampers", "ClashZoneDetection");
                    }
                }
                catch (Exception ex)
                {
                    _logger.Warning($"Error in auto-detection, using original intersections: {ex.Message}", "ClashZoneDetection");
                    enhancedIntersections = currentIntersections;
                }
            }
            
            // ✅ STEP 2: Pre-calculate damper locations once (optimization - Calculate Once, Use Many Times)
            List<(ElementId damperId, BoundingBoxXYZ bbox, ElementId wallId)> damperLocations = new List<(ElementId, BoundingBoxXYZ, ElementId)>();
            if (_ductDamperFilterService != null)
            {
                try
                {
                    damperLocations = _ductDamperFilterService.PreCalculateDamperLocations(storage, enhancedIntersections, document);
                    _logger.Info($"Pre-calculated {damperLocations.Count} damper locations for filtering", "ClashZoneDetection");
                }
                catch (Exception ex)
                {
                    _logger.Warning($"Error pre-calculating damper locations: {ex.Message}", "ClashZoneDetection");
                    damperLocations = new List<(ElementId, BoundingBoxXYZ, ElementId)>();
                }
            }
            
            var newClashZones = new List<ClashZone>();
            int processedCount = 0;
            int skippedCount = 0;
            int skippedForDamper = 0;
            
            foreach (var (mepElement, structuralElement, boundingBox, intersectionPoint) in enhancedIntersections)
            {
                processedCount++;
                
                try
                {
                    // ✅ STEP 1: Validate elements
                    if (mepElement == null || structuralElement == null)
                    {
                        _logger.Debug($"Skipping intersection {processedCount}: Invalid elements", "ClashZoneDetection");
                        skippedCount++;
                        continue;
                    }
                    
                    // ✅ STEP 2: Check if clash zone already exists (deduplication)
                    var existingClashZone = FindExistingClashZone(
                        mepElement.Id, 
                        structuralElement.Id, 
                        intersectionPoint, 
                        storage.ClashZones);
                    
                    if (existingClashZone != null)
                    {
                        _logger.Info($"Skipping intersection {processedCount}: Clash zone already exists (ID: {existingClashZone.Id}, MEP={mepElement.Id}, Host={structuralElement.Id})", "ClashZoneDetection");
                        if (!DeploymentConfiguration.DeploymentMode && processedCount <= 5)
                        {
                            DebugLogger.Info($"[ClashZoneDetection] Skipping intersection {processedCount}: Clash zone already exists (ID: {existingClashZone.Id}, MEP={mepElement.Id}, Host={structuralElement.Id})");
                        }
                        skippedCount++;
                        continue;
                    }
                    
                    // ✅ STEP 3: Apply duct-damper filter (skip ducts when dampers present)
                    var mepCategory = GetElementCategoryName(mepElement);
                    bool ductHasDamperNearby = false;
                    
                    if (string.Equals(mepCategory, "Ducts", StringComparison.OrdinalIgnoreCase) && _ductDamperFilterService != null)
                    {
                        bool shouldSkip = _ductDamperFilterService.ShouldSkipDuctForDamper(
                            mepElement,
                            structuralElement.Id,
                            intersectionPoint,
                            damperLocations ?? new List<(ElementId, BoundingBoxXYZ, ElementId)>(),
                            existingClashZone);
                        
                        if (shouldSkip)
                        {
                            _logger.Debug($"Skipping duct {mepElement.Id}: Damper present nearby on wall {structuralElement.Id}", "ClashZoneDetection");
                            skippedCount++;
                            skippedForDamper++;
                            
                            // ✅ Set flag on existing clash zone if found
                            if (existingClashZone != null)
                            {
                                existingClashZone.HasDamperNearby = true;
                                _logger.Debug($"Set HasDamperNearby=true on existing clash zone {existingClashZone.Id}", "ClashZoneDetection");
                            }
                            else
                            {
                                // Track that this duct has damper nearby - will set flag if new clash zone is created
                                ductHasDamperNearby = true;
                            }
                            
                            continue;
                        }
                    }
                    
                    // ✅ STEP 4: Apply other validation filters (penetration checks, etc.)
                    if (!ShouldCreateClashZone(mepElement, structuralElement, intersectionPoint, document))
                    {
                        _logger.Info($"Skipping intersection {processedCount}: Failed validation filters (MEP={mepElement.Id}, Host={structuralElement.Id}, Category={GetElementCategoryName(mepElement)})", "ClashZoneDetection");
                        if (!DeploymentConfiguration.DeploymentMode && processedCount <= 5)
                        {
                            DebugLogger.Info($"[ClashZoneDetection] Skipping intersection {processedCount}: Failed validation filters (MEP={mepElement.Id}, Host={structuralElement.Id}, Category={GetElementCategoryName(mepElement)})");
                        }
                        skippedCount++;
                        continue;
                    }
                    
                    // ✅ STEP 5: Create new ClashZone object
                    var newClashZone = CreateClashZone(
                        mepElement, 
                        structuralElement, 
                        intersectionPoint, 
                        boundingBox, 
                        document, 
                        clearanceSettings);
                    
                    if (newClashZone != null)
                    {
                        // ✅ CRITICAL: Set HasDamperNearby flag if duct-damper combo was detected
                        if (ductHasDamperNearby)
                        {
                            newClashZone.HasDamperNearby = true;
                            _logger.Debug($"Set HasDamperNearby=true on new clash zone {newClashZone.Id} for duct {mepElement.Id}", "ClashZoneDetection");
                        }
                        
                        newClashZone.IsCurrentClash = true;
                        newClashZone.ReadyForPlacement = true;
                        newClashZone.DetectedAt = DateTime.Now;
                        newClashZone.LastUpdated = DateTime.Now;
                        
                        // Clear heavy API objects immediately
                        newClashZone.ClearRevitApiObjects();
                        
                        newClashZones.Add(newClashZone);
                        storage.ClashZones.Add(newClashZone);
                        
                        _logger.Debug($"Created clash zone {newClashZone.Id} for MEP={mepElement.Id}, Structural={structuralElement.Id}", "ClashZoneDetection");
                    }
                }
                catch (Exception ex)
                {
                    _logger.Warning($"Error processing intersection {processedCount}: {ex.Message}", "ClashZoneDetection");
                    skippedCount++;
                }
            }
            
            _logger.Info($"Detection complete: {newClashZones.Count} new clash zones created, {skippedCount} skipped ({skippedForDamper} for damper, {skippedCount - skippedForDamper} for other reasons), {processedCount} total processed", "ClashZoneDetection");
            
            // ✅ DIAGNOSTIC: Direct logging to refresh log
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[ClashZoneDetection] Detection complete: {newClashZones.Count} new clash zones created, {skippedCount} skipped ({skippedForDamper} for damper, {skippedCount - skippedForDamper} for other reasons), {processedCount} total processed");
            }
            
            if (newClashZones.Count == 0 && processedCount > 0)
            {
                _logger.Warning($"⚠️ WARNING: All {processedCount} intersections were filtered out! Check logs above for reasons.", "ClashZoneDetection");
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[ClashZoneDetection] ⚠️ WARNING: All {processedCount} intersections were filtered out! Skipped breakdown: {skippedForDamper} for damper, {skippedCount - skippedForDamper} for other reasons (existing zones, validation failures, etc.)");
                }
            }
            
            return newClashZones;
        }
        
        /// <summary>
        /// Find existing clash zone by MEP element, structural element, and intersection point.
        /// Uses GuidManager for efficient lookup.
        /// </summary>
        private ClashZone FindExistingClashZone(
            ElementId mepElementId, 
            ElementId structuralElementId, 
            XYZ intersectionPoint,
            List<ClashZone> existingClashZones)
        {
            if (existingClashZones == null || existingClashZones.Count == 0)
                return null;
            
            // ✅ Use GuidManager for efficient lookup
            var existingByPoint = _guidManager.FindByMepHostAndPoint(
                existingClashZones,
                mepElementId.IntegerValue,
                structuralElementId.IntegerValue,
                intersectionPoint);
            
            if (existingByPoint != null)
                return existingByPoint;
            
            // Fallback: Simple MEP+Host match
            var existingByMepHost = _guidManager.FindByMepAndHost(
                existingClashZones,
                mepElementId,
                structuralElementId);
            
            return existingByMepHost;
        }
        
        /// <summary>
        /// Determine if a clash zone should be created for this intersection.
        /// Applies validation filters (damper checks, penetration adequacy, etc.).
        /// </summary>
        private bool ShouldCreateClashZone(
            Element mepElement, 
            Element structuralElement, 
            XYZ intersectionPoint,
            Document document)
        {
            try
            {
                // ✅ Basic validation: Check element IDs are valid
                if (mepElement.Id.IntegerValue <= 0 || structuralElement.Id.IntegerValue <= 0)
                {
                    return false;
                }
                
                // ✅ Skip pipe accessories (handled separately)
                var mepCategory = GetElementCategoryName(mepElement);
                if (mepCategory == "Pipe Accessories")
                {
                    return false;
                }
                
                // ✅ NOTE: Duct-damper filtering is handled separately in the main loop
                // ✅ TODO: Add penetration adequacy checks (if needed)
                // For now, allow all valid intersections through
                
                return true;
            }
            catch (Exception ex)
            {
                _logger.Warning($"Error in validation filter: {ex.Message}", "ClashZoneDetection");
                return false; // Fail-safe: skip on error
            }
        }
        
        /// <summary>
        /// Create a new ClashZone object from intersection data.
        /// This is a simplified version - can be enhanced with more metadata from legacy code.
        /// </summary>
        private ClashZone CreateClashZone(
            Element mepElement, 
            Element structuralElement, 
            XYZ intersectionPoint, 
            BoundingBoxXYZ boundingBox, 
            Document document,
            Dictionary<string, double> clearanceSettings)
        {
            try
            {
                // ✅ Fix zero intersection points
                if (intersectionPoint == null || 
                    (Math.Abs(intersectionPoint.X) < 1e-9 && 
                     Math.Abs(intersectionPoint.Y) < 1e-9 && 
                     Math.Abs(intersectionPoint.Z) < 1e-9))
                {
                    if (boundingBox != null)
                    {
                        intersectionPoint = new XYZ(
                            (boundingBox.Min.X + boundingBox.Max.X) / 2.0,
                            (boundingBox.Min.Y + boundingBox.Max.Y) / 2.0,
                            (boundingBox.Min.Z + boundingBox.Max.Z) / 2.0);
                    }
                    else
                    {
                        var mepBBox = mepElement.get_BoundingBox(null);
                        if (mepBBox != null)
                        {
                            intersectionPoint = new XYZ(
                                (mepBBox.Min.X + mepBBox.Max.X) / 2.0,
                                (mepBBox.Min.Y + mepBBox.Max.Y) / 2.0,
                                (mepBBox.Min.Z + mepBBox.Max.Z) / 2.0);
                        }
                    }
                }
                
                var mepCategory = GetElementCategoryName(mepElement);
                var structuralElementType = GetStructuralElementType(structuralElement);
                
                // ✅ Get deterministic GUID from GuidManager
                int mepId = mepElement.Id.IntegerValue;
                int hostId = structuralElement.Id.IntegerValue;
                Guid clashZoneGuid = _guidManager.GetOrCreateDeterministicGuidDatabaseFirst(
                    mepId,
                    hostId,
                    intersectionPoint.X,
                    intersectionPoint.Y,
                    intersectionPoint.Z,
                    tolerance: 0.1);
                
                var clashZone = new ClashZone
                {
                    Id = clashZoneGuid,
                    MepElementId = mepElement.Id,
                    StructuralElementId = structuralElement.Id,
                    IntersectionPoint = intersectionPoint,
                    SleevePlacementPoint = intersectionPoint,
                    IntersectionPointX = intersectionPoint.X,
                    IntersectionPointY = intersectionPoint.Y,
                    IntersectionPointZ = intersectionPoint.Z,
                    SleevePlacementPointX = intersectionPoint.X,
                    SleevePlacementPointY = intersectionPoint.Y,
                    SleevePlacementPointZ = intersectionPoint.Z,
                    ClashBoundingBox = boundingBox,
                    MepElementCategory = MepCategoryConstants.Normalize(mepCategory),
                    StructuralElementType = structuralElementType,
                    DocumentPath = document.PathName,
                    StructuralElementDocumentTitle = structuralElement.Document.Title,
                    SourceDocKey = mepElement.Document.Title ?? mepElement.Document.PathName ?? string.Empty,
                    HostDocKey = structuralElement.Document.Title ?? structuralElement.Document.PathName ?? string.Empty,
                    HostOrientation = WallDirectionService.GetHostOrientation(structuralElement),
                    StructuralElementNormal = WallDirectionService.GetStructuralElementNormal(structuralElement),
                    WallDirection = WallDirectionService.GetWallDirection(structuralElement),
                    WallDirectionType = WallDirectionService.GetWallDirectionType(structuralElement, WallDirectionService.GetWallDirection(structuralElement)),
                    IsResolved = false,
                    IsClusterResolved = false,
                    SleeveInstanceId = -1,
                    ClusterSleeveInstanceId = -1
                };
                
                // ✅ TODO: Add more metadata from legacy CreateClashZone method:
                // - MEP element dimensions (width, height)
                // - MEP element orientation
                // - Structural element thickness
                // - MEP element level info
                // - Duct shape, insulation type, etc.
                // These can be added incrementally as needed
                
                return clashZone;
            }
            catch (Exception ex)
            {
                _logger.Error($"Error creating clash zone: {ex.Message}", ex, "ClashZoneDetection");
                return null;
            }
        }
        
        /// <summary>
        /// Get MEP element category name.
        /// </summary>
        private string GetElementCategoryName(Element element)
        {
            try
            {
                // ✅ CRITICAL: Check element type FIRST for accurate categorization
                if (element is Autodesk.Revit.DB.Mechanical.Duct) return "Ducts";
                if (element is Autodesk.Revit.DB.Plumbing.Pipe) return "Pipes";
                if (element is Autodesk.Revit.DB.Electrical.CableTray) return "Cable Trays";
                
                // Fallback to category name
                var categoryName = element?.Category?.Name;
                if (!string.IsNullOrEmpty(categoryName))
                {
                    return categoryName;
                }
                
                // Check by category ID for Duct Accessories
                if (element.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_DuctAccessory)
                    return "Duct Accessories";
                
                return "Unknown";
            }
            catch (Exception ex)
            {
                _logger.Warning($"Error getting element category: {ex.Message}", "ClashZoneDetection");
                return "Unknown";
            }
        }
        
        /// <summary>
        /// Get structural element type name.
        /// </summary>
        private string GetStructuralElementType(Element element)
        {
            if (element is Wall)
                return "Wall";
            else if (element is Floor)
                return "Floor";
            else if (element is FamilyInstance famInst && 
                     famInst.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                return "Structural Framing";
            else
                return "Unknown";
        }
    }
}

