using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Services.Logging;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ClashZoneManagement
{
    /// <summary>
    /// Team G: SOLID-compliant damper auto-detection service.
    /// 
    /// This service automatically detects dampers that user forgot to select,
    /// preventing incorrect sleeve placement when ducts have dampers at their ends.
    /// SOLID: Single Responsibility - auto-detection only.
    /// 
    /// ✅ PRESERVES FOOLPROOF LOGIC:
    /// - Detects if user selected ducts but forgot Duct Accessories
    /// - Finds dampers intersecting same walls as ducts
    /// - Adds auto-detected dampers to intersection list
    /// </summary>
    public class DamperAutoDetectionService : IDamperAutoDetectionService
    {
        private readonly IDuctDamperFilterService _damperFilterService;
        private readonly ILogger _logger;
        
        /// <summary>
        /// Creates a new damper auto-detection service.
        /// </summary>
        /// <param name="damperFilterService">Duct-damper filter service (for IsDamperElement check)</param>
        /// <param name="logger">Optional logger for tracking operations</param>
        public DamperAutoDetectionService(IDuctDamperFilterService damperFilterService, ILogger logger = null)
        {
            _damperFilterService = damperFilterService ?? throw new ArgumentNullException(nameof(damperFilterService));
            _logger = logger ?? LoggerAdapter.Default;
        }
        
        /// <summary>
        /// Auto-detects missing dampers when user selected ducts but forgot Duct Accessories.
        /// </summary>
        public List<(Element, Element, BoundingBoxXYZ, XYZ)> AutoDetectMissingDampers(
            Document document,
            List<(Element, Element, BoundingBoxXYZ, XYZ)> currentIntersections)
        {
            var enhancedIntersections = new List<(Element, Element, BoundingBoxXYZ, XYZ)>(currentIntersections);
            
            try
            {
                // ✅ Check if user selected ducts but forgot duct accessories
                var hasDucts = currentIntersections.Any(i => 
                {
                    var mepCat = GetElementCategoryName(i.Item1);
                    return string.Equals(mepCat, "Ducts", StringComparison.OrdinalIgnoreCase);
                });
                
                var hasDuctAccessories = currentIntersections.Any(i => _damperFilterService.IsDamperElement(i.Item1));
                
                if (hasDucts && !hasDuctAccessories)
                {
                    _logger.Info("User selected ducts but forgot Duct Accessories - auto-detecting dampers", "DamperAutoDetection");
                    
                    // Get all structural elements (walls) that ducts intersect with
                    var ductWallIds = currentIntersections
                        .Where(i => string.Equals(GetElementCategoryName(i.Item1), "Ducts", StringComparison.OrdinalIgnoreCase))
                        .Select(i => i.Item2.Id)
                        .Distinct()
                        .ToList();
                    
                    // Find dampers that intersect with the same walls
                    var autoDetectedDampers = FindDampersIntersectingSameWalls(document, ductWallIds);
                    
                    foreach (var (damperElement, wallElement, boundingBox, intersectionPoint) in autoDetectedDampers)
                    {
                        // Avoid duplicates
                        if (!enhancedIntersections.Any(i => i.Item1.Id == damperElement.Id))
                        {
                            enhancedIntersections.Add((damperElement, wallElement, boundingBox, intersectionPoint));
                            _logger.Debug($"Auto-detected damper {damperElement.Id} intersecting wall {wallElement.Id}", "DamperAutoDetection");
                        }
                    }
                    
                    _logger.Info($"Auto-detected {autoDetectedDampers.Count} dampers that user missed", "DamperAutoDetection");
                }
                else if (hasDuctAccessories)
                {
                    _logger.Debug("User correctly selected Duct Accessories - no auto-detection needed", "DamperAutoDetection");
                }
                else
                {
                    _logger.Debug("No ducts selected - no auto-detection needed", "DamperAutoDetection");
                }
            }
            catch (Exception ex)
            {
                _logger.Warning($"Error in auto-detection: {ex.Message}", "DamperAutoDetection");
                // Fallback to original intersections
            }
            
            return enhancedIntersections;
        }
        
        /// <summary>
        /// Find dampers that intersect with the same walls as ducts.
        /// </summary>
        private List<(Element, Element, BoundingBoxXYZ, XYZ)> FindDampersIntersectingSameWalls(
            Document document,
            List<ElementId> wallIds)
        {
            var damperIntersections = new List<(Element, Element, BoundingBoxXYZ, XYZ)>();
            
            try
            {
                // ✅ Get all duct accessories (dampers) from the document
                var ductAccessories = new FilteredElementCollector(document)
                    .OfCategory(BuiltInCategory.OST_DuctAccessory)
                    .WhereElementIsNotElementType()
                    .Cast<Element>()
                    .ToList();
                
                // Also check family instances with "damper" in the name
                var damperFamilyInstances = new FilteredElementCollector(document)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol?.Family?.Name?.ToLowerInvariant().Contains("damper") == true)
                    .Cast<Element>()
                    .ToList();
                
                var allDampers = ductAccessories.Concat(damperFamilyInstances).ToList();
                
                foreach (var damper in allDampers)
                {
                    var damperBbox = damper.get_BoundingBox(null);
                    if (damperBbox == null) continue;
                    
                    // Check if damper intersects with any of the walls that ducts intersect
                    foreach (var wallId in wallIds)
                    {
                        var wallElement = document.GetElement(wallId);
                        if (wallElement == null) continue;
                        
                        var wallBbox = wallElement.get_BoundingBox(null);
                        if (wallBbox == null) continue;
                        
                        // ✅ Check if bounding boxes intersect using BoundingBoxService
                        if (BoundingBoxService.BoundingBoxesIntersect(damperBbox, wallBbox))
                        {
                            // Calculate intersection point (center of intersection)
                            var intersectionMin = new XYZ(
                                Math.Max(damperBbox.Min.X, wallBbox.Min.X),
                                Math.Max(damperBbox.Min.Y, wallBbox.Min.Y),
                                Math.Max(damperBbox.Min.Z, wallBbox.Min.Z)
                            );
                            
                            var intersectionMax = new XYZ(
                                Math.Min(damperBbox.Max.X, wallBbox.Max.X),
                                Math.Min(damperBbox.Max.Y, wallBbox.Max.Y),
                                Math.Min(damperBbox.Max.Z, wallBbox.Max.Z)
                            );
                            
                            var intersectionPoint = new XYZ(
                                (intersectionMin.X + intersectionMax.X) / 2,
                                (intersectionMin.Y + intersectionMax.Y) / 2,
                                (intersectionMin.Z + intersectionMax.Z) / 2
                            );
                            
                            var intersectionBbox = new BoundingBoxXYZ
                            {
                                Min = intersectionMin,
                                Max = intersectionMax
                            };
                            
                            damperIntersections.Add((damper, wallElement, intersectionBbox, intersectionPoint));
                            break; // Found intersection with this wall, move to next damper
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Warning($"Error finding dampers intersecting same walls: {ex.Message}", "DamperAutoDetection");
            }
            
            return damperIntersections;
        }
        
        /// <summary>
        /// Get MEP element category name.
        /// </summary>
        private string GetElementCategoryName(Element element)
        {
            try
            {
                if (element is Autodesk.Revit.DB.Mechanical.Duct) return "Ducts";
                if (element is Autodesk.Revit.DB.Plumbing.Pipe) return "Pipes";
                if (element is Autodesk.Revit.DB.Electrical.CableTray) return "Cable Trays";
                
                var categoryName = element?.Category?.Name;
                if (!string.IsNullOrEmpty(categoryName))
                {
                    return categoryName;
                }
                
                if (element.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_DuctAccessory)
                    return "Duct Accessories";
                
                return "Unknown";
            }
            catch
            {
                return "Unknown";
            }
        }
    }
}

