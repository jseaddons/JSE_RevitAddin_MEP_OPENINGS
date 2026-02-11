using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.MultiFloor
{
    /// <summary>
    /// Detects clashes for a single floor in isolation
    /// Uses spatial indexing for O(N) performance
    /// </summary>
    public class FloorIsolatedClashDetector
    {
        private readonly Document _doc;
        private readonly SharedResourceCache _cache;
        
        public FloorIsolatedClashDetector(Document _doc, SharedResourceCache _cache)
        {
            this._doc = _doc;
            this._cache = _cache;
        }
        
        /// <summary>
        /// Detect clashes for a single floor (thread-safe)
        /// </summary>
        public List<ClashZone> DetectClashesForFloor(
            Level level,
            List<string> categories)
        {
            // Get MEP elements ONLY for this floor
            var elements = GetMEPElementsForFloor(level, categories);
            
            if (elements.Count == 0)
                return new List<ClashZone>();
            
            // Get structural hosts ONLY for this floor
            var hosts = GetStructuralHostsForFloor(level);
            
            // Build spatial grid (floor-local, thread-safe)
            var grid = new FloorSpatialGrid(elements, hosts, cellSize: 5.0);
            
            // Detect clashes using spatial indexing
            var clashes = new List<ClashZone>();
            foreach (var element in elements)
            {
                var candidates = grid.GetCandidates(element);
                clashes.AddRange(DetectClashesForElement(element, candidates, level));
            }
            
            return clashes;
        }
        
        /// <summary>
        /// Get MEP elements that belong to this floor only
        /// </summary>
        private List<Element> GetMEPElementsForFloor(Level level, List<string> categories)
        {
            var levelZ = level.Elevation;
            var nextLevel = GetNextLevel(level);
            var levelHeight = nextLevel != null 
                ? (nextLevel.Elevation - levelZ) 
                : 10.0; // Default 10 feet
            
            var elements = new List<Element>();
            
            foreach (var categoryName in categories)
            {
                var builtInCat = GetBuiltInCategory(categoryName);
                if (builtInCat == BuiltInCategory.INVALID) continue;
                
                var collector = new FilteredElementCollector(_doc)
                    .OfCategory(builtInCat)
                    .WhereElementIsNotElementType();
                
                foreach (var elem in collector)
                {
                    if (IsElementOnLevel(elem, levelZ, levelHeight))
                    {
                        elements.Add(elem);
                    }
                }
            }
            
            return elements;
        }
        
        /// <summary>
        /// Get structural hosts (walls, floors) for this floor
        /// </summary>
        private List<Element> GetStructuralHostsForFloor(Level level)
        {
            var hosts = new List<Element>();
            
            // Get walls on this level
            if (_cache.Walls != null)
            {
                var levelWalls = _cache.Walls.Values
                    .Where(w => w.LevelId == level.Id)
                    .Cast<Element>()
                    .ToList();
                hosts.AddRange(levelWalls);
            }
            
            // Get floors (slabs)
            if (_cache.Floors != null)
            {
                var levelFloors = _cache.Floors.Values
                    .Where(f => f.LevelId == level.Id)
                    .Cast<Element>()
                    .ToList();
                hosts.AddRange(levelFloors);
            }
            
            return hosts;
        }
        
        /// <summary>
        /// Check if element is on this level by Z-coordinate
        /// </summary>
        private bool IsElementOnLevel(Element elem, double levelZ, double levelHeight)
        {
            var bbox = elem.get_BoundingBox(null);
            if (bbox == null) return false;
            
            // Element is on level if its bounding box overlaps level's Z range
            // We use a small epsilon or buffer to include elements sitting exactly on the level
            const double epsilon = 0.01;
            return bbox.Min.Z <= (levelZ + levelHeight - epsilon) && bbox.Max.Z >= levelZ - epsilon;
        }
        
        private Level? GetNextLevel(Level current)
        {
            return new FilteredElementCollector(_doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .Where(l => l.Elevation > current.Elevation)
                .OrderBy(l => l.Elevation)
                .FirstOrDefault();
        }
        
        private BuiltInCategory GetBuiltInCategory(string category)
        {
            return category.ToLower() switch
            {
                "ducts" => BuiltInCategory.OST_DuctCurves,
                "duct accessories" => BuiltInCategory.OST_DuctAccessory,
                "pipes" => BuiltInCategory.OST_PipeCurves,
                "cable trays" => BuiltInCategory.OST_CableTray,
                _ => BuiltInCategory.INVALID
            };
        }
        
        private List<ClashZone> DetectClashesForElement(Element element, List<Element> candidates, Level level)
        {
            var clashes = new List<ClashZone>();
            
            // Implementation of clash detection logic between the MEP element and candidates (hosts)
            // This would normally call into the existing IntersectionDetectionService or similar
            // For now, this is a placeholder to be integrated with the project's actual intersection logic.
            // TODO: Integrate with IntersectionOptimizationService or MepIntersectionService
            
            return clashes;
        }
    }
}
