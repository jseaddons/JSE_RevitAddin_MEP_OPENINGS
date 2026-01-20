using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services.Combined.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Combined.Spatial;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Combined
{
    /// <summary>
    /// Service for detecting cross-category proximity groups.
    /// Uses spatial indexing for efficient O(n log n) proximity detection.
    /// </summary>
    public class CrossCategoryProximityService : ICrossCategoryProximityService
    {
        private readonly Action<string> _logger;
        
        public CrossCategoryProximityService()
        {
            _logger = msg => DebugLogger.Info(msg);
        }
        
        /// <summary>
        /// Detects proximity groups from a list of unified sleeves.
        /// Algorithm:
        /// 1. Build spatial index for fast proximity queries
        /// 2. For each sleeve, find nearby sleeves (different category)
        /// 3. Group mutually proximate sleeves using union-find
        /// 4. Return groups with 2+ sleeves from different categories
        /// </summary>
        public List<ProximityGroup> DetectProximityGroups(
            List<UnifiedSleeve> sleeves,
            double proximityThreshold)
        {
            if (sleeves == null || sleeves.Count == 0)
                return new List<ProximityGroup>();
            
            _logger($"[CrossCategoryProximity] Starting detection for {sleeves.Count} sleeves, threshold={proximityThreshold:F2} ft");
            
            // 1. Build spatial index
            var spatialIndex = BuildSpatialIndex(sleeves);
            
            // 2. Find proximity relationships
            var proximityPairs = FindProximityPairs(sleeves, spatialIndex, proximityThreshold);
            
            _logger($"[CrossCategoryProximity] Found {proximityPairs.Count} proximity pairs");
            
            // 3. Group sleeves using union-find
            var groups = GroupSleeves(sleeves, proximityPairs);
            
            // 4. Filter to cross-category groups only
            var crossCategoryGroups = groups
                .Where(g => g.IsCrossCategory() && g.IsValid())
                .ToList();
            
            _logger($"[CrossCategoryProximity] Created {crossCategoryGroups.Count} cross-category proximity groups");
            
            foreach (var group in crossCategoryGroups)
            {
                _logger($"[CrossCategoryProximity]   - {group.GetSummary()}");
            }
            
            return crossCategoryGroups;
        }
        
        /// <summary>
        /// Builds a spatial index from the list of sleeves
        /// </summary>
        private SimplifiedSpatialIndex BuildSpatialIndex(List<UnifiedSleeve> sleeves)
        {
            var index = new SimplifiedSpatialIndex();
            
            foreach (var sleeve in sleeves)
            {
                if (sleeve.BoundingBox != null)
                {
                    _logger($"[SpatialIndex] Inserting {sleeve.Id} ({sleeve.Category}): Min={sleeve.BoundingBox.Min}, Max={sleeve.BoundingBox.Max}");
                    index.Insert(sleeve.Id, sleeve.BoundingBox.Min, sleeve.BoundingBox.Max);
                }
                else
                {
                    _logger($"[SpatialIndex] SKIPPING {sleeve.Id}: BoundingBox is null!");
                }
            }
            
            index.Build();
            _logger($"[SpatialIndex] Build complete.");
            
            return index;
        }
        
        /// <summary>
        /// Finds all pairs of sleeves that are within proximity threshold
        /// and from different categories
        /// </summary>
        private List<(string id1, string id2)> FindProximityPairs(
            List<UnifiedSleeve> sleeves,
            SimplifiedSpatialIndex spatialIndex,
            double proximityThreshold)
        {
            var pairs = new List<(string, string)>();
            var processed = new HashSet<string>();
            
            foreach (var sleeve in sleeves)
            {
                // Expand bounding box by threshold for search
                var searchMin = sleeve.BoundingBox.Min - new XYZ(proximityThreshold, proximityThreshold, proximityThreshold);
                var searchMax = sleeve.BoundingBox.Max + new XYZ(proximityThreshold, proximityThreshold, proximityThreshold);
                
                // Query spatial index for nearby sleeves
                var nearbyIds = spatialIndex.Query(searchMin, searchMax);
                
                // DIAGNOSTIC LOGGING
                if (nearbyIds.Count > 0)
                {
                   // _logger($"[ProximityDebug] Sleeve {sleeve.Id} ({sleeve.Category}) found {nearbyIds.Count} potential neighbors in spatial index.");
                }

                foreach (var nearbyId in nearbyIds)
                {
                    // Skip self
                    if (nearbyId == sleeve.Id)
                        continue;
                    
                    // Skip if already processed this pair
                    var pairKey = GetPairKey(sleeve.Id, nearbyId);
                    if (processed.Contains(pairKey))
                        continue;
                    
                    // Get the nearby sleeve
                    var nearbySleeve = sleeves.FirstOrDefault(s => s.Id == nearbyId);
                    if (nearbySleeve == null)
                        continue;
                    
                    // Only consider cross-category pairs
                    if (sleeve.Category == nearbySleeve.Category)
                        continue;
                    
                    // Check proximity using bounding box edges (accounts for element size)
                    bool intersects = sleeve.BoundingBoxIntersects(nearbySleeve, proximityThreshold);
                    
                    
                    _logger($"[ProximityDetail] Checking {sleeve.Id} vs {nearbyId}:");
                    _logger($"  - Sleeve 1 BBox: {sleeve.BoundingBox.Min} to {sleeve.BoundingBox.Max}");
                    _logger($"  - Sleeve 2 BBox: {nearbySleeve.BoundingBox.Min} to {nearbySleeve.BoundingBox.Max}");
                    _logger($"  - Distance Result: {intersects}");
                    

                    if (intersects)
                    {
                        pairs.Add((sleeve.Id, nearbyId));
                        processed.Add(pairKey);
                        _logger($"[ProximityMatch] FOUND PAIR: {sleeve.Id} ({sleeve.Category}) <-> {nearbySleeve.Id} ({nearbySleeve.Category})");
                    }
                }
            }
            
            return pairs;
        }
        
        /// <summary>
        /// Groups sleeves based on proximity pairs using union-find algorithm
        /// </summary>
        private List<ProximityGroup> GroupSleeves(
            List<UnifiedSleeve> sleeves,
            List<(string id1, string id2)> proximityPairs)
        {
            // Union-Find data structure
            var parent = new Dictionary<string, string>();
            
            // Initialize: each sleeve is its own parent
            foreach (var sleeve in sleeves)
            {
                parent[sleeve.Id] = sleeve.Id;
            }
            
            // Union operation for each proximity pair
            foreach (var (id1, id2) in proximityPairs)
            {
                var root1 = Find(parent, id1);
                var root2 = Find(parent, id2);
                
                if (root1 != root2)
                {
                    parent[root2] = root1; // Union
                }
            }
            
            // Group sleeves by their root parent
            var groupMap = new Dictionary<string, ProximityGroup>();
            
            foreach (var sleeve in sleeves)
            {
                var root = Find(parent, sleeve.Id);
                
                if (!groupMap.ContainsKey(root))
                {
                    groupMap[root] = new ProximityGroup();
                }
                
                groupMap[root].Add(sleeve);
            }
            
            // Return groups with 2+ sleeves
            return groupMap.Values
                .Where(g => g.Count >= 2)
                .ToList();
        }
        
        /// <summary>
        /// Union-Find: Find operation with path compression
        /// </summary>
        private string Find(Dictionary<string, string> parent, string id)
        {
            if (parent[id] != id)
            {
                parent[id] = Find(parent, parent[id]); // Path compression
            }
            return parent[id];
        }
        
        /// <summary>
        /// Gets a unique key for a pair of sleeve IDs (order-independent)
        /// </summary>
        private string GetPairKey(string id1, string id2)
        {
            // Sort to ensure order-independence
            return string.CompareOrdinal(id1, id2) < 0
                ? $"{id1}|{id2}"
                : $"{id2}|{id1}";
        }
    }
}
