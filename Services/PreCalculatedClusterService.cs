using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Pre-calculates cluster groupings from XML data using "calculate once, use many times" approach
    /// This eliminates the need for expensive Revit API bounding box calculations during clustering
    /// </summary>
    public class PreCalculatedClusterService
    {
        private readonly Dictionary<string, List<List<ClashZone>>> _preCalculatedClusters = new Dictionary<string, List<List<ClashZone>>>();
        private readonly Dictionary<string, double> _toleranceDistances = new Dictionary<string, double>();
        
        /// <summary>
        /// Pre-calculate cluster groupings for all categories from XML files
        /// This should be called once during initialization, before any clustering operations
        /// </summary>
        public void PreCalculateClustersFromXml(string xmlFilePath, double toleranceDistance)
        {
            try
            {
                DebugLogger.Info($"[PreCalculatedClusterService] Pre-calculating clusters from {xmlFilePath}");
                
                // Load only CURRENT clash zones from XML
                var clashZones = LoadClashZonesFromXml(xmlFilePath);
                PreCalculateClustersFromClashZones(clashZones, toleranceDistance);
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[PreCalculatedClusterService] Error pre-calculating clusters: {ex.Message}");
            }
        }

        public void PreCalculateClustersFromClashZones(List<ClashZone> clashZones, double toleranceDistance)
        {
            try
            {
                DebugLogger.Info($"[PreCalculatedClusterService] Pre-calculating clusters from {clashZones?.Count ?? 0} clash zones");
                
                if (clashZones == null || clashZones.Count == 0)
                {
                    DebugLogger.Warning($"[PreCalculatedClusterService] No current clash zones found");
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", 
                        $"No current clash zones found - skipping pre-calculation\n");
                    return;
                }
                
                DebugLogger.Info($"[PreCalculatedClusterService] Loaded {clashZones.Count} CURRENT clash zones from XML");
                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", 
                    $"===== PROCESSING {clashZones.Count} CURRENT CLASH ZONES =====\n");
                
                // Group clash zones by category and host type
                var groupedClashZones = GroupClashZonesByCategoryAndHost(clashZones);
                
                // Pre-calculate clusters for each group
                foreach (var group in groupedClashZones)
                {
                    var groupKey = group.Key;
                    var zones = group.Value;
                    
                    DebugLogger.Info($"[PreCalculatedClusterService] Processing group: {groupKey} with {zones.Count} zones");
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", 
                        $"GROUP: {groupKey} ({zones.Count} zones)\n");
                    
                    // Calculate clusters for this group
                    var clusters = CalculateClustersForGroup(zones, toleranceDistance, groupKey);
                    
                    // Store pre-calculated clusters
                    _preCalculatedClusters[groupKey] = clusters;
                    _toleranceDistances[groupKey] = toleranceDistance;
                    
                    DebugLogger.Info($"[PreCalculatedClusterService] Pre-calculated {clusters.Count} clusters for group {groupKey}");
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", 
                        $"RESULT: {clusters.Count} clusters formed from {zones.Count} zones\n");
                }
                
                DebugLogger.Info($"[PreCalculatedClusterService] Pre-calculation complete for {clashZones?.Count ?? 0} clash zones");
                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", 
                    $"===== PRE-CALCULATION COMPLETED =====\n\n");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[PreCalculatedClusterService] Error pre-calculating clusters: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Get pre-calculated clusters for a specific category
        /// Returns the clusters that were calculated during PreCalculateClustersFromXml
        /// </summary>
        public List<List<ClashZone>> GetPreCalculatedClusters(string category, string hostType, string orientation)
        {
            var groupKey = $"{category}_{hostType}_{orientation}";
            
            if (_preCalculatedClusters.TryGetValue(groupKey, out var clusters))
            {
                DebugLogger.Info($"[PreCalculatedClusterService] Retrieved {clusters.Count} pre-calculated clusters for {groupKey}");
                return clusters;
            }
            
            DebugLogger.Warning($"[PreCalculatedClusterService] No pre-calculated clusters found for {groupKey}");
            return new List<List<ClashZone>>();
        }
        
        /// <summary>
        /// Check if clusters are pre-calculated for a specific category
        /// </summary>
        public bool HasPreCalculatedClusters(string category, string hostType, string orientation)
        {
            var groupKey = $"{category}_{hostType}_{orientation}";
            return _preCalculatedClusters.ContainsKey(groupKey);
        }
        
        /// <summary>
        /// Load clash zones from XML file
        /// </summary>
        private List<ClashZone> LoadClashZonesFromXml(string xmlFilePath)
        {
            try
            {
                var serializer = new XmlSerializer(typeof(OpeningFilter));
                using (var reader = new StreamReader(xmlFilePath))
                {
                    var filter = (OpeningFilter)serializer.Deserialize(reader);
                    var allClashZones = filter?.ClashZoneStorage?.ClashZones ?? new List<ClashZone>();
                    
                    // ✅ CRITICAL FILTER: Only process CURRENT clashes (not old saved clashes)
                    var currentClashZones = allClashZones.Where(cz => cz.IsCurrentClash).ToList();
                    
                    // ✅ CRITICAL FIX: Reconstruct SleevePlacementPoint from XML-serializable properties
                    foreach (var clashZone in currentClashZones)
                    {
                        DebugLogger.Log($"[PreCalculatedClusterService] BEFORE reconstruction: Zone {clashZone.Id}, SleevePlacementPoint={clashZone.SleevePlacementPoint}, X={clashZone.SleevePlacementPointX}, Y={clashZone.SleevePlacementPointY}, Z={clashZone.SleevePlacementPointZ}");
                        clashZone.EnsureSleevePlacementPointReconstructed();
                        clashZone.EnsureSleevePlacementPointActiveDocumentReconstructed();
                        DebugLogger.Log($"[PreCalculatedClusterService] AFTER reconstruction: Zone {clashZone.Id}, SleevePlacementPoint={clashZone.SleevePlacementPoint}, SleevePlacementPointActiveDocument={clashZone.SleevePlacementPointActiveDocument}");
                    }
                    
                    DebugLogger.Log($"[PreCalculatedClusterService] Loaded {allClashZones.Count} total clash zones, filtering to {currentClashZones.Count} current clashes");
                    
                    return currentClashZones;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[PreCalculatedClusterService] Error loading XML: {ex.Message}");
                return new List<ClashZone>();
            }
        }
        
        /// <summary>
        /// Group clash zones by category, host type, and orientation
        /// </summary>
        private Dictionary<string, List<ClashZone>> GroupClashZonesByCategoryAndHost(List<ClashZone> clashZones)
        {
            var grouped = new Dictionary<string, List<ClashZone>>();
            
            foreach (var zone in clashZones)
            {
                var category = zone.MepElementCategory ?? "Unknown";
                var hostType = zone.StructuralElementType ?? "Unknown";
                var orientation = zone.HostOrientation ?? "";
                
                var groupKey = $"{category}_{hostType}_{orientation}";
                
                if (!grouped.ContainsKey(groupKey))
                {
                    grouped[groupKey] = new List<ClashZone>();
                }
                
                grouped[groupKey].Add(zone);
            }
            
            return grouped;
        }
        
        /// <summary>
        /// ⚠️ CRITICAL PROTECTION: Calculate clusters for a specific group using XML intersection points
        /// 
        /// CORRECTED FLAG PROTECTION LOGIC:
        /// - IsClustered = true → Already clustered (SKIP completely)
        /// - IsClustered = false → Processed but not proximate (SKIP - keep as individual)
        /// - IsClustered = null → Not yet processed (ALLOW clustering - check proximity)
        ///
        /// NOTE: Only IsClustered = null sleeves should be processed for clustering.
        /// After proximity check: proximate sleeves → IsClustered = true, non-proximate → IsClustered = false
        /// 
        /// This prevents duplicate clustering of sleeves that were already clustered in previous runs.
        /// This is the core "calculate once" logic that replaces expensive Revit API calls
        /// </summary>
        private List<List<ClashZone>> CalculateClustersForGroup(List<ClashZone> zones, double toleranceDistance, string groupKey)
        {
            var clusters = new List<List<ClashZone>>();
            
            // ✅ CRITICAL PROTECTION: Only process zones that haven't been processed for clustering yet
            // MarkedForClusteringSleeveProcess = null → Not yet processed (ALLOW clustering)
            // MarkedForClusteringSleeveProcess = true → Already marked for clustering (SKIP)
            // MarkedForClusteringSleeveProcess = false → Processed but not proximate (SKIP)
            var unprocessedZones = zones.Where(z => z.MarkedForClusteringSleeveProcess == null).ToList();
            
            if (unprocessedZones.Count != zones.Count)
            {
                var processedCount = zones.Count - unprocessedZones.Count;
                DebugLogger.Info($"[PreCalculatedClusterService] Filtered out {processedCount} already processed zones, processing {unprocessedZones.Count} unprocessed zones");
                
                // ⚠️ CRITICAL: Log flag states of filtered zones
                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\flag_state_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] [PRE-CALC-FILTER] Filtered out {processedCount} already processed zones\n");
                
                foreach (var zone in zones.Where(z => z.MarkedForClusteringSleeveProcess != null).Take(3)) // Log first 3 filtered zones
                {
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\flag_state_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] [PRE-CALC-FILTER] FILTERED ClashZone {zone.Id}: MarkedForClustering={zone.MarkedForClusteringSleeveProcess}, IsResolved={zone.IsResolved}, IsClusterResolved={zone.IsClusterResolved}\n");
                }
            }
            
            var zonesToProcess = new HashSet<ClashZone>(unprocessedZones);
            
            // Extract host type and orientation from group key
            var parts = groupKey.Split('_');
            var hostType = parts.Length > 1 ? parts[1] : "Unknown";
            var orientation = parts.Length > 2 ? parts[2] : "";
            
            DebugLogger.Info($"[PreCalculatedClusterService] Calculating clusters for {zones.Count} zones, hostType={hostType}, orientation={orientation}");
            
            while (zonesToProcess.Count > 0)
            {
                var startZone = zonesToProcess.First();
                var cluster = new List<ClashZone>();
                var queue = new Queue<ClashZone>();
                
                queue.Enqueue(startZone);
                zonesToProcess.Remove(startZone);
                
                while (queue.Count > 0)
                {
                    var currentZone = queue.Dequeue();
                    cluster.Add(currentZone);
                    
                    // Find proximate zones using XML intersection points
                    var proximateZones = FindProximateZones(currentZone, zonesToProcess, toleranceDistance, hostType, orientation);
                    
                    foreach (var proximateZone in proximateZones)
                    {
                        if (zonesToProcess.Remove(proximateZone))
                        {
                            queue.Enqueue(proximateZone);
                        }
                    }
                }
                
                // Only add clusters with more than 1 zone
                if (cluster.Count > 1)
                {
                    clusters.Add(cluster);
                    DebugLogger.Info($"[PreCalculatedClusterService] Formed cluster with {cluster.Count} zones");
                }
                else
                {
                    DebugLogger.Info($"[PreCalculatedClusterService] Individual zone (no proximate neighbors)");
                }
            }
            
            // ✅ CRITICAL: Ensure all sleeves have a flag value after proximity checking
            // Any sleeves that weren't processed (isolated) should be marked as MarkedForClusteringSleeveProcess = false
            foreach (var zone in zones)
            {
                if (zone.MarkedForClusteringSleeveProcess == null)
                {
                    zone.MarkedForClusteringSleeveProcess = false;
                    DebugLogger.Info($"[PreCalculatedClusterService] Set MarkedForClusteringSleeveProcess = false for isolated ClashZone {zone.Id} (no proximate neighbors)");
                }
            }
            
            return clusters;
        }
        
        /// <summary>
        /// Find proximate zones using XML intersection points
        /// This replaces the expensive Revit API bounding box calculations
        /// </summary>
        private List<ClashZone> FindProximateZones(ClashZone currentZone, HashSet<ClashZone> candidateZones, double toleranceDistance, string hostType, string orientation)
        {
            var proximateZones = new List<ClashZone>();
            
            foreach (var candidateZone in candidateZones)
            {
                if (candidateZone == currentZone) continue;
                
                // Calculate distance using XML intersection points
                double distance = CalculateDistanceBetweenZones(currentZone, candidateZone, hostType, orientation);
                
                if (distance <= toleranceDistance)
                {
                    proximateZones.Add(candidateZone);
                    
                    // ✅ CRITICAL: Set MarkedForClusteringSleeveProcess = true for proximate sleeves (should be clustered)
                    // This happens DURING proximity check, not after cluster formation
                    currentZone.MarkedForClusteringSleeveProcess = true;
                    candidateZone.MarkedForClusteringSleeveProcess = true;
                    
                    DebugLogger.Info($"[PreCalculatedClusterService] Set MarkedForClusteringSleeveProcess = true for ClashZone {currentZone.Id} and {candidateZone.Id} (proximate - will be clustered)");
                    
                    // Log proximity analysis to main cluster debug log
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log",
                        $"✓ PROXIMITY: Zone {currentZone.Id} -> {candidateZone.Id}: {UnitUtils.ConvertFromInternalUnits(distance, UnitTypeId.Millimeters):F1}mm <= {UnitUtils.ConvertFromInternalUnits(toleranceDistance, UnitTypeId.Millimeters):F1}mm\n");
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log",
                        $"  MEP1: {currentZone.MepElementId.IntegerValue} ({currentZone.MepElementCategory}) -> MEP2: {candidateZone.MepElementId.IntegerValue} ({candidateZone.MepElementCategory})\n");
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log",
                        $"  Sleeve1: {UnitUtils.ConvertFromInternalUnits(currentZone.SleeveWidth, UnitTypeId.Millimeters):F1}x{UnitUtils.ConvertFromInternalUnits(currentZone.SleeveHeight, UnitTypeId.Millimeters):F1}mm at {currentZone.SleevePlacementPoint}\n");
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log",
                        $"  Sleeve2: {UnitUtils.ConvertFromInternalUnits(candidateZone.SleeveWidth, UnitTypeId.Millimeters):F1}x{UnitUtils.ConvertFromInternalUnits(candidateZone.SleeveHeight, UnitTypeId.Millimeters):F1}mm at {candidateZone.SleevePlacementPoint}\n");
                }
                else
                {
                    // Log non-proximate analysis to main cluster debug log
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log",
                        $"✗ NO PROXIMITY: Zone {currentZone.Id} -> {candidateZone.Id}: {UnitUtils.ConvertFromInternalUnits(distance, UnitTypeId.Millimeters):F1}mm > {UnitUtils.ConvertFromInternalUnits(toleranceDistance, UnitTypeId.Millimeters):F1}mm\n");
                    
                    // ✅ CRITICAL: Set MarkedForClusteringSleeveProcess = false for non-proximate sleeves (should remain individual)
                    // Only set to false if it hasn't been set to true by another proximity check
                    if (currentZone.MarkedForClusteringSleeveProcess != true)
                    {
                        currentZone.MarkedForClusteringSleeveProcess = false;
                        DebugLogger.Info($"[PreCalculatedClusterService] Set MarkedForClusteringSleeveProcess = false for ClashZone {currentZone.Id} (not proximate - keep as individual)");
                    }
                    if (candidateZone.MarkedForClusteringSleeveProcess != true)
                    {
                        candidateZone.MarkedForClusteringSleeveProcess = false;
                        DebugLogger.Info($"[PreCalculatedClusterService] Set MarkedForClusteringSleeveProcess = false for ClashZone {candidateZone.Id} (not proximate - keep as individual)");
                    }
                    
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\precalculated_cluster_debug.log",
                        $"XML NO-PROXIMITY: Zone {currentZone.Id} -> {candidateZone.Id}: {UnitUtils.ConvertFromInternalUnits(distance, UnitTypeId.Millimeters):F1}mm > {UnitUtils.ConvertFromInternalUnits(toleranceDistance, UnitTypeId.Millimeters):F1}mm ✗\n");
                }
            }
            
            return proximateZones;
        }
        
        /// <summary>
        /// Calculate minimum distance between two rectangular sleeves using corner coordinates
        /// This properly detects overlapping sleeves even when centers are far apart
        /// </summary>
        private double CalculateDistanceBetweenZones(ClashZone zone1, ClashZone zone2, string hostType, string orientation)
        {
            // ✅ CRITICAL: Use saved sleeve corner coordinates (4 corners of actual placed sleeves)
            // ✅ CRITICAL: Use active document coordinates for proximity calculation
            if (zone1.SleevePlacementPointActiveDocument == null || zone2.SleevePlacementPointActiveDocument == null)
            {
                DebugLogger.Warning($"[PreCalculatedClusterService] Missing active document coordinates: Zone1={zone1.SleevePlacementPointActiveDocument}, Zone2={zone2.SleevePlacementPointActiveDocument}");
                return double.MaxValue; // No sleeve placement point data
            }
            
            // Get the 4 corner coordinates of each placed sleeve directly
            var corners1 = GetSleeveCorners(zone1.SleevePlacementPointActiveDocument, zone1.SleeveWidth, zone1.SleeveHeight, hostType, orientation);
            var corners2 = GetSleeveCorners(zone2.SleevePlacementPointActiveDocument, zone2.SleeveWidth, zone2.SleeveHeight, hostType, orientation);
            
            // ✅ CRITICAL DEBUG: Log the actual coordinates being used
            DebugLogger.Info($"[PreCalculatedClusterService] Zone1({zone1.MepElementId.IntegerValue}): Center={zone1.SleevePlacementPointActiveDocument}, W={UnitUtils.ConvertFromInternalUnits(zone1.SleeveWidth, UnitTypeId.Millimeters):F1}mm, H={UnitUtils.ConvertFromInternalUnits(zone1.SleeveHeight, UnitTypeId.Millimeters):F1}mm");
            DebugLogger.Info($"[PreCalculatedClusterService] Zone2({zone2.MepElementId.IntegerValue}): Center={zone2.SleevePlacementPointActiveDocument}, W={UnitUtils.ConvertFromInternalUnits(zone2.SleeveWidth, UnitTypeId.Millimeters):F1}mm, H={UnitUtils.ConvertFromInternalUnits(zone2.SleeveHeight, UnitTypeId.Millimeters):F1}mm");
            
            // 🔥 DEBUG: Log detailed coordinate and dimension info for distance calculation
            File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                $"[DISTANCE-CALC] Zone1({zone1.MepElementId.IntegerValue}): ActiveDoc={zone1.SleevePlacementPointActiveDocument}, W={UnitUtils.ConvertFromInternalUnits(zone1.SleeveWidth, UnitTypeId.Millimeters):F1}mm, H={UnitUtils.ConvertFromInternalUnits(zone1.SleeveHeight, UnitTypeId.Millimeters):F1}mm\n");
            File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                $"[DISTANCE-CALC] Zone2({zone2.MepElementId.IntegerValue}): ActiveDoc={zone2.SleevePlacementPointActiveDocument}, W={UnitUtils.ConvertFromInternalUnits(zone2.SleeveWidth, UnitTypeId.Millimeters):F1}mm, H={UnitUtils.ConvertFromInternalUnits(zone2.SleeveHeight, UnitTypeId.Millimeters):F1}mm\n");
            
            // Find minimum distance between any corner of sleeve1 to any corner of sleeve2
            double minDistance = double.MaxValue;
            
            foreach (var corner1 in corners1)
            {
                foreach (var corner2 in corners2)
                {
                    double distance = Calculate2DDistance(corner1, corner2, hostType, orientation);
                    if (distance < minDistance)
                    {
                        minDistance = distance;
                    }
                }
            }
            
            // ✅ CRITICAL DEBUG: Log the calculated distance
            DebugLogger.Info($"[PreCalculatedClusterService] Rectangle distance: Zone1({zone1.MepElementId.IntegerValue}) to Zone2({zone2.MepElementId.IntegerValue}) = {UnitUtils.ConvertFromInternalUnits(minDistance, UnitTypeId.Millimeters):F1}mm");
            
            // 🔥 DEBUG: Log final distance calculation result
            File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                $"[DISTANCE-RESULT] Zone1({zone1.MepElementId.IntegerValue}) to Zone2({zone2.MepElementId.IntegerValue}) = {UnitUtils.ConvertFromInternalUnits(minDistance, UnitTypeId.Millimeters):F1}mm\n");
            
            return minDistance;
        }
        
        /// <summary>
        /// Get 4 corner coordinates of a rectangular sleeve
        /// </summary>
        private List<XYZ> GetSleeveCorners(XYZ center, double width, double height, string hostType, string orientation, double slabThickness = 0.5)
        {
            var corners = new List<XYZ>();
            
            double halfWidth = width / 2.0;
            double halfHeight = height / 2.0;
            
            if (hostType == "Floor")
            {
                // Floor: Use X and Y coordinates, Z varies by slab thickness
                // For floor sleeves, corners should be at top and bottom of slab
                corners.Add(new XYZ(center.X - halfWidth, center.Y - halfHeight, center.Z - slabThickness/2)); // Bottom-left (bottom of slab)
                corners.Add(new XYZ(center.X + halfWidth, center.Y - halfHeight, center.Z - slabThickness/2)); // Bottom-right (bottom of slab)
                corners.Add(new XYZ(center.X - halfWidth, center.Y + halfHeight, center.Z + slabThickness/2)); // Top-left (top of slab)
                corners.Add(new XYZ(center.X + halfWidth, center.Y + halfHeight, center.Z + slabThickness/2)); // Top-right (top of slab)
            }
            else if (hostType == "Wall" || hostType == "Structural Framing")
            {
                if (orientation == "X")
                {
                    // X-oriented wall: Use Y and Z coordinates
                    corners.Add(new XYZ(center.X, center.Y - halfWidth, center.Z - halfHeight)); // Bottom-left
                    corners.Add(new XYZ(center.X, center.Y + halfWidth, center.Z - halfHeight)); // Bottom-right
                    corners.Add(new XYZ(center.X, center.Y - halfWidth, center.Z + halfHeight)); // Top-left
                    corners.Add(new XYZ(center.X, center.Y + halfWidth, center.Z + halfHeight)); // Top-right
                }
                else
                {
                    // Y-oriented wall: Use X and Z coordinates
                    corners.Add(new XYZ(center.X - halfWidth, center.Y, center.Z - halfHeight)); // Bottom-left
                    corners.Add(new XYZ(center.X + halfWidth, center.Y, center.Z - halfHeight)); // Bottom-right
                    corners.Add(new XYZ(center.X - halfWidth, center.Y, center.Z + halfHeight)); // Top-left
                    corners.Add(new XYZ(center.X + halfWidth, center.Y, center.Z + halfHeight)); // Top-right
                }
            }
            else
            {
                // Default: Use all coordinates
                corners.Add(new XYZ(center.X - halfWidth, center.Y - halfHeight, center.Z - halfHeight));
                corners.Add(new XYZ(center.X + halfWidth, center.Y - halfHeight, center.Z - halfHeight));
                corners.Add(new XYZ(center.X - halfWidth, center.Y + halfHeight, center.Z + halfHeight));
                corners.Add(new XYZ(center.X + halfWidth, center.Y + halfHeight, center.Z + halfHeight));
            }
            
            return corners;
        }
        
        /// <summary>
        /// Calculate 2D distance based on host type and orientation
        /// </summary>
        private double Calculate2DDistance(XYZ point1, XYZ point2, string hostType, string orientation)
        {
            if (hostType == "Floor")
            {
                // Floor: Use X and Y only
                double dx = Math.Abs(point1.X - point2.X);
                double dy = Math.Abs(point1.Y - point2.Y);
                return Math.Sqrt(dx * dx + dy * dy);
            }
            else if (hostType == "Wall" || hostType == "Structural Framing")
            {
                if (orientation == "X")
                {
                    // X-oriented wall: Use Y and Z only
                    double dy = Math.Abs(point1.Y - point2.Y);
                    double dz = Math.Abs(point1.Z - point2.Z);
                    return Math.Sqrt(dy * dy + dz * dz);
                }
                else
                {
                    // Y-oriented wall: Use X and Z only
                    double dx = Math.Abs(point1.X - point2.X);
                    double dz = Math.Abs(point1.Z - point2.Z);
                    return Math.Sqrt(dx * dx + dz * dz);
                }
            }
            else
            {
                // Default: Use all coordinates
                double dx = Math.Abs(point1.X - point2.X);
                double dy = Math.Abs(point1.Y - point2.Y);
                double dz = Math.Abs(point1.Z - point2.Z);
                return Math.Sqrt(dx * dx + dy * dy + dz * dz);
            }
        }
        
        /// <summary>
        /// Clear all pre-calculated clusters (useful for testing or re-calculation)
        /// </summary>
        public void ClearPreCalculatedClusters()
        {
            _preCalculatedClusters.Clear();
            _toleranceDistances.Clear();
            DebugLogger.Info("[PreCalculatedClusterService] Cleared all pre-calculated clusters");
        }
        
        /// <summary>
        /// Get statistics about pre-calculated clusters
        /// </summary>
        public Dictionary<string, object> GetClusterStatistics()
        {
            var stats = new Dictionary<string, object>
            {
                ["TotalGroups"] = _preCalculatedClusters.Count,
                ["TotalClusters"] = _preCalculatedClusters.Values.Sum(clusters => clusters.Count),
                ["TotalIndividualZones"] = _preCalculatedClusters.Values.Sum(clusters => clusters.Count(c => c.Count == 1)),
                ["TotalClusteredZones"] = _preCalculatedClusters.Values.Sum(clusters => clusters.Where(c => c.Count > 1).Sum(c => c.Count))
            };
            
            return stats;
        }
    }
}
