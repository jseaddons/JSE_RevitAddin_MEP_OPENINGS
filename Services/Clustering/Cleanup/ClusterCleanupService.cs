using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Cleanup;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Cleanup
{
    /// <summary>
    /// Extracted cleanup logic from UniversalClusterService. Simplified: retains safety checks and protection set.
    /// </summary>
    public class ClusterCleanupService : IClusterCleanupService
    {
        /// <summary>
        /// ✅ CRITICAL FIX: Accept deferred parameters dictionary to read correct dimensions when batching is enabled.
        /// This prevents cleanup from using stale parameter values from Revit before flush/regeneration.
        /// </summary>
        public int CleanupSleevesWithinClusters(Document doc, List<FamilyInstance> placedClusters, Dictionary<ElementId, Dictionary<string, object>> deferredParameters = null, string targetCategory = null)
        {
            int deletedCount = 0;
            try
            {
                if (placedClusters == null || placedClusters.Count == 0)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: No placed clusters, skipping cleanup\n");
                    return 0;
                }

                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Starting cleanup with {placedClusters.Count} placed clusters (TargetCategory={targetCategory ?? "NONE"})\n");

                // ✅ CRITICAL PROTECTION: Build protection set from FRESH elements
                var clusterSleeveIds = new HashSet<int>();
                foreach (var cluster in placedClusters)
                {
                    if (cluster == null) continue;
                    int clusterId = cluster.Id.IntegerValue;
                    
                    var freshCluster = doc.GetElement(cluster.Id) as FamilyInstance;
                    if (freshCluster != null && freshCluster.IsValidObject)
                    {
                        clusterSleeveIds.Add(clusterId);
                    }
                }

                // Collect all potential sleeve family instances first (coarse filter)
                var potentialSleeves = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(s =>
                    {
                        string fam = s.Symbol?.FamilyName ?? string.Empty;
                        bool keyword = fam.Contains("Sleeve", StringComparison.OrdinalIgnoreCase) || fam.Contains("Opening", StringComparison.OrdinalIgnoreCase);
                        bool known = fam.Contains("CircularOpening", StringComparison.OrdinalIgnoreCase) || fam.Contains("RectangularOpening", StringComparison.OrdinalIgnoreCase);
                        
                        string catName = s.Category?.Name ?? "";
                        bool validCategory = catName == "Generic Models" || 
                                             catName == "Structural Connections" || 
                                             catName == "Duct Accessories" ||
                                             catName == "Pipe Accessories" ||
                                             catName == "Mechanical Equipment";
                                             
                        return validCategory && (keyword || known);
                    })
                    .ToList();
                    
                // ✅ DB LOOKUP: Get MEP Categories using "dump once, use many times" principle
                // Querying DB for all potential sleeves to get their source-of-truth category
                var sleeveCategories = new Dictionary<int, string>();
                try
                {
                    using (var context = new SleeveDbContext(doc))
                    {
                        var repo = new ClashZoneRepository(context, msg => SafeFileLogger.SafeAppendText("cluster_debug.log", $"[Repo] {msg}\n"));
                        var ids = potentialSleeves.Select(s => s.Id.IntegerValue).ToList();
                        sleeveCategories = repo.GetMepCategoriesForSleeveIds(ids);
                        
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Retrieved {sleeveCategories.Count} category records from DB for {ids.Count} potential sleeves\n");
                    }
                }
                catch (Exception ex)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ CLEANUP: Failed to read categories from DB: {ex.Message} - Falling back to element parameters\n");
                }

                // Filter sleeves using DB category (primary) or element parameter (fallback)
                var allSleeves = potentialSleeves.Where(s =>
                {
                    // ✅ CATEGORY SPECIFIC FILTERING TO PREVENT CROSS-CATEGORY DELETION
                    bool isCrossCategory = false;
                    string catName = s.Category?.Name ?? "";
                    string fam = s.Symbol?.FamilyName ?? "";
                    
                    // 1. Check "MEP Category" - Priority: DB -> Parameter -> Empty
                    string mepCatValue = "";
                    
                    if (sleeveCategories.ContainsKey(s.Id.IntegerValue))
                    {
                        mepCatValue = sleeveCategories[s.Id.IntegerValue];
                    }
                    else
                    {
                        // Fallback to parameter if not in DB
                        var mepCatParam = s.LookupParameter("MEP Category");
                        mepCatValue = mepCatParam?.AsString() ?? "";
                    }

                    if (!string.IsNullOrEmpty(targetCategory))
                    {
                        // If sleeve has explicit MEP Category, check compatibility directly
                        if (!string.IsNullOrEmpty(mepCatValue))
                        {
                            // If I am placing "Pipes" -> Protect Ducts, Trays, and ALL Accessories/Dampers
                            if (targetCategory.Contains("Pipe") && (mepCatValue.Contains("Duct") || mepCatValue.Contains("Tray") || mepCatValue.Contains("Accessory") || mepCatValue.Contains("Damper"))) isCrossCategory = true;
                            
                            // If I am placing "Ducts" -> Protect Pipes, Trays, and Dampers/Accessories (don't absorb dampers into duct clusters)
                            // User Request: "delete only sleeves belog to that mep cat" (Damper != Duct)
                            if (targetCategory.Contains("Duct") && (mepCatValue.Contains("Pipe") || mepCatValue.Contains("Tray") || mepCatValue.Contains("Accessory") || mepCatValue.Contains("Damper"))) isCrossCategory = true;
                            
                            // If I am placing "Trays" -> Protect everything else
                            if (targetCategory.Contains("Tray") && !mepCatValue.Contains("Tray")) isCrossCategory = true;
                        }
                        else 
                        {
                            // Fallback: Check Revit Category & Family Name
                            if (targetCategory.Contains("Pipe"))
                            {
                                // Protect Ducts, Accessories, Dampers
                                if (catName.Contains("Duct") || catName.Contains("Accessory") || fam.Contains("Damper", StringComparison.OrdinalIgnoreCase)) isCrossCategory = true;
                            }
                            else if (targetCategory.Contains("Duct"))
                            {
                                // Protect Pipes, Accessories, Dampers
                                if (catName.Contains("Pipe") || catName.Contains("Accessory") || fam.Contains("Damper", StringComparison.OrdinalIgnoreCase)) isCrossCategory = true;
                            }
                            else if (targetCategory.Contains("Tray"))
                            {
                                // Protect Pipes, Ducts
                                if (catName.Contains("Pipe") || catName.Contains("Duct")) isCrossCategory = true;
                            }
                        }
                    }

                    bool matched = !isCrossCategory;

                    if (matched)
                    {
                         // Log matched candidates temporarily
                         SafeFileLogger.SafeAppendText("cluster_debug_verbose.log", 
                             $"[{DateTime.Now:HH:mm:ss}] 🗑 CANDIDATE: ID={s.Id}, Cat={catName}, Fam={fam}, MepCat='{mepCatValue}' (Source={(sleeveCategories.ContainsKey(s.Id.IntegerValue) ? "DB" : "Param")}), Target={targetCategory ?? "NULL"}\n");
                    }
                    else if (isCrossCategory)
                    {
                         SafeFileLogger.SafeAppendText("cluster_debug_verbose.log", 
                             $"[{DateTime.Now:HH:mm:ss}] 🛡 PROTECTED: ID={s.Id}, Cat={catName}, Fam={fam}, MepCat='{mepCatValue}' (Source={(sleeveCategories.ContainsKey(s.Id.IntegerValue) ? "DB" : "Param")}) (Cross-Category with {targetCategory})\n");
                    }

                    return matched;
                })
                .ToList();

                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Found {allSleeves.Count} potential sleeves (Filtered for Category Compatibility)\n");

                var individualSleeves = new List<FamilyInstance>();
                foreach (var s in allSleeves)
                {
                    int id = s.Id.IntegerValue;
                    if (clusterSleeveIds.Contains(id)) continue; 
                    
                    var clusterParam = s.LookupParameter("Cluster Sleeve Instance ID");
                    var sleeveInstanceParam = s.LookupParameter("Sleeve Instance ID");
                    int clusterValue = clusterParam?.AsInteger() ?? -1;
                    int sleeveInstanceValue = sleeveInstanceParam?.AsInteger() ?? -999;
                    
                    // ✅ RELAXED SAFETY CHECK: Even if sleeveInstanceValue is -1, it might be an orphaned individual sleeve.
                    // We only skip if it IS one of the newly placed clusters.
                    // (sleeveInstanceValue == -1) check removed to ensure persistent orphans are caught.
                    if (clusterValue > 0 && clusterValue == id) continue;

                    individualSleeves.Add(s);
                }

                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Found {individualSleeves.Count} candidate individual sleeves for overlap check\n");

                if (individualSleeves.Count == 0) return 0;

                // Build cluster sleeve bounding boxes (axis-aligned)
                // ✅ CRITICAL: Use active view for bounding box calculation (matches methodology requirement)
                // According to methodology: "Check for other individual sleeves that fall within the cluster bounding box"
                // The bounding box must be calculated from the actual placed cluster sleeve element
                var clusterBboxes = new List<BoundingBoxXYZ>();
                var clusterBboxMap = new Dictionary<int, BoundingBoxXYZ>(); // Map cluster ID to bbox for logging
                foreach (var cluster in placedClusters)
                {
                    if (cluster == null) continue;
                    int clusterId = cluster.Id.IntegerValue;
                    
                    // ✅ CRITICAL FIX: Refresh element and get fresh bounding box
                    // Sometimes bounding boxes are stale after placement, so we get the element fresh
                    var freshCluster = doc.GetElement(cluster.Id) as FamilyInstance;
                    if (freshCluster == null)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ CLEANUP: Cluster {clusterId} not found in document - SKIPPING\n");
                        continue;
                    }
                    
                    // ✅ CRITICAL FIX: Calculate bounding box from placement point + parameters
                    // Don't rely on get_BoundingBox() which may be stale if parameters haven't been flushed/regenerated
                    // Get placement point from instance
                    XYZ placementPoint = null;
                    if (freshCluster.Location is LocationPoint locationPoint)
                    {
                        placementPoint = locationPoint.Point;
                    }
                    else if (freshCluster.Location is LocationCurve locationCurve)
                    {
                        var curve = locationCurve.Curve;
                        if (curve != null)
                        {
                            placementPoint = curve.GetEndPoint(0);
                        }
                    }
                    
                    // ✅ CRITICAL FIX: Read dimensions from deferred parameters first (if batching enabled), then fallback to Revit element
                    // This prevents reading stale values before flush/regeneration
                    double paramWidth = 0.0;
                    double paramHeight = 0.0;
                    double paramDepth = 0.0;
                    
                    // Try to read from deferred parameters first (if batching is enabled)
                    if (OptimizationFlags.UseBatchedParameterWrites && deferredParameters != null && deferredParameters.ContainsKey(freshCluster.Id))
                    {
                        var deferredParams = deferredParameters[freshCluster.Id];
                        if (deferredParams.ContainsKey("Width") && deferredParams["Width"] is double w) paramWidth = w;
                        if (deferredParams.ContainsKey("Height") && deferredParams["Height"] is double h) paramHeight = h;
                        if (deferredParams.ContainsKey("Depth") && deferredParams["Depth"] is double d) paramDepth = d;
                        
                        if (!DeploymentConfiguration.DeploymentMode && (paramWidth > 0.001 || paramHeight > 0.001 || paramDepth > 0.001))
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Cluster {clusterId} dimensions from DEFERRED DICT - W={paramWidth*304.8:F1}mm, H={paramHeight*304.8:F1}mm, D={paramDepth*304.8:F1}mm\n");
                        }
                    }
                    
                    // Fallback to Revit element parameters if deferred parameters not available
                    if (paramWidth <= 0.001 || paramHeight <= 0.001 || paramDepth <= 0.001)
                    {
                        var widthParam = freshCluster.LookupParameter("Width");
                        var heightParam = freshCluster.LookupParameter("Height");
                        var depthParam = freshCluster.LookupParameter("Depth");
                        if (paramWidth <= 0.001) paramWidth = widthParam?.AsDouble() ?? 0.0;
                        if (paramHeight <= 0.001) paramHeight = heightParam?.AsDouble() ?? 0.0;
                        if (paramDepth <= 0.001) paramDepth = depthParam?.AsDouble() ?? 0.0;
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Cluster {clusterId} dimensions from REVIT ELEMENT - W={paramWidth*304.8:F1}mm, H={paramHeight*304.8:F1}mm, D={paramDepth*304.8:F1}mm\n");
                        }
                    }
                    
                    // ✅ CRITICAL: Determine which BBox calculation to use based on safety flags
                    bool handled = false;
                    
                    // Identify host type to pick correct flag
                    bool isHostedOnFloor = freshCluster.Host is Floor;
                    bool isHostedOnWall = freshCluster.Host is Wall;
                    bool useGeometricCenter = (isHostedOnFloor && OptimizationFlags.UseGeometricCenterForFloors) || 
                                             (isHostedOnWall && OptimizationFlags.UseGeometricCenterForWalls);

                    if (useGeometricCenter)
                    {
                        // Prioritize Revit's native BoundingBox (Correct for Rotated Clusters)
                        var nativeBbox = freshCluster.get_BoundingBox(null);
                        if (nativeBbox != null && nativeBbox.Enabled)
                        {
                            clusterBboxes.Add(nativeBbox);
                            clusterBboxMap[clusterId] = nativeBbox;
                            
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Cluster {clusterId} bbox (NATIVE get_BoundingBox - FLAG ENABLED) - " +
                                $"Host={(isHostedOnFloor ? "FLOOR" : isHostedOnWall ? "WALL" : "OTHER")}, " +
                                $"Min=({nativeBbox.Min.X:F3}, {nativeBbox.Min.Y:F3}, {nativeBbox.Min.Z:F3}), " +
                                $"Max=({nativeBbox.Max.X:F3}, {nativeBbox.Max.Y:F3}, {nativeBbox.Max.Z:F3})\n");
                            
                            handled = true;
                        }
                    }

                    if (!handled && placementPoint != null && paramWidth > 0.001 && paramHeight > 0.001 && paramDepth > 0.001)
                    {
                        // LEGACY: Calculate bounding box from placement point + dimensions
                        double halfWidth = paramWidth / 2.0;
                        double halfHeight = paramHeight / 2.0;
                        double halfDepth = paramDepth / 2.0;
                        
                        var calculatedBbox = new BoundingBoxXYZ();
                        calculatedBbox.Min = new XYZ(
                            placementPoint.X - halfWidth,
                            placementPoint.Y - halfHeight,
                            placementPoint.Z - halfDepth
                        );
                        calculatedBbox.Max = new XYZ(
                            placementPoint.X + halfWidth,
                            placementPoint.Y + halfHeight,
                            placementPoint.Z + halfDepth
                        );
                        calculatedBbox.Enabled = true;
                        
                        clusterBboxes.Add(calculatedBbox);
                        clusterBboxMap[clusterId] = calculatedBbox;
                        
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Cluster {clusterId} bbox (CALCULATED from params - FLAG DISABLED) - " +
                            $"PlacementPoint=({placementPoint.X:F3}, {placementPoint.Y:F3}, {placementPoint.Z:F3}), " +
                            $"Min=({calculatedBbox.Min.X:F3}, {calculatedBbox.Min.Y:F3}, {calculatedBbox.Min.Z:F3}), " +
                            $"Max=({calculatedBbox.Max.X:F3}, {calculatedBbox.Max.Y:F3}, {calculatedBbox.Max.Z:F3}), " +
                            $"BBoxSize=({paramWidth*304.8:F1}mm x {paramHeight*304.8:F1}mm x {paramDepth*304.8:F1}mm), " +
                            $"Params=(W={paramWidth*304.8:F1}mm, H={paramHeight*304.8:F1}mm, D={paramDepth*304.8:F1}mm)\n");
                        
                        handled = true;
                    }

                    if (!handled)
                    {
                        // Fallback to get_BoundingBox() if parameters are missing
                        var bbox = freshCluster.get_BoundingBox(null);
                        if (bbox != null && bbox.Enabled)
                        {
                            clusterBboxes.Add(bbox);
                            clusterBboxMap[clusterId] = bbox;
                            
                            double bboxWidth = (bbox.Max.X - bbox.Min.X) * 304.8;
                            double bboxHeight = (bbox.Max.Y - bbox.Min.Y) * 304.8;
                            double bboxDepth = (bbox.Max.Z - bbox.Min.Z) * 304.8;
                            
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Cluster {clusterId} bbox (FALLBACK from get_BoundingBox) - " +
                                $"Min=({bbox.Min.X:F3}, {bbox.Min.Y:F3}, {bbox.Min.Z:F3}), " +
                                $"Max=({bbox.Max.X:F3}, {bbox.Max.Y:F3}, {bbox.Max.Z:F3}), " +
                                $"BBoxSize=({bboxWidth:F1}mm x {bboxHeight:F1}mm x {bboxDepth:F1}mm), " +
                                $"Params=(W={paramWidth*304.8:F1}mm, H={paramHeight*304.8:F1}mm, D={paramDepth*304.8:F1}mm)\n");
                        }
                        else
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ CLEANUP: Cluster {clusterId} has no valid placement point or parameters - SKIPPING\n");
                        }
                    }
                }

                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Built {clusterBboxes.Count} cluster bounding boxes from {placedClusters.Count} placed clusters\n");

                // ✅ NEW: DATABASE-DRIVEN FALLBACK (User Request)
                // Use the constituent GUIDs from ClusterSleeves_v2 to identify sleeves that MUST be deleted.
                // This catches cases where geometric checks (proximity) fail but DB knows the mapping.
                var dbTrackedSleeveIds = new List<ElementId>();
                try
                {
                    using (var context = new SleeveDbContext(doc))
                    {
                        var repo = new ClashZoneRepository(context, _ => { });
                        foreach (var cluster in placedClusters)
                        {
                            if (cluster == null) continue;
                            int clusterRevId = cluster.Id.IntegerValue;
                            
                            // Query ClusterSleeves_v2 for constituents
                            string cGuidsStr = "";
                            using (var qCmd = context.Connection.CreateCommand())
                            {
                                qCmd.CommandText = "SELECT ConstituentZoneGuids FROM ClusterSleeves_v2 WHERE ClusterInstanceId = @id";
                                qCmd.Parameters.AddWithValue("@id", clusterRevId);
                                var qRes = qCmd.ExecuteScalar();
                                if (qRes != null && qRes != DBNull.Value) cGuidsStr = qRes.ToString();
                            }
                            
                            if (!string.IsNullOrEmpty(cGuidsStr))
                            {
                                var gList = cGuidsStr.Split(new[] { ',', ';', ' ', '[', ']', '\"' }, StringSplitOptions.RemoveEmptyEntries)
                                                     .Select(g => Guid.TryParse(g.Trim(), out var guid) ? guid : Guid.Empty)
                                                     .Where(g => g != Guid.Empty)
                                                     .ToList();
                                
                                if (gList.Any())
                                {
                                    var zones = repo.GetClashZonesByGuids(gList);
                                    foreach (var z in zones)
                                    {
                                        // Use AfterClusterSleevePlacedSleeveInstanceId (captured before reset) OR SleeveInstanceId (if not yet reset)
                                        int targetSleeveId = z.AfterClusterSleevePlacedSleeveInstanceId > 0 ? z.AfterClusterSleevePlacedSleeveInstanceId : z.SleeveInstanceId;
                                        
                                        if (targetSleeveId > 0)
                                        {
                                            var eid = new ElementId(targetSleeveId);
                                            // Check if element still exists in Revit
                                            var el = doc.GetElement(eid);
                                            if (el != null && el.IsValidObject)
                                            {
                                                if (!clusterSleeveIds.Contains(targetSleeveId)) // Safety check: not a cluster itself
                                                {
                                                    dbTrackedSleeveIds.Add(eid);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception dbEx)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ CLEANUP: DB fallback failed: {dbEx.Message}\n");
                }

                if (clusterBboxes.Count == 0) return 0;

                var toDelete = new List<ElementId>();
                foreach (var individual in individualSleeves)
                {
                    int individualId = individual.Id.IntegerValue;
                    
                    // ✅ CRITICAL SAFETY CHECK: Double-check this is NOT a cluster sleeve
                    // This prevents any cluster sleeve from being marked for deletion, even if it somehow got into individualSleeves
                    if (clusterSleeveIds.Contains(individualId))
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ SAFETY CHECK: Sleeve {individualId} found in individualSleeves but is in protection set (cluster sleeve) - SKIPPING containment check\n");
                        continue;
                    }
                    
                    // ✅ IMPROVED LOGIC: Get individual sleeve bounding box for overlap check
                    // Check if individual sleeve's bounding box overlaps with cluster bounding box
                    // This is more robust than just checking placement point
                    var individualBbox = individual.get_BoundingBox(null);
                    if (individualBbox == null || !individualBbox.Enabled)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Individual sleeve {individualId} has no bounding box or is disabled - SKIPPING\n");
                        continue;
                    }
                    
                    // Get placement point for logging
                    XYZ sleevePlacementPoint = null;
                    if (individual.Location is LocationPoint locationPoint)
                    {
                        sleevePlacementPoint = locationPoint.Point;
                    }
                    else
                    {
                        sleevePlacementPoint = (individualBbox.Min + individualBbox.Max) / 2.0;
                    }
                    
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Checking individual sleeve {individualId} - " +
                        $"PlacementPoint=({sleevePlacementPoint.X:F3}, {sleevePlacementPoint.Y:F3}, {sleevePlacementPoint.Z:F3}), " +
                        $"BBox=({individualBbox.Min.X:F3},{individualBbox.Min.Y:F3},{individualBbox.Min.Z:F3}) to ({individualBbox.Max.X:F3},{individualBbox.Max.Y:F3},{individualBbox.Max.Z:F3})\n");
                    
                    bool matchedAnyCluster = false;
                    int clusterIndex = 0;
                    foreach (var kvp in clusterBboxMap)
                    {
                        int clusterId = kvp.Key;
                        var cbbox = kvp.Value;
                        clusterIndex++;
                        if (cbbox == null) continue;
                        
                        // ✅ IMPROVED LOGIC: Check bounding box overlap instead of just placement point
                        // Two bounding boxes overlap if they intersect on all three axes
                        // This catches edge cases where placement point is outside but bounding box overlaps
                        bool bboxOverlaps = individualBbox.Max.X >= cbbox.Min.X && individualBbox.Min.X <= cbbox.Max.X &&
                                           individualBbox.Max.Y >= cbbox.Min.Y && individualBbox.Min.Y <= cbbox.Max.Y &&
                                           individualBbox.Max.Z >= cbbox.Min.Z && individualBbox.Min.Z <= cbbox.Max.Z;
                        
                        // Also check if placement point is inside (for small sleeves where bbox might not overlap)
                        bool placementPointInside = sleevePlacementPoint.X >= cbbox.Min.X && sleevePlacementPoint.X <= cbbox.Max.X &&
                                                    sleevePlacementPoint.Y >= cbbox.Min.Y && sleevePlacementPoint.Y <= cbbox.Max.Y &&
                                                    sleevePlacementPoint.Z >= cbbox.Min.Z && sleevePlacementPoint.Z <= cbbox.Max.Z;
                        
                        // Delete if either bounding box overlaps OR placement point is inside
                        bool shouldDelete = bboxOverlaps || placementPointInside;
                        
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP:   vs Cluster #{clusterIndex} (ID={clusterId}) bbox - " +
                            $"Cluster=({cbbox.Min.X:F3},{cbbox.Min.Y:F3},{cbbox.Min.Z:F3}) to ({cbbox.Max.X:F3},{cbbox.Max.Y:F3},{cbbox.Max.Z:F3}), " +
                            $"IndividualBBox=({individualBbox.Min.X:F3},{individualBbox.Min.Y:F3},{individualBbox.Min.Z:F3}) to ({individualBbox.Max.X:F3},{individualBbox.Max.Y:F3},{individualBbox.Max.Z:F3}), " +
                            $"PlacementPoint=({sleevePlacementPoint.X:F3},{sleevePlacementPoint.Y:F3},{sleevePlacementPoint.Z:F3}), " +
                            $"bboxOverlaps={bboxOverlaps}, placementPointInside={placementPointInside}, shouldDelete={shouldDelete}\n");
                        
                        // ✅ CRITICAL: Delete if bounding box overlaps OR placement point is inside cluster bbox
                        if (shouldDelete)
                        {
                            // ✅ CRITICAL SAFETY CHECK: Triple-check this is NOT a cluster sleeve before marking for deletion
                            if (clusterSleeveIds.Contains(individualId))
                            {
                                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ CRITICAL SAFETY: Sleeve {individualId} matched containment check but is a cluster sleeve - NOT MARKING FOR DELETION\n");
                                matchedAnyCluster = true;
                                break;
                            }
                            
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: ✅ Individual sleeve {individualId} placement point is inside cluster #{clusterIndex} (ID={clusterId}) bbox - MARKING FOR DELETION\n");
                            toDelete.Add(individual.Id);
                            matchedAnyCluster = true;
                            break;
                        }
                    }
                    
                    if (!matchedAnyCluster)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: ❌ Individual sleeve {individualId} placement point is NOT inside any cluster bbox - KEEPING\n");
                    }
                }

                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Found {toDelete.Count} individual sleeves to delete via geometric check\n");

                // ✅ MERGE DB-TRACKED SLEEVES
                if (dbTrackedSleeveIds.Any())
                {
                    int addedFromDb = 0;
                    foreach (var dbId in dbTrackedSleeveIds)
                    {
                        if (!toDelete.Contains(dbId))
                        {
                            toDelete.Add(dbId);
                            addedFromDb++;
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: ✅ Individual sleeve {dbId.IntegerValue} added via DB mapping\n");
                        }
                    }
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Added {addedFromDb} additional sleeves via DB mapping (Total to delete: {toDelete.Count})\n");
                }

                if (toDelete.Count == 0) return 0;

                // ✅ CRITICAL PROTECTION: Final check before deletion - ensure NO cluster sleeve IDs are in toDelete list
                // This is a safety net in case filtering logic above failed
                var protectedIds = new HashSet<int>(clusterSleeveIds);
                var filteredToDelete = toDelete.Where(id => !protectedIds.Contains(id.IntegerValue)).ToList();
                int removedClusterSleeves = toDelete.Count - filteredToDelete.Count;
                
                if (removedClusterSleeves > 0)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ CRITICAL PROTECTION: Removed {removedClusterSleeves} cluster sleeve ID(s) from deletion list before batch delete!\n");
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Protection set: {string.Join(", ", protectedIds)}\n");
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Original toDelete IDs: {string.Join(", ", toDelete.Select(id => id.IntegerValue))}\n");
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Filtered toDelete IDs: {string.Join(", ", filteredToDelete.Select(id => id.IntegerValue))}\n");
                }
                
                if (filteredToDelete.Count == 0)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: After protection filter, no sleeves to delete\n");
                    return 0;
                }

                // ✅ CRITICAL: Check if we're already in a transaction
                // If doc.IsModifiable is true, we're in a transaction - don't create a new one
                bool alreadyInTransaction = doc.IsModifiable;
                
                if (alreadyInTransaction)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Already in transaction, batch deleting {filteredToDelete.Count} sleeves (protected {removedClusterSleeves} cluster sleeves)\n");
                    
                    // ✅ PERFORMANCE: Use batch delete instead of deleting one by one
                    try
                    {
                        var deletedIds = doc.Delete(filteredToDelete);
                        deletedCount = deletedIds.Count;
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: ✅ Batch deleted {deletedCount} individual sleeves (IDs: {string.Join(", ", deletedIds.Select(id => id.IntegerValue))})\n");
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: ❌ Batch delete failed: {ex.Message}, falling back to individual deletes\n");
                        
                        // Fallback to individual deletes if batch fails
                        foreach (var id in filteredToDelete)
                        {
                            try
                            {
                                // ✅ FINAL PROTECTION: Double-check before individual delete
                                if (protectedIds.Contains(id.IntegerValue))
                                {
                                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ CRITICAL PROTECTION: Skipping cluster sleeve {id.IntegerValue} in fallback delete!\n");
                                    continue;
                                }
                                
                                doc.Delete(id);
                                deletedCount++;
                            }
                            catch (Exception delEx)
                            {
                                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: ❌ Failed deleting sleeve {id.IntegerValue}: {delEx.Message}\n");
                                DebugLogger.Warning($"[ClusterCleanupService] Failed deleting sleeve {id.IntegerValue}: {delEx.Message}");
                            }
                        }
                    }
                }
                else
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Not in transaction, creating new transaction for batch delete of {filteredToDelete.Count} sleeves (protected {removedClusterSleeves} cluster sleeves)\n");
                    
                    using (var tx = new Transaction(doc, "Delete Individual Sleeves Within Clusters"))
                    {
                        tx.Start();
                        
                        // ✅ PERFORMANCE: Use batch delete instead of deleting one by one
                        try
                        {
                            var deletedIds = doc.Delete(filteredToDelete);
                            deletedCount = deletedIds.Count;
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: ✅ Batch deleted {deletedCount} individual sleeves (IDs: {string.Join(", ", deletedIds.Select(id => id.IntegerValue))})\n");
                        }
                        catch (Exception ex)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: ❌ Batch delete failed: {ex.Message}, falling back to individual deletes\n");
                            
                            // Fallback to individual deletes if batch fails
                            foreach (var id in filteredToDelete)
                            {
                                try
                                {
                                    // ✅ FINAL PROTECTION: Double-check before individual delete
                                    if (protectedIds.Contains(id.IntegerValue))
                                    {
                                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                            $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ CRITICAL PROTECTION: Skipping cluster sleeve {id.IntegerValue} in fallback delete!\n");
                                        continue;
                                    }
                                    
                                    doc.Delete(id);
                                    deletedCount++;
                                }
                                catch (Exception delEx)
                                {
                                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: ❌ Failed deleting sleeve {id.IntegerValue}: {delEx.Message}\n");
                                    DebugLogger.Warning($"[ClusterCleanupService] Failed deleting sleeve {id.IntegerValue}: {delEx.Message}");
                                }
                            }
                        }
                        
                        tx.Commit();
                    }
                }
                
                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: ✅ Completed cleanup: Deleted {deletedCount} individual sleeves\n");
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: ❌ ERROR during cleanup: {ex.Message}\n{ex.StackTrace}\n");
                DebugLogger.Error($"[ClusterCleanupService] Error during cleanup: {ex.Message}");
            }
            return deletedCount;
        }

        public void ResetClusterFlagsForDeletedSleeves(Document doc, string? xmlFilePath = null)
        {
            try
            {
                int resetCount = 0;
                var resetClashZones = new List<ClashZone>();
                var oldClusterInstanceIds = new Dictionary<Guid, int>(); // For potential future diagnostics
                var categories = new[] { "Ducts", "Pipes", "Cable Trays", "Duct Accessories", "Cable Tray Fittings" };

                using var context = new SleeveDbContext(doc, _ => { });
                var repository = new ClashZoneRepository(context, _ => { });

                foreach (var category in categories)
                {
                    var dbZones = repository.GetClashZonesByCategory(category)
                        ?.Where(z => z != null && z.IsClusterResolved && z.ClusterSleeveInstanceId > 0)
                        .ToList();
                    if (dbZones == null || dbZones.Count == 0) continue;
                    foreach (var clashZone in dbZones)
                    {
                        var clusterSleeveId = new ElementId(clashZone.ClusterSleeveInstanceId);
                        var clusterSleeve = doc.GetElement(clusterSleeveId);
                        if (clusterSleeve == null)
                        {
                            oldClusterInstanceIds[clashZone.Id] = clashZone.ClusterSleeveInstanceId;
                            clashZone.IsClusterResolved = false;
                            clashZone.ClusterSleeveInstanceId = -1;
                            resetClashZones.Add(clashZone);
                            resetCount++;
                        }
                    }
                }

                if (resetClashZones.Count > 0)
                {
                    // Direct DB update (repository has no bulk UpdateClashZones API)
                    using (var tx = context.Connection.BeginTransaction())
                    {
                        foreach (var z in resetClashZones)
                        {
                            try
                            {
                                using var cmd = context.Connection.CreateCommand();
                                cmd.Transaction = tx;
                                cmd.CommandText = @"UPDATE ClashZones
                                    SET IsClusterResolvedFlag = 0,
                                        ClusterInstanceId = -1,
                                        UpdatedAt = CURRENT_TIMESTAMP
                                    WHERE UPPER(ClashZoneGuid) = UPPER(@Guid)
                                      AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL";
                                cmd.Parameters.AddWithValue("@Guid", z.Id.ToString());
                                cmd.ExecuteNonQuery();
                            }
                            catch (Exception updEx)
                            {
                                DebugLogger.Warning($"[ClusterCleanupService] Failed DB flag reset for zone {z.Id}: {updEx.Message}");
                            }
                        }
                        tx.Commit();
                    }
                }
                DebugLogger.Info($"[ClusterCleanupService] Reset {resetCount} cluster flags for deleted sleeves.");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[ClusterCleanupService] Error resetting cluster flags: {ex.Message}");
            }
        }

        /// <summary>
        /// ✅ DB-ONLY CLEANUP: Uses bounding boxes from database (zero Revit queries).
        /// Finds individual sleeves within cluster bounding boxes using DB data only.
        /// Combines Stage 1 and Stage 2 cleanup into one operation.
        /// </summary>
        public int CleanupSleevesWithinClustersFromDatabase(Document doc, string targetCategory = null, List<int> clusterInstanceIds = null)
        {
            int deletedCount = 0;
            try
            {
                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 DB-ONLY CLEANUP: Starting (Category={targetCategory ?? "ALL"}, ClusterIds={clusterInstanceIds?.Count ?? 0})\n");

                using (var context = new SleeveDbContext(doc))
                {
                    var repo = new ClashZoneRepository(context, msg => SafeFileLogger.SafeAppendText("cluster_debug.log", $"[Repo] {msg}\n"));

                    // Step 1: Query individual sleeves with valid bounding boxes from DB
                    var individualSleeves = new List<(int SleeveInstanceId, double MinX, double MinY, double MinZ, double MaxX, double MaxY, double MaxZ, string MepCategory)>();
                    
                    using (var cmd = context.Connection.CreateCommand())
                    {
                        var query = @"
                            SELECT SleeveInstanceId, 
                                   SleeveBoundingBoxMinX, SleeveBoundingBoxMinY, SleeveBoundingBoxMinZ,
                                   SleeveBoundingBoxMaxX, SleeveBoundingBoxMaxY, SleeveBoundingBoxMaxZ,
                                   MepElementCategory
                            FROM ClashZones
                            WHERE SleeveInstanceId > 0
                              AND SleeveBoundingBoxMinX != 0.0 AND SleeveBoundingBoxMinY != 0.0 AND SleeveBoundingBoxMinZ != 0.0
                              AND SleeveBoundingBoxMaxX != 0.0 AND SleeveBoundingBoxMaxY != 0.0 AND SleeveBoundingBoxMaxZ != 0.0
                              AND (ClusterSleeveInstanceId = 0 OR ClusterSleeveInstanceId IS NULL)";

                        if (!string.IsNullOrEmpty(targetCategory))
                        {
                            query += " AND MepElementCategory = @targetCategory";
                            cmd.Parameters.AddWithValue("@targetCategory", targetCategory);
                        }

                        cmd.CommandText = query;

                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                var sleeveId = reader.GetInt32(reader.GetOrdinal("SleeveInstanceId"));
                                var minX = reader.GetDouble(reader.GetOrdinal("SleeveBoundingBoxMinX"));
                                var minY = reader.GetDouble(reader.GetOrdinal("SleeveBoundingBoxMinY"));
                                var minZ = reader.GetDouble(reader.GetOrdinal("SleeveBoundingBoxMinZ"));
                                var maxX = reader.GetDouble(reader.GetOrdinal("SleeveBoundingBoxMaxX"));
                                var maxY = reader.GetDouble(reader.GetOrdinal("SleeveBoundingBoxMaxY"));
                                var maxZ = reader.GetDouble(reader.GetOrdinal("SleeveBoundingBoxMaxZ"));
                                var mepCat = reader.IsDBNull(reader.GetOrdinal("MepElementCategory")) ? "" : reader.GetString(reader.GetOrdinal("MepElementCategory"));

                                // Validate bounding box
                                if (minX < maxX && minY < maxY && minZ < maxZ)
                                {
                                    individualSleeves.Add((sleeveId, minX, minY, minZ, maxX, maxY, maxZ, mepCat));
                                }
                            }
                        }
                    }

                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🧹 DB-ONLY CLEANUP: Found {individualSleeves.Count} individual sleeves with valid bounding boxes\n");

                    // Step 2: Query cluster sleeves with valid bounding boxes from DB
                    var clusterSleeves = new List<(int ClusterInstanceId, double MinX, double MinY, double MinZ, double MaxX, double MaxY, double MaxZ, string MepCategory)>();
                    
                    using (var cmd = context.Connection.CreateCommand())
                    {
                        var query = @"
                            SELECT DISTINCT ClusterSleeveInstanceId,
                                   ClusterSleeveBoundingBoxMinX, ClusterSleeveBoundingBoxMinY, ClusterSleeveBoundingBoxMinZ,
                                   ClusterSleeveBoundingBoxMaxX, ClusterSleeveBoundingBoxMaxY, ClusterSleeveBoundingBoxMaxZ,
                                   MepElementCategory
                            FROM ClashZones
                            WHERE ClusterSleeveInstanceId > 0
                              AND ClusterSleeveBoundingBoxMinX != 0.0 AND ClusterSleeveBoundingBoxMinY != 0.0 AND ClusterSleeveBoundingBoxMinZ != 0.0
                              AND ClusterSleeveBoundingBoxMaxX != 0.0 AND ClusterSleeveBoundingBoxMaxY != 0.0 AND ClusterSleeveBoundingBoxMaxZ != 0.0";

                        if (clusterInstanceIds != null && clusterInstanceIds.Count > 0)
                        {
                            var idPlaceholders = string.Join(",", clusterInstanceIds.Select((_, i) => $"@clusterId{i}"));
                            query += $" AND ClusterSleeveInstanceId IN ({idPlaceholders})";
                            for (int i = 0; i < clusterInstanceIds.Count; i++)
                            {
                                cmd.Parameters.AddWithValue($"@clusterId{i}", clusterInstanceIds[i]);
                            }
                        }

                        if (!string.IsNullOrEmpty(targetCategory))
                        {
                            query += " AND MepElementCategory = @targetCategory2";
                            cmd.Parameters.AddWithValue("@targetCategory2", targetCategory);
                        }

                        cmd.CommandText = query;

                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                var clusterId = reader.GetInt32(reader.GetOrdinal("ClusterSleeveInstanceId"));
                                var minX = reader.GetDouble(reader.GetOrdinal("ClusterSleeveBoundingBoxMinX"));
                                var minY = reader.GetDouble(reader.GetOrdinal("ClusterSleeveBoundingBoxMinY"));
                                var minZ = reader.GetDouble(reader.GetOrdinal("ClusterSleeveBoundingBoxMinZ"));
                                var maxX = reader.GetDouble(reader.GetOrdinal("ClusterSleeveBoundingBoxMaxX"));
                                var maxY = reader.GetDouble(reader.GetOrdinal("ClusterSleeveBoundingBoxMaxY"));
                                var maxZ = reader.GetDouble(reader.GetOrdinal("ClusterSleeveBoundingBoxMaxZ"));
                                var mepCat = reader.IsDBNull(reader.GetOrdinal("MepElementCategory")) ? "" : reader.GetString(reader.GetOrdinal("MepElementCategory"));

                                // Validate bounding box
                                if (minX < maxX && minY < maxY && minZ < maxZ)
                                {
                                    clusterSleeves.Add((clusterId, minX, minY, minZ, maxX, maxY, maxZ, mepCat));
                                }
                            }
                        }
                    }

                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🧹 DB-ONLY CLEANUP: Found {clusterSleeves.Count} cluster sleeves with valid bounding boxes\n");

                    if (clusterSleeves.Count == 0 || individualSleeves.Count == 0)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🧹 DB-ONLY CLEANUP: No clusters or no individual sleeves, skipping\n");
                        return 0;
                    }

                    // Step 3: Find overlaps (pure C# math, no Revit API)
                    var sleevesToDelete = new HashSet<int>();
                    var clusterIdsSet = new HashSet<int>(clusterSleeves.Select(c => c.ClusterInstanceId));

                    foreach (var individual in individualSleeves)
                    {
                        // Skip if this individual sleeve is actually a cluster sleeve
                        if (clusterIdsSet.Contains(individual.SleeveInstanceId))
                        {
                            continue;
                        }

                        // Check category compatibility
                        if (!string.IsNullOrEmpty(targetCategory) && !string.IsNullOrEmpty(individual.MepCategory))
                        {
                            bool isCrossCategory = false;
                            if (targetCategory.Contains("Pipe") && (individual.MepCategory.Contains("Duct") || individual.MepCategory.Contains("Tray") || individual.MepCategory.Contains("Accessory") || individual.MepCategory.Contains("Damper"))) isCrossCategory = true;
                            if (targetCategory.Contains("Duct") && (individual.MepCategory.Contains("Pipe") || individual.MepCategory.Contains("Tray") || individual.MepCategory.Contains("Accessory") || individual.MepCategory.Contains("Damper"))) isCrossCategory = true;
                            if (targetCategory.Contains("Tray") && !individual.MepCategory.Contains("Tray")) isCrossCategory = true;
                            
                            if (isCrossCategory)
                            {
                                continue; // Protect cross-category sleeves
                            }
                        }

                        // Check if individual sleeve bounding box overlaps with any cluster bounding box
                        foreach (var cluster in clusterSleeves)
                        {
                            // Bounding box overlap check: two boxes overlap if they intersect on all three axes
                            bool bboxOverlaps = individual.MaxX >= cluster.MinX && individual.MinX <= cluster.MaxX &&
                                               individual.MaxY >= cluster.MinY && individual.MinY <= cluster.MaxY &&
                                               individual.MaxZ >= cluster.MinZ && individual.MinZ <= cluster.MaxZ;

                            if (bboxOverlaps)
                            {
                                sleevesToDelete.Add(individual.SleeveInstanceId);
                                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] 🗑 OVERLAP: Individual sleeve {individual.SleeveInstanceId} overlaps with cluster {cluster.ClusterInstanceId}\n");
                                break; // Found overlap, no need to check other clusters
                            }
                        }
                    }

                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🧹 DB-ONLY CLEANUP: Found {sleevesToDelete.Count} sleeves to delete\n");

                    // Step 4: Delete sleeves in one batch operation
                    if (sleevesToDelete.Count > 0)
                    {
                        var elementIdsToDelete = sleevesToDelete.Select(id => new ElementId(id)).ToList();
                        
                        bool alreadyInTransaction = doc.IsModifiable;
                        if (alreadyInTransaction)
                        {
                            // We're inside an existing transaction - delete directly
                            var deletedIds = doc.Delete(elementIdsToDelete);
                            deletedCount = deletedIds.Count;
                        }
                        else
                        {
                            // No active transaction - create one for deletion
                            using (Transaction tx = new Transaction(doc, "Delete Individual Sleeves Within Clusters (DB-Only)"))
                            {
                                tx.Start();
                                var deletedIds = doc.Delete(elementIdsToDelete);
                                deletedCount = deletedIds.Count;
                                tx.Commit();
                            }
                        }

                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🧹 DB-ONLY CLEANUP: Deleted {deletedCount} individual sleeves\n");
                    }
                }

                return deletedCount;
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("placement_errors.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ DB-ONLY CLEANUP FAILED: {ex.Message}\n{ex.StackTrace}\n");
                return 0;
            }
        }
    }
}
