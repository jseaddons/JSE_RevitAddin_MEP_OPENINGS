using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Combined.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Geometry;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Combined
{
    /// <summary>
    /// Integration service that combines Agent A's repository layer with Agent B's proximity detection.
    /// Handles the complete workflow: detection → placement → persistence.
    /// </summary>
    public class CombinedSleevePlacementService
    {
        private readonly Document _doc;
        private readonly ICombinedSleeveRepository _repository;
        private readonly ICrossCategoryProximityService _proximityService;
        private readonly ISleeveCornerCalculationService _cornerService;
        private readonly Action<string> _logger;
        
        public CombinedSleevePlacementService(
            Document doc,
            ICombinedSleeveRepository repository,
            ICrossCategoryProximityService proximityService,
            ISleeveCornerCalculationService cornerService)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _proximityService = proximityService ?? throw new ArgumentNullException(nameof(proximityService));
            _cornerService = cornerService ?? throw new ArgumentNullException(nameof(cornerService));
            _logger = msg => DebugLogger.Info(msg);
        }
        
        /// <summary>
        /// Main entry point: Detects proximity groups and places combined sleeves.
        /// </summary>
        /// <param name="individualSleeves">Individual sleeves (ClashZones) from all categories</param>
        /// <param name="clusterSleeves">Cluster sleeves from all categories</param>
        /// <param name="comboId">File combination ID</param>
        /// <param name="filterId">Filter ID</param>
        /// <param name="proximityThreshold">Proximity threshold in feet (default: 1.0)</param>
        /// <returns>List of placed combined sleeves</returns>
        public List<CombinedSleeve> PlaceCombinedSleeves(
            List<ClashZone> individualSleeves,
            List<ClusterSleeveData> clusterSleeves,
            int comboId,
            int filterId,
            double proximityThreshold = 1.0)
        {
            _logger($"[CombinedSleevePlacement] Starting cross-category combined sleeve placement");
            _logger($"[CombinedSleevePlacement] Individual sleeves: {individualSleeves?.Count ?? 0}, Cluster sleeves: {clusterSleeves?.Count ?? 0}");
            _logger($"[CombinedSleevePlacement] Proximity threshold: {proximityThreshold:F2} ft");
            
            if ((individualSleeves == null || individualSleeves.Count == 0) &&
                (clusterSleeves == null || clusterSleeves.Count == 0))
            {
                _logger($"[CombinedSleevePlacement] No sleeves to process");
                return new List<CombinedSleeve>();
            }
            
            try
            {
                // Step 1: Convert to unified sleeves (Agent B abstraction)
                var unifiedSleeves = ConvertToUnifiedSleeves(individualSleeves, clusterSleeves);
                _logger($"[CombinedSleevePlacement] Converted {unifiedSleeves.Count} sleeves to unified format");
                
                // Step 2: Detect proximity groups (Agent B algorithm)
                var proximityGroups = _proximityService.DetectProximityGroups(unifiedSleeves, proximityThreshold);
                _logger($"[CombinedSleevePlacement] Detected {proximityGroups.Count} cross-category proximity groups");
                
                if (proximityGroups.Count == 0)
                {
                    _logger($"[CombinedSleevePlacement] No cross-category proximity groups found");
                    return new List<CombinedSleeve>();
                }
                
                // Step 3: Place combined sleeves in Revit
                var placedCombinedSleeves = PlaceProximityGroups(proximityGroups, comboId, filterId);
                _logger($"[CombinedSleevePlacement] ✅ Placed {placedCombinedSleeves.Count} combined sleeves");
                
                return placedCombinedSleeves;
            }
            catch (Exception ex)
            {
                _logger($"[CombinedSleevePlacement] ❌ Error: {ex.Message}");
                DebugLogger.Error($"[CombinedSleevePlacement] Exception: {ex}");
                throw;
            }
        }
        
        /// <summary>
        /// Converts individual and cluster sleeves to unified sleeve abstraction
        /// </summary>
        private List<UnifiedSleeve> ConvertToUnifiedSleeves(
            List<ClashZone> individualSleeves,
            List<ClusterSleeveData> clusterSleeves)
        {
            var unified = new List<UnifiedSleeve>();
            
            // Convert individual sleeves
            if (individualSleeves != null)
            {
                foreach (var sleeve in individualSleeves)
                {
                    try
                    {
                        unified.Add(UnifiedSleeve.FromClashZone(sleeve));
                    }
                    catch (Exception ex)
                    {
                        _logger($"[CombinedSleevePlacement] Warning: Failed to convert individual sleeve {sleeve.Id}: {ex.Message}");
                    }
                }
            }
            
            // Convert cluster sleeves
            if (clusterSleeves != null)
            {
                foreach (var cluster in clusterSleeves)
                {
                    try
                    {
                        unified.Add(UnifiedSleeve.FromClusterSleeve(cluster));
                    }
                    catch (Exception ex)
                    {
                        _logger($"[CombinedSleevePlacement] Warning: Failed to convert cluster sleeve {cluster.ClusterInstanceId}: {ex.Message}");
                    }
                }
            }
            
            return unified;
        }
        
        /// <summary>
        /// Places combined sleeves in Revit for each proximity group.
        /// Handles Revit Transaction, Database Persistence, and Cleanup.
        /// </summary>
        public List<CombinedSleeve> PlaceProximityGroups(
            List<ProximityGroup> proximityGroups,
            int comboId,
            int filterId)
        {
            var placedCombinedSleeves = new List<CombinedSleeve>();
            
            // CRITICAL: Separate Revit transaction from database operations
            // Per REVIT_TRANSACTION_MANAGEMENT_SAFE_PLAN.md:
            // "Never mix database operations with Revit transactions"
            
            // Step 1: Place in Revit (Revit transaction)
            using (var transaction = new Transaction(_doc, "Place Cross-Category Combined Sleeves"))
            {
                transaction.Start();
                
                try
                {
                    foreach (var group in proximityGroups)
                    {
                        try
                        {
                            // Only place in Revit, don't save to database yet
                            var combinedSleeve = PlaceSingleCombinedSleeveInRevit(group, comboId, filterId);
                            // ✅ Only add if placement succeeded (CombinedInstanceId > 0)
                            if (combinedSleeve != null && combinedSleeve.CombinedInstanceId > 0)
                            {
                                placedCombinedSleeves.Add(combinedSleeve);
                            }
                            else if (combinedSleeve != null)
                            {
                                _logger($"[CombinedSleevePlacement] ⚠️ Skipping failed placement (CombinedInstanceId={combinedSleeve.CombinedInstanceId})");
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger($"[CombinedSleevePlacement] ⚠️ Failed to place combined sleeve for group: {ex.Message}");
                        }
                    }
                    
                    _doc.Regenerate(); // ✅ REGEN: Requested by user ("regen revit one time")
                    
                    transaction.Commit();
                    _logger($"[CombinedSleevePlacement] Revit transaction committed: {placedCombinedSleeves.Count} combined sleeves placed");
                }
                catch (Exception ex)
                {
                    transaction.RollBack();
                    _logger($"[CombinedSleevePlacement] ❌ Revit transaction rolled back: {ex.Message}");
                    throw;
                }
            }
            
            // Step 2: Save to database (OUTSIDE Revit transaction)
            if (placedCombinedSleeves.Count > 0)
            {
                try
                {
                    // ✅ BATCH SAVE & FLAG UPDATE: Use single transaction for efficiency
                    // Filter out invalid ones first
                    var validSleeves = placedCombinedSleeves
                        .Where(cs => cs.CombinedInstanceId > 0)
                        .ToList();

                    if (validSleeves.Count < placedCombinedSleeves.Count)
                    {
                         _logger($"[CombinedSleevePlacement] ⚠️ Skipping {placedCombinedSleeves.Count - validSleeves.Count} invalid sleeves from DB save");
                    }
                    
                    if (validSleeves.Count > 0)
                    {
                        _repository.SaveCombinedSleevesBatch(validSleeves);
                        _logger($"[CombinedSleevePlacement] ✅ Batch saved {validSleeves.Count} combined sleeves to database");
                    }
                }
                catch (Exception ex)
                {
                    _logger($"[CombinedSleevePlacement] ⚠️ Database save failed: {ex.Message}");
                    // Don't throw - Revit placement succeeded, database is secondary
                }
            }

            // Step 3: OPTIONAL CLEANUP – remove constituent sleeves that are now covered by combined sleeves
            // NOTE: This uses CombinedSleeveConstituents as the single source of truth, instead of
            // geometric point-in-box checks. It is intentionally separate from cluster cleanup.
            try
            {
                CleanupConstituentSleevesForCombined(placedCombinedSleeves);
            }
            catch (Exception ex)
            {
                _logger($"[CombinedSleevePlacement] ⚠️ Combined cleanup failed: {ex.Message}");
                // Do not throw – placement and DB save already succeeded.
            }
            
            return placedCombinedSleeves;
        }
        
        /// <summary>
        /// Places a single combined sleeve for a proximity group (Revit operations only, no database)
        /// </summary>
        public CombinedSleeve PlaceSingleCombinedSleeveInRevit(
            ProximityGroup group,
            int comboId,
            int filterId)
        {
            // Calculate combined geometry (Agent B logic)
            var bbox = group.CalculateCombinedBoundingBox();
            var result = group.CalculateCombinedDimensions();
            var width = result.width;
            var height = result.height;
            var depth = result.depth;

            // ✅ FIX: Ensure minimum depth for visibility/validity
            // User reported "Extrusion is too thin" errors. Enforcing default ~100mm if too small.
            double minDepth = 0.35; // ~106mm
            if (depth < 0.05) // < 15mm (treat as zero/invalid)
            {
                 _logger($"[CombinedSleevePlacement] ⚠️ Depth too thin ({depth:F3} ft). Enforcing minimum depth of {minDepth:F2} ft (~106mm) for visibility.");
                 depth = minDepth;
            }
            var rotation = group.CalculateCombinedRotation();
            var placementPoint = group.CalculateCombinedPlacementPoint();
            
            _logger($"[CombinedSleevePlacement] Placing combined sleeve: {group.GetSummary()}");
            _logger($"[CombinedSleevePlacement] Placing combined sleeve: {group.GetSummary()}");
            _logger($"[CombinedSleevePlacement]   Dimensions: W={width:F2}, H={height:F2}, D={depth:F2}, Rot={rotation * 180/Math.PI:F0}°");

            _logger($"[CombinedSleevePlacement]   Placement: ({placementPoint.X:F2}, {placementPoint.Y:F2}, {placementPoint.Z:F2})");

            FamilyInstance placedInstance = null;

            try
            {
                // 1. Determine Family Name
                string familyName = "RectangularOpeningOnWall"; // Default
                string hostType = group.GetHostType();
                
                if (hostType != null && (hostType.Contains("Floor") || hostType.Contains("Slab")))
                {
                    familyName = "RectangularOpeningOnSlab";
                }

                // 2. Load Symbol
                FamilySymbol symbol = new FilteredElementCollector(_doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(x => x.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase) || 
                                         x.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase));
                                         
                if (symbol != null)
                {
                    if (!symbol.IsActive) symbol.Activate();

                    // 3. Create Instance
                    // Use NonStructural for openings
                    placedInstance = _doc.Create.NewFamilyInstance(placementPoint, symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                    
                    if (placedInstance != null)
                    {
                        // 3b. Apply Rotation if needed
                        if (Math.Abs(rotation) > 0.001)
                        {
                             // Rotate around Z axis at placement point
                             Line axis = Line.CreateBound(placementPoint, placementPoint + XYZ.BasisZ);
                             ElementTransformUtils.RotateElement(_doc, placedInstance.Id, axis, rotation);
                        }

                        // 4. Set Parameters
                        var pWidth = placedInstance.LookupParameter("Width");
                        var pHeight = placedInstance.LookupParameter("Height");
                        // Depth might be controlled by host or instance parameter
                        var pDepth = placedInstance.LookupParameter("Depth");
                        var pLength = placedInstance.LookupParameter("Length"); // Some families use Length for depth

                        if (pWidth != null) pWidth.Set(width);
                        if (pHeight != null) pHeight.Set(height);
                        if (pDepth != null) pDepth.Set(depth);
                        else if (pLength != null) pLength.Set(depth);
                        
                        // Set Comments
                        var pComments = placedInstance.LookupParameter("Comments");
                        if (pComments != null)
                        {
                            var summary = group.GetSummary();
                            // Truncate if too long
                            if (summary.Length > 200) summary = summary.Substring(0, 197) + "...";
                            pComments.Set(summary);
                        }

                        // 4b. Set Instance ID (Address user request: "setting the combined sleeve instance id")
                        // ✅ FIX: Case-insensitive lookup with space variations
                        Parameter pInstanceId = null;
                        foreach (Parameter p in placedInstance.Parameters)
                        {
                            var paramName = p.Definition.Name.Replace(" ", "").Replace("_", "").ToLower();
                            if (paramName == "combinedsleeveinstanceid" || 
                                paramName == "combinedinstanceid" ||
                                paramName == "instanceid" || 
                                paramName == "sleeveid")
                            {
                                pInstanceId = p;
                                break;
                            }
                        }
                                        
                        if (pInstanceId != null && !pInstanceId.IsReadOnly)
                        {
                            try
                            {
                                if (pInstanceId.StorageType == StorageType.ElementId)
                                {
                                    pInstanceId.Set(placedInstance.Id);
                                }
                                else if (pInstanceId.StorageType == StorageType.Integer)
                                {
                                    pInstanceId.Set(placedInstance.Id.IntegerValue);
                                }
                                else if (pInstanceId.StorageType == StorageType.String)
                                {
                                    pInstanceId.Set(placedInstance.Id.IntegerValue.ToString());
                                }
                                _logger($"[CombinedSleevePlacement] ✅ Set Instance ID parameter '{pInstanceId.Definition.Name}' = {placedInstance.Id.IntegerValue}");
                            }
                            catch (Exception paramEx)
                            {
                                _logger($"[CombinedSleevePlacement] ⚠️ Failed to set Instance ID parameter: {paramEx.Message}");
                            }
                        }
                        else
                        {
                            _logger($"[CombinedSleevePlacement] ⚠️ No CombinedInstanceId parameter found in opening family");
                        }
                        
                        // ✅ SCHEDULE LEVEL & ELEVATION FROM LEVEL: Set from first sleeve's MEP element level (matches cluster sleeve logic)
                        try
                        {
                            // Get first sleeve from proximity group to extract level information
                            var firstSleeve = group.Sleeves.FirstOrDefault();
                            if (firstSleeve != null)
                            {
                                string levelName = null;
                                
                                // Extract level name from first sleeve (Individual or Cluster)
                                if (firstSleeve.SourceData is ClashZone cz)
                                {
                                    levelName = cz.MepElementLevelName;
                                }
                                else if (firstSleeve.SourceData is ClusterSleeveData cluster)
                                {
                                    // For cluster sleeves, try to get level from first constituent
                                    // This is a simplified approach - ideally we'd query the ClashZone for the cluster
                                    levelName = null; // TODO: Implement cluster level extraction if needed
                                }
                                
                                if (!string.IsNullOrWhiteSpace(levelName))
                                {
                                    // Find level in document
                                    Level mepLevel = new FilteredElementCollector(_doc)
                                        .OfClass(typeof(Level))
                                        .Cast<Level>()
                                        .FirstOrDefault(l => string.Equals(l.Name, levelName, StringComparison.OrdinalIgnoreCase));
                                    
                                    if (mepLevel != null)
                                    {
                                        // Set Schedule Level parameter (try multiple variations)
                                        var scheduleLevelParam = placedInstance.LookupParameter("Schedule of Level")
                                                             ?? placedInstance.LookupParameter("Schedule Level")
                                                             ?? placedInstance.LookupParameter("ScheduleLevel")
                                                             ?? placedInstance.Symbol?.LookupParameter("Schedule of Level")
                                                             ?? placedInstance.Symbol?.LookupParameter("Schedule Level")
                                                             ?? placedInstance.Symbol?.LookupParameter("ScheduleLevel");
                                        
                                        if (scheduleLevelParam != null && !scheduleLevelParam.IsReadOnly)
                                        {
                                            if (scheduleLevelParam.StorageType == StorageType.ElementId)
                                            {
                                                scheduleLevelParam.Set(mepLevel.Id);
                                            }
                                            else if (scheduleLevelParam.StorageType == StorageType.String)
                                            {
                                                scheduleLevelParam.Set(mepLevel.Name);
                                            }
                                            
                                            _logger($"[CombinedSleevePlacement] ✅ Set Schedule Level to '{mepLevel.Name}' on combined sleeve {placedInstance.Id.IntegerValue}");
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception levelEx)
                        {
                            _logger($"[CombinedSleevePlacement] ⚠️ Error setting Schedule Level: {levelEx.Message}");
                        }
                        
                        // ✅ BOTTOM OF OPENING: Calculate and set "Bottom of Opening" AFTER Schedule Level is set
                        // Formula: Bottom of Opening = Elevation from Level - (Height / 2)
                        // Only for WALL families (not slabs/floors)
                        if (familyName.IndexOf("Wall", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            try
                            {
                                // Step 1: Read "Elevation from Level" from parameter (Revit calculates this after Schedule Level is set)
                                var elevationParam = placedInstance.LookupParameter("Elevation from Level");
                                double? elevationFromLevel = null;
                                
                                if (elevationParam != null && elevationParam.StorageType == StorageType.Double)
                                {
                                    elevationFromLevel = elevationParam.AsDouble();
                                }
                                
                                if (elevationFromLevel.HasValue)
                                {
                                    // Step 2: Calculate Bottom of Opening = Elevation from Level - (Height / 2)
                                    double bottomOfOpening = elevationFromLevel.Value - (height / 2.0);
                                    
                                    // Step 3: Set "Bottom of Opening" parameter (try multiple variations)
                                    var bottomParam = placedInstance.LookupParameter("Bottom Of Opening")
                                                   ?? placedInstance.LookupParameter("Bottom of Opening")
                                                   ?? placedInstance.LookupParameter("BottomOfOpening")
                                                   ?? placedInstance.Symbol?.LookupParameter("Bottom Of Opening")
                                                   ?? placedInstance.Symbol?.LookupParameter("Bottom of Opening")
                                                   ?? placedInstance.Symbol?.LookupParameter("BottomOfOpening");
                                    
                                    if (bottomParam != null && !bottomParam.IsReadOnly)
                                    {
                                        bottomParam.Set(bottomOfOpening);
                                        _logger($"[CombinedSleevePlacement] ✅ Set Bottom of Opening = {bottomOfOpening * 304.8:F1}mm " +
                                               $"(Elevation from Level = {elevationFromLevel.Value * 304.8:F1}mm, Height = {height * 304.8:F1}mm)");
                                    }
                                    else
                                    {
                                        _logger($"[CombinedSleevePlacement] ⚠️ 'Bottom of Opening' parameter not found or read-only");
                                    }
                                }
                                else
                                {
                                    _logger($"[CombinedSleevePlacement] ⚠️ 'Elevation from Level' parameter not available - skipping Bottom of Opening calculation");
                                }
                            }
                            catch (Exception bottomEx)
                            {
                                _logger($"[CombinedSleevePlacement] ⚠️ Error setting Bottom of Opening: {bottomEx.Message}");
                            }
                        }

                        // 5. AUTO-JOIN (CRITICAL FIX)
                        try
                        {
                            // Try to find a host to join with if not automatically hosted
                            Element host = placedInstance.Host;
                            
                            // If auto-hosting didn't work (e.g. creating in open space), try to find intersecting host
                            if (host == null)
                            {
                                // Simple proximity search for Wall/Floor at placement point
                                var potentialHosts = new FilteredElementCollector(_doc)
                                    .OfClass(typeof(HostObject))
                                    .WherePasses(new BoundingBoxIntersectsFilter(new Outline(placementPoint - new XYZ(0.5, 0.5, 0.5), placementPoint + new XYZ(0.5, 0.5, 0.5))))
                                    .Cast<HostObject>()
                                    .ToList();

                                if (potentialHosts.Count > 0)
                                {
                                    host = potentialHosts.OrderBy(h => h.Location is LocationCurve lc ? lc.Curve.Distance(placementPoint) : 100).FirstOrDefault();
                                }
                            }

                            if (host != null)
                            {
                                if (!JoinGeometryUtils.AreElementsJoined(_doc, host, placedInstance))
                                {
                                    JoinGeometryUtils.JoinGeometry(_doc, host, placedInstance);
                                    _logger($"[CombinedSleevePlacement] ✅ Joined combined sleeve {placedInstance.Id} with host {host.Id} ({host.Category?.Name})");
                                }
                            }
                            else
                            {
                                _logger($"[CombinedSleevePlacement] ⚠️ No host found to join for sleeve {placedInstance.Id}");
                            }

                            // 5b. Instance Void Cut (Alternative for families that use voids)
                            if (host != null)
                            {
                                try
                                {
                                    InstanceVoidCutUtils.AddInstanceVoidCut(_doc, host, placedInstance);
                                    _logger($"[CombinedSleevePlacement] ✅ Added Void Cut for sleeve {placedInstance.Id} on host {host.Id}");
                                }
                                catch
                                {
                                    // Ignore failures (e.g. not a void family, or already cut)
                                }
                            }
                        }
                        catch (Exception joinEx)
                        {
                            _logger($"[CombinedSleevePlacement] ⚠️ Auto-Join/Cut failed: {joinEx.Message}");
                        }
                    }
                }
                else
                {
                     _logger($"[CombinedSleevePlacement] ❌ Family symbol '{familyName}' not found!");
                }
            }
            catch (Exception ex)
            {
                _logger($"[CombinedSleevePlacement] ❌ Error creating Revit bundle: {ex.Message}");
            }
            
            // Calculate rotation angle (use first sleeve's rotation or group average)
            var rotationAngle = rotation * (180.0 / Math.PI);
            
            // Create combined sleeve data model
            var combinedSleeve = new CombinedSleeve
            {
                CombinedInstanceId = placedInstance?.Id.IntegerValue ?? -1, // Use actual ID or -1 if failed
                ComboId = comboId,
                FilterId = filterId,
                Categories = group.Categories,
                
                BoundingBoxMinX = bbox.Min.X,
                BoundingBoxMinY = bbox.Min.Y,
                BoundingBoxMinZ = bbox.Min.Z,
                BoundingBoxMaxX = bbox.Max.X,
                BoundingBoxMaxY = bbox.Max.Y,
                BoundingBoxMaxZ = bbox.Max.Z,
                
                CombinedWidth = width,
                CombinedHeight = height,
                CombinedDepth = depth,
                
                PlacementX = placementPoint.X,
                PlacementY = placementPoint.Y,
                PlacementZ = placementPoint.Z,
                RotationAngleDeg = rotationAngle,
                
                HostType = group.GetHostType(),
                HostOrientation = group.GetHostOrientation(),
                
                Constituents = CreateConstituents(group)
            };
            
            // Calculate and save corners (in-memory only, database save happens later)
            CalculateAndSaveCorners(combinedSleeve, placementPoint, width, height, rotationAngle);

            // ✅ CLEANUP: Delete original sleeves (Individual and Cluster) if placement was successful
            // Matches user request: "delete those sleeves that forms the combined sleeve"
            if (placedInstance != null && placedInstance.IsValidObject)
            {
                try 
                {
                    DeleteConstituents(_doc, group);
                }
                catch (Exception ex)
                {
                    _logger($"[CombinedSleevePlacement] ⚠️ Cleanup failed: {ex.Message}");
                }
            }
            
            return combinedSleeve;
        }

        /// <summary>
        /// Deletes constituent sleeves (individual + cluster) whose IDs are recorded in
        /// CombinedSleeveConstituents for the given combined sleeves.
        /// This mirrors cluster cleanup intent but uses explicit constituent mapping instead
        /// of geometry-based point-in-box.
        /// </summary>
        /// <remarks>
        /// - Only runs when OptimizationFlags.UseCombinedClustering is true.
        /// - Never deletes the combined sleeves themselves – only their recorded constituents.
        /// - Safe to call even if Constituents collection is empty or not loaded; it will
        ///   re-read from the repository by CombinedSleeveId.
        /// </remarks>
        private void CleanupConstituentSleevesForCombined(List<CombinedSleeve> placedCombinedSleeves)
        {
            if (!OptimizationFlags.UseCombinedClustering)
            {
                _logger("[CombinedSleevePlacement] 🧹 Combined cleanup skipped (UseCombinedClustering=false)");
                return;
            }

            if (placedCombinedSleeves == null || placedCombinedSleeves.Count == 0)
            {
                return;
            }

            // Collect all Revit element IDs to delete from constituents
            var elementIdsToDelete = new HashSet<int>();
            var individualClashZoneIds = new HashSet<int>();

            foreach (var cs in placedCombinedSleeves)
            {
                if (cs == null) continue;

                // Ensure we have constituen ts – reload from repo if needed
                var constituents = cs.Constituents;
                if (constituents == null || constituents.Count == 0)
                {
                    if (cs.CombinedSleeveId > 0)
                    {
                        constituents = _repository.GetConstituents(cs.CombinedSleeveId);
                    }
                }

                if (constituents == null || constituents.Count == 0)
                    continue;

                foreach (var c in constituents)
                {
                    if (c == null) continue;

                    // Individual sleeves: record ClashZoneIds – we'll resolve SleeveInstanceId via DB
                    if (c.Type == ConstituentType.Individual && c.ClashZoneId.HasValue)
                    {
                        if (c.ClashZoneId.Value > 0)
                        {
                            individualClashZoneIds.Add(c.ClashZoneId.Value);
                        }
                    }

                    // Cluster sleeves: we have ClusterInstanceId (Revit family instance id)
                    if (c.Type == ConstituentType.Cluster && c.ClusterInstanceId.HasValue)
                    {
                        var id = c.ClusterInstanceId.Value;
                        if (id > 0)
                        {
                            elementIdsToDelete.Add(id);
                        }
                    }
                }
            }

            // Resolve individual ClashZoneIds -> SleeveInstanceId (Revit element ids)
            if (individualClashZoneIds.Count > 0)
            {
                try
                {
                    using (var context = new SleeveDbContext(_doc))
                    {
                        using (var cmd = context.Connection.CreateCommand())
                        {
                            var idPlaceholders = string.Join(",", individualClashZoneIds.Select((_, i) => $"@cz{i}"));
                            cmd.CommandText = $@"
                                SELECT ClashZoneId, SleeveInstanceId 
                                FROM ClashZones 
                                WHERE SleeveInstanceId > 0 AND ClashZoneId IN ({idPlaceholders})";

                            int idx = 0;
                            foreach (var czId in individualClashZoneIds)
                            {
                                cmd.Parameters.AddWithValue($"@cz{idx++}", czId);
                            }

                            using (var reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    int sleeveId = reader.GetInt32(reader.GetOrdinal("SleeveInstanceId"));
                                    if (sleeveId > 0)
                                    {
                                        elementIdsToDelete.Add(sleeveId);
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger($"[CombinedSleevePlacement] ⚠️ Failed to resolve individual constituents for cleanup: {ex.Message}");
                }
            }

            if (elementIdsToDelete.Count == 0)
            {
                _logger("[CombinedSleevePlacement] 🧹 Combined cleanup: no constituent sleeves to delete");
                return;
            }

            _logger($"[CombinedSleevePlacement] 🧹 Combined cleanup: deleting {elementIdsToDelete.Count} constituent sleeves");

            using (var tx = new Transaction(_doc, "Combined Sleeves Cleanup"))
            {
                tx.Start();
                try
                {
                    var revitIds = elementIdsToDelete
                        .Where(id => id > 0)
                        .Select(id => new ElementId(id))
                        .ToList();

                    if (revitIds.Count > 0)
                    {
                        _doc.Delete(revitIds);
                    }

                    tx.Commit();
                }
                catch (Exception ex)
                {
                    _logger($"[CombinedSleevePlacement] ⚠️ Combined cleanup transaction failed: {ex.Message}");
                    tx.RollBack();
                }
            }
        }
        
        /// <summary>
        /// Creates constituent records from proximity group sleeves
        /// </summary>
        private List<SleeveConstituent> CreateConstituents(ProximityGroup group)
        {
            var constituents = new List<SleeveConstituent>();
            
            foreach (var sleeve in group.Sleeves)
            {
                var constituent = new SleeveConstituent
                {
                    Type = sleeve.Type == SleeveType.Individual ? ConstituentType.Individual : ConstituentType.Cluster,
                    Category = sleeve.Category
                };
                
                if (sleeve.Type == SleeveType.Individual)
                {
                    // Individual sleeve
                    if (sleeve.SourceData is ClashZone clashZone)
                    {
                        constituent.ClashZoneGuid = clashZone.Id;
                    }
                }
                else if (sleeve.Type == SleeveType.Cluster)
                {
                    // Cluster sleeve
                    if (sleeve.SourceData is ClusterSleeveData cluster)
                    {
                        constituent.ClusterInstanceId = cluster.ClusterInstanceId;
                    }
                    else if (sleeve.SourceData is ClashZone cz && cz.ClusterSleeveInstanceId > 0)
                    {
                        // Supports ClusterSleeveInfo -> ClashZone mapping
                        constituent.ClusterInstanceId = cz.ClusterSleeveInstanceId;
                    }
                }
                
                constituents.Add(constituent);
            }
            
            return constituents;
        }

        /// <summary>
        /// Deletes the constituent sleeves (Individual and Cluster) that formed the group.
        /// </summary>
        private void DeleteConstituents(Document doc, ProximityGroup group)
        {
            try
            {
                var idsToDelete = new List<ElementId>();
                _logger($"[CombinedSleeveCleanup] Reviewing {group.Sleeves.Count} constituent sleeves for cleanup...");

                // Collect IDs to delete
                foreach (var sleeve in group.Sleeves)
                {
                    if (sleeve.Type == SleeveType.Individual)
                    {
                        if (sleeve.SourceData is ClashZone cz)
                        {
                            // ✅ CRITICAL FIX: Check if zone is part of a cluster
                            // If IsClusterResolved=true, the original individual sleeve was already deleted
                            // and replaced by a cluster sleeve. We need to delete the cluster sleeve instead.
                            if (cz.IsClusterResolved && cz.ClusterSleeveInstanceId > 0)
                            {
                                // Zone is part of a cluster - delete the CLUSTER sleeve, not the old individual
                                // Avoid duplicates by checking if we already added this cluster ID
                                var clusterElementId = new ElementId(cz.ClusterSleeveInstanceId);
                                if (!idsToDelete.Contains(clusterElementId))
                                {
                                    idsToDelete.Add(clusterElementId);
                                    _logger($"[CombinedSleeveCleanup]   MARKED: Cluster Sleeve {cz.ClusterSleeveInstanceId} (for clustered zone GUID={cz.Id})");
                                }
                                else
                                {
                                    _logger($"[CombinedSleeveCleanup]   SKIP: Cluster Sleeve {cz.ClusterSleeveInstanceId} already marked (zone GUID={cz.Id})");
                                }
                            }
                            else if (cz.SleeveInstanceId > 0)
                            {
                                // Zone is NOT clustered - delete the individual sleeve
                                idsToDelete.Add(new ElementId(cz.SleeveInstanceId));
                                _logger($"[CombinedSleeveCleanup]   MARKED: Individual Sleeve {cz.SleeveInstanceId} (GUID={cz.Id})");
                            }
                            else
                            {
                                _logger($"[CombinedSleeveCleanup]   SKIP: Individual Sleeve (GUID={cz.Id}) has invalid SleeveInstanceId={cz.SleeveInstanceId}");
                            }
                        }
                        else
                        {
                            _logger($"[CombinedSleeveCleanup]   SKIP: Individual Sleeve {sleeve.Id} - SourceData is null or not ClashZone");
                        }
                    }
                    else if (sleeve.Type == SleeveType.Cluster)
                    {
                        // For clusters, Id property holds the Revit Element ID as string
                        // Also check SourceData for robustness
                        int cid = -1;
                        if (sleeve.SourceData is ClusterSleeveData csd && csd.ClusterInstanceId > 0)
                        {
                            cid = csd.ClusterInstanceId;
                            _logger($"[CombinedSleeveCleanup]   MARKED: Cluster Sleeve {cid} (from SourceData ClusterData)");
                        }
                        else if (sleeve.SourceData is ClashZone cz && cz.ClusterSleeveInstanceId > 0)
                        {
                            cid = cz.ClusterSleeveInstanceId;
                            _logger($"[CombinedSleeveCleanup]   MARKED: Cluster Sleeve {cid} (from SourceData ClashZone)");
                        }
                        else if (int.TryParse(sleeve.Id, out int parsedId) && parsedId > 0)
                        {
                            cid = parsedId;
                            _logger($"[CombinedSleeveCleanup]   MARKED: Cluster Sleeve {cid} (from ID parsing)");
                        }
                        else if (sleeve.Id.StartsWith("C_") && int.TryParse(sleeve.Id.Substring(2), out int parsedIdC) && parsedIdC > 0)
                        {
                             cid = parsedIdC;
                             _logger($"[CombinedSleeveCleanup]   MARKED: Cluster Sleeve {cid} (from ID parsing C_ prefix)");
                        }
                        else
                        {
                            _logger($"[CombinedSleeveCleanup]   SKIP: Cluster Sleeve {sleeve.Id} - Could not resolve Element ID");
                        }
                        
                        if (cid > 0)
                        {
                            idsToDelete.Add(new ElementId(cid));
                        }
                    }
                }

                _logger($"[CombinedSleeveCleanup] Found {idsToDelete.Count} potential sleeves to delete.");

                // Perform Deletion
                if (idsToDelete.Count > 0)
                {
                    // Verify elements exist before deleting to avoid exceptions
                    var validIds = new List<ElementId>();
                    foreach (var id in idsToDelete)
                    {
                        var elem = doc.GetElement(id);
                        if (elem != null && elem.IsValidObject)
                        {
                            validIds.Add(id);
                        }
                        else
                        {
                            _logger($"[CombinedSleeveCleanup]   ⚠️ Element {id} not found in document or invalid (already deleted?)");
                        }
                    }

                    if (validIds.Count > 0)
                    {
                        doc.Delete(validIds);
                        _logger($"[CombinedSleeveCleanup] ✅ SUCCESS: Deleted {validIds.Count} constituent sleeves.");
                    }
                    else
                    {
                        _logger($"[CombinedSleeveCleanup] ⚠️ No valid elements found to delete (all were missing/invalid).");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger($"[CombinedSleeveCleanup] ❌ CRITICAL ERROR in DeleteConstituents: {ex.Message}\n{ex.StackTrace}");
            }
        }


        
        /// <summary>
        /// Calculates and saves corner coordinates for combined sleeve
        /// </summary>
        private void CalculateAndSaveCorners(
            CombinedSleeve combinedSleeve,
            XYZ placementPoint,
            double width,
            double height,
            double rotationAngle)
        {
            try
            {
                // Calculate corners using corner service
                var cornersResult = _cornerService.CalculateCorners(placementPoint, width, height, rotationAngle);
                
                if (cornersResult.HasValue)
                {
                    var corners = cornersResult.Value;
                    
                    combinedSleeve.Corner1X = corners.corner1.X;
                    combinedSleeve.Corner1Y = corners.corner1.Y;
                    combinedSleeve.Corner1Z = corners.corner1.Z;
                    
                    combinedSleeve.Corner2X = corners.corner2.X;
                    combinedSleeve.Corner2Y = corners.corner2.Y;
                    combinedSleeve.Corner2Z = corners.corner2.Z;
                    
                    combinedSleeve.Corner3X = corners.corner3.X;
                    combinedSleeve.Corner3Y = corners.corner3.Y;
                    combinedSleeve.Corner3Z = corners.corner3.Z;
                    
                    combinedSleeve.Corner4X = corners.corner4.X;
                    combinedSleeve.Corner4Y = corners.corner4.Y;
                    combinedSleeve.Corner4Z = corners.corner4.Z;
                    
                    _logger($"[CombinedSleevePlacement] Calculated corners for combined sleeve");
                }
                else
                {
                    _logger($"[CombinedSleevePlacement] ⚠️ Corner calculation returned null");
                }
            }
            catch (Exception ex)
            {
                _logger($"[CombinedSleevePlacement] ⚠️ Failed to calculate corners: {ex.Message}");
            }
        }
    }
}
