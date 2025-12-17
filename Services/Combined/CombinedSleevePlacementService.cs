using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
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
            List<JSE_RevitAddin_MEP_OPENINGS.Models.ClashZone> individualSleeves,
            List<JSE_RevitAddin_MEP_OPENINGS.Data.Repositories.ClusterSleeveData> clusterSleeves,
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
                var placedCombinedSleeves = PlaceCombinedSleevesInRevit(proximityGroups, comboId, filterId);
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
            List<JSE_RevitAddin_MEP_OPENINGS.Models.ClashZone> individualSleeves,
            List<JSE_RevitAddin_MEP_OPENINGS.Data.Repositories.ClusterSleeveData> clusterSleeves)
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
        /// Places combined sleeves in Revit for each proximity group
        /// </summary>
        private List<CombinedSleeve> PlaceCombinedSleevesInRevit(
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
                            if (combinedSleeve != null)
                            {
                                placedCombinedSleeves.Add(combinedSleeve);
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger($"[CombinedSleevePlacement] ⚠️ Failed to place combined sleeve for group: {ex.Message}");
                        }
                    }
                    
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
                    foreach (var combinedSleeve in placedCombinedSleeves)
                    {
                        // Save to database (Agent A repository)
                        var combinedSleeveId = _repository.SaveCombinedSleeve(combinedSleeve);
                        combinedSleeve.CombinedSleeveId = combinedSleeveId;
                        
                        // Mark constituents as resolved (Agent A repository)
                        _repository.MarkConstituentsAsResolved(combinedSleeve.Constituents);
                        
                        _logger($"[CombinedSleevePlacement] ✅ Saved combined sleeve {combinedSleeveId} to database");
                    }
                }
                catch (Exception ex)
                {
                    _logger($"[CombinedSleevePlacement] ⚠️ Database save failed: {ex.Message}");
                    // Don't throw - Revit placement succeeded, database is secondary
                }
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
            var placementPoint = group.CalculateCombinedPlacementPoint();
            
            _logger($"[CombinedSleevePlacement] Placing combined sleeve: {group.GetSummary()}");
            _logger($"[CombinedSleevePlacement]   Dimensions: W={width:F2}, H={height:F2}, D={depth:F2}");
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
            var rotationAngle = group.Sleeves.FirstOrDefault()?.RotationAngleDeg ?? 0.0;
            
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
            
            return combinedSleeve;
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
                    if (sleeve.SourceData is JSE_RevitAddin_MEP_OPENINGS.Models.ClashZone clashZone)
                    {
                        constituent.ClashZoneGuid = clashZone.Id;
                    }
                }
                else if (sleeve.Type == SleeveType.Cluster)
                {
                    // Cluster sleeve
                    if (sleeve.SourceData is JSE_RevitAddin_MEP_OPENINGS.Data.Repositories.ClusterSleeveData cluster)
                    {
                        constituent.ClusterInstanceId = cluster.ClusterInstanceId;
                    }
                }
                
                constituents.Add(constituent);
            }
            
            return constituents;
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
