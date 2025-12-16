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
            
            using (var transaction = new Transaction(_doc, "Place Cross-Category Combined Sleeves"))
            {
                transaction.Start();
                
                try
                {
                    foreach (var group in proximityGroups)
                    {
                        try
                        {
                            var combinedSleeve = PlaceSingleCombinedSleeve(group, comboId, filterId);
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
                    _logger($"[CombinedSleevePlacement] Transaction committed: {placedCombinedSleeves.Count} combined sleeves placed");
                }
                catch (Exception ex)
                {
                    transaction.RollBack();
                    _logger($"[CombinedSleevePlacement] ❌ Transaction rolled back: {ex.Message}");
                    throw;
                }
            }
            
            return placedCombinedSleeves;
        }
        
        /// <summary>
        /// Places a single combined sleeve for a proximity group
        /// </summary>
        private CombinedSleeve PlaceSingleCombinedSleeve(
            ProximityGroup group,
            int comboId,
            int filterId)
        {
            // Calculate combined geometry (Agent B logic)
            var bbox = group.CalculateCombinedBoundingBox();
            var (width, height, depth) = group.CalculateCombinedDimensions();
            var placementPoint = group.CalculateCombinedPlacementPoint();
            
            _logger($"[CombinedSleevePlacement] Placing combined sleeve: {group.GetSummary()}");
            _logger($"[CombinedSleevePlacement]   Dimensions: W={width:F2}, H={height:F2}, D={depth:F2}");
            _logger($"[CombinedSleevePlacement]   Placement: ({placementPoint.X:F2}, {placementPoint.Y:F2}, {placementPoint.Z:F2})");
            
            // TODO: Place actual Revit family instance
            // For now, create a placeholder ElementId
            var placedInstanceId = 999999; // Placeholder - replace with actual Revit placement
            
            // Calculate rotation angle (use first sleeve's rotation)
            var rotationAngle = group.Sleeves.FirstOrDefault()?.RotationAngleDeg ?? 0.0;
            
            // Create combined sleeve data model
            var combinedSleeve = new CombinedSleeve
            {
                CombinedInstanceId = placedInstanceId,
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
            
            // Calculate and save corners
            CalculateAndSaveCorners(combinedSleeve, placementPoint, width, height, rotationAngle);
            
            // Save to database (Agent A repository)
            var combinedSleeveId = _repository.SaveCombinedSleeve(combinedSleeve);
            combinedSleeve.CombinedSleeveId = combinedSleeveId;
            
            // Mark constituents as resolved (Agent A repository)
            _repository.MarkConstituentsAsResolved(combinedSleeve.Constituents);
            
            _logger($"[CombinedSleevePlacement] ✅ Saved combined sleeve {combinedSleeveId} to database");
            
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
                    // Individual sleeve - get from source data
                    if (sleeve.SourceData is JSE_RevitAddin_MEP_OPENINGS.Models.ClashZone clashZone)
                    {
                        constituent.ClashZoneGuid = clashZone.Id;
                        // ClashZoneId will be looked up by repository if needed
                    }
                }
                else if (sleeve.Type == SleeveType.Cluster)
                {
                    // Cluster sleeve - get from source data
                    if (sleeve.SourceData is JSE_RevitAddin_MEP_OPENINGS.Data.Repositories.ClusterSleeveData cluster)
                    {
                        constituent.ClusterInstanceId = cluster.ClusterInstanceId;
                        // ClusterSleeveId will be looked up by repository if needed
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
