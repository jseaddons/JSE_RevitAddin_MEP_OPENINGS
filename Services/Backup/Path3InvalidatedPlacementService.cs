using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Strategies;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// ✅ PATH 3 INVALIDATED PLACEMENT: Distinct placement flow for invalidated zones
    /// Handles moved elements, deletion of affected sleeves, flag reset, sizing, placement, and clustering
    /// </summary>
    public class Path3InvalidatedPlacementService
    {
        private readonly Document _document;
        private readonly FlagManager _flagManager;
        
        public Path3InvalidatedPlacementService(Document document, FlagManager flagManager = null)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _flagManager = flagManager ?? new FlagManager(document);
        }
        
        /// <summary>
        /// ✅ PATH 3 INVALIDATED: Execute distinct placement flow for invalidated zones
        /// Flow: Detect moved → Delete affected sleeves → Reset flags → Calculate sizes → Place new sleeves → Recalculate clusters
        /// </summary>
        public Path3InvalidatedPlacementResult ExecutePlacement(
            List<ClashZone> invalidatedZones,
            string filterName,
            string category,
            OpeningConditions conditions,
            ISleevePlacementStrategy strategy,
            Dictionary<string, double> clearanceSettings)
        {
            var result = new Path3InvalidatedPlacementResult();
            
            if (invalidatedZones == null || invalidatedZones.Count == 0)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[PATH3-INVALIDATED] No invalidated zones to process");
                }
                return result;
            }
            
            try
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[PATH3-INVALIDATED] Starting placement for {invalidatedZones.Count} invalidated zones");
                }
                
                // Step 1: Delete affected sleeves (individual sleeves at old intersection points)
                var deletedSleeves = DeleteAffectedSleeves(invalidatedZones, category);
                result.DeletedSleeveCount = deletedSleeves.Count;
                
                // Step 2: Reset flags for deleted sleeves (IsResolved = true to prevent re-placement)
                ResetFlagsForDeletedSleeves(invalidatedZones, deletedSleeves, category);
                
                // Step 3: Calculate new sizes (apply clearance settings)
                CalculateNewSizes(invalidatedZones, filterName, conditions, strategy, clearanceSettings);
                
                // Step 4: Place new sleeves at new intersection points
                var placementResult = PlaceNewSleeves(invalidatedZones, filterName, category, conditions, strategy, clearanceSettings);
                result.PlacedCount = placementResult.PlacedCount;
                result.ErrorCount = placementResult.ErrorCount;
                
                // Step 5: Recalculate clusters (always calculate, geometry changed)
                // Note: Clustering will be handled separately by the orchestrator
                // This service just prepares the zones for placement
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[PATH3-INVALIDATED] ✅ Placement complete: {result.PlacedCount} placed, {result.DeletedSleeveCount} deleted, {result.ErrorCount} errors");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[PATH3-INVALIDATED] ❌ Error in placement: {ex.Message}\n{ex.StackTrace}");
                }
                result.ErrorCount++;
                throw;
            }
            
            return result;
        }
        
        /// <summary>
        /// Step 1: Delete affected placed sleeves (individual sleeves at old intersection points)
        /// </summary>
        private List<int> DeleteAffectedSleeves(List<ClashZone> invalidatedZones, string category)
        {
            var deletedSleeveIds = new List<int>();
            
            try
            {
                foreach (var zone in invalidatedZones)
                {
                    try
                    {
                        // Delete individual sleeve if it exists
                        if (zone.SleeveInstanceId > 0)
                        {
                            var sleeveId = new ElementId(zone.SleeveInstanceId);
                            var sleeve = _document.GetElement(sleeveId) as FamilyInstance;
                            
                            if (sleeve != null)
                            {
                                _document.Delete(sleeveId);
                                deletedSleeveIds.Add(zone.SleeveInstanceId);
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[PATH3-INVALIDATED] Deleted affected sleeve {zone.SleeveInstanceId} for zone {zone.Id}");
                                }
                            }
                        }
                        
                        // Delete cluster sleeve if it exists
                        if (zone.ClusterSleeveInstanceId > 0)
                        {
                            var clusterId = new ElementId(zone.ClusterSleeveInstanceId);
                            var clusterSleeve = _document.GetElement(clusterId) as FamilyInstance;
                            
                            if (clusterSleeve != null)
                            {
                                _document.Delete(clusterId);
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[PATH3-INVALIDATED] Deleted affected cluster sleeve {zone.ClusterSleeveInstanceId} for zone {zone.Id}");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[PATH3-INVALIDATED] Error deleting sleeve for zone {zone.Id}: {ex.Message}");
                        }
                        // Continue with next zone
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[PATH3-INVALIDATED] Error in DeleteAffectedSleeves: {ex.Message}");
                }
                throw;
            }
            
            return deletedSleeveIds;
        }
        
        /// <summary>
        /// Step 2: Reset flags for deleted sleeves (IsResolved = true to prevent re-placement)
        /// </summary>
        private void ResetFlagsForDeletedSleeves(
            List<ClashZone> invalidatedZones, 
            List<int> deletedSleeveIds, 
            string category)
        {
            try
            {
                using (var dbContext = new SleeveDbContext(_document))
                {
                    var clashZoneRepository = new ClashZoneRepository(dbContext);
                    
                    foreach (var zone in invalidatedZones)
                    {
                        try
                        {
                            // Reset flags for deleted sleeves
                            if (deletedSleeveIds.Contains(zone.SleeveInstanceId))
                            {
                                // Mark as resolved to prevent re-placement
                                zone.IsResolved = true;
                                zone.SleeveInstanceId = 0;
                                
                                // Update in database
                                clashZoneRepository.BatchUpdateFlags(new[]
                                {
                                    (zone.Id, 
                                     IsResolved: true, 
                                     IsClusterResolved: zone.IsClusterResolved, 
                                     SleeveInstanceId: 0, 
                                     ClusterInstanceId: zone.ClusterSleeveInstanceId,
                                     zone.MepElementIdValue, 
                                     zone.StructuralElementIdValue, 
                                     zone.IntersectionPointX, 
                                     zone.IntersectionPointY, 
                                     zone.IntersectionPointZ,
                                     OldSleeveInstanceId: zone.SleeveInstanceId,
                                     OldClusterInstanceId: zone.ClusterSleeveInstanceId,
                                     MarkedForClusterProcess: zone.MarkedForClusteringSleeveProcess, // ✅ EDGE CASE: Pass through if available
                                     AfterClusterSleeveId: zone.AfterClusterSleevePlacedSleeveInstanceId, // ✅ EDGE CASE: Pass through if available
                                     IsClusteredFlag: (bool?)null) // ✅ EDGE CASE: Deprecated, set to null
                                });
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[PATH3-INVALIDATED] Reset flags for deleted sleeve in zone {zone.Id}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Warning($"[PATH3-INVALIDATED] Error resetting flags for zone {zone.Id}: {ex.Message}");
                            }
                            // Continue with next zone
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[PATH3-INVALIDATED] Error in ResetFlagsForDeletedSleeves: {ex.Message}");
                }
                throw;
            }
        }
        
        /// <summary>
        /// Step 3: Calculate new sleeve sizes (apply clearance settings)
        /// </summary>
        private void CalculateNewSizes(
            List<ClashZone> invalidatedZones,
            string filterName,
            OpeningConditions conditions,
            ISleevePlacementStrategy strategy,
            Dictionary<string, double> clearanceSettings)
        {
            try
            {
                foreach (var zone in invalidatedZones)
                {
                    try
                    {
                        // ✅ PATH 3 INVALIDATED: Calculate new size using PATH 2 logic (full calculation)
                        // This ensures sizes are recalculated based on new intersection points and conditions
                        var placerService = new NewSleevePlacerService(
                            _document,
                            conditions,
                            strategy,
                            clearanceSettings,
                            null, // sleeveRepository
                            null, // zoneFilterService
                            null, // familyManager
                            _flagManager,
                            false, // isReplayPath
                            filterName ?? "Unknown");
                        
                        // Calculate size (this will update zone.SleeveWidth, SleeveHeight, SleeveDepth)
                        // Note: This is a simplified approach - in production, you might want to call
                        // a dedicated size calculation method instead of full placement service
                        // For now, we'll rely on the placement service to calculate sizes during placement
                    }
                    catch (Exception ex)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[PATH3-INVALIDATED] Error calculating size for zone {zone.Id}: {ex.Message}");
                        }
                        // Continue with next zone
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[PATH3-INVALIDATED] Error in CalculateNewSizes: {ex.Message}");
                }
                throw;
            }
        }
        
        /// <summary>
        /// Step 4: Place new sleeves at new intersection points (no flag check before placement)
        /// </summary>
        private PlacementResult PlaceNewSleeves(
            List<ClashZone> invalidatedZones,
            string filterName,
            string category,
            OpeningConditions conditions,
            ISleevePlacementStrategy strategy,
            Dictionary<string, double> clearanceSettings)
        {
            var result = new PlacementResult();
            
            try
            {
                // ✅ PATH 3 INVALIDATED: Use PATH 2 placement logic (sizing and placement, no flag check)
                var placerService = new NewSleevePlacerService(
                    _document,
                    conditions,
                    strategy,
                    clearanceSettings,
                    null, // sleeveRepository
                    null, // zoneFilterService
                    null, // familyManager
                    _flagManager,
                    false, // isReplayPath
                    filterName);
                
                var placementOutcome = placerService.PlaceAllSleevesInTransaction(invalidatedZones);
                
                result.PlacedCount = placementOutcome.PlacedCount;
                result.ErrorCount = placementOutcome.ErrorCount;
                
                // Update flags after placement (IsResolved = true, SleeveInstanceId)
                UpdateFlagsAfterPlacement(invalidatedZones, category);
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[PATH3-INVALIDATED] Error in PlaceNewSleeves: {ex.Message}\n{ex.StackTrace}");
                }
                result.ErrorCount++;
                throw;
            }
            
            return result;
        }
        
        /// <summary>
        /// Update flags after placement (IsResolved = true, SleeveInstanceId)
        /// </summary>
        private void UpdateFlagsAfterPlacement(List<ClashZone> invalidatedZones, string category)
        {
            try
            {
                using (var dbContext = new SleeveDbContext(_document))
                {
                    var clashZoneRepository = new ClashZoneRepository(dbContext);
                    
                    var updates = new List<(Guid ClashZoneId, bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterInstanceId, int MepElementId, int StructuralElementId, double IntersectionPointX, double IntersectionPointY, double IntersectionPointZ, int OldSleeveInstanceId, int OldClusterInstanceId, bool? MarkedForClusterProcess, int AfterClusterSleeveId, bool? IsClusteredFlag)>();
                    
                    foreach (var zone in invalidatedZones)
                    {
                        if (zone.SleeveInstanceId > 0)
                        {
                            updates.Add((
                                zone.Id,
                                IsResolved: true,
                                IsClusterResolved: zone.IsClusterResolved,
                                SleeveInstanceId: zone.SleeveInstanceId,
                                ClusterInstanceId: zone.ClusterSleeveInstanceId,
                                zone.MepElementIdValue,
                                zone.StructuralElementIdValue,
                                zone.IntersectionPointX,
                                zone.IntersectionPointY,
                                zone.IntersectionPointZ,
                                OldSleeveInstanceId: 0, // Was deleted, so old ID is 0
                                OldClusterInstanceId: zone.ClusterSleeveInstanceId,
                                MarkedForClusterProcess: zone.MarkedForClusteringSleeveProcess, // ✅ EDGE CASE: Pass through if available
                                AfterClusterSleeveId: zone.AfterClusterSleevePlacedSleeveInstanceId, // ✅ EDGE CASE: Pass through if available
                                IsClusteredFlag: (bool?)null)); // ✅ EDGE CASE: Deprecated, set to null
                        }
                    }
                    
                    if (updates.Count > 0)
                    {
                        clashZoneRepository.BatchUpdateFlags(updates);
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[PATH3-INVALIDATED] Updated flags for {updates.Count} placed sleeves");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[PATH3-INVALIDATED] Error updating flags: {ex.Message}");
                }
                throw;
            }
        }
        
        /// <summary>
        /// Result of PATH 3 invalidated placement operation
        /// </summary>
        public class Path3InvalidatedPlacementResult
        {
            public int PlacedCount { get; set; }
            public int DeletedSleeveCount { get; set; }
            public int ErrorCount { get; set; }
        }
        
        /// <summary>
        /// Result of placement operation
        /// </summary>
        private class PlacementResult
        {
            public int PlacedCount { get; set; }
            public int ErrorCount { get; set; }
        }
    }
}

