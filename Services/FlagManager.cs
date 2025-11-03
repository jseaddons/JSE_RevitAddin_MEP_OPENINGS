using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    // ✅ DEPLOYMENT MODE: All logging wrapped - tested pattern for deployment
    // Pattern: if (!DeploymentConfiguration.DeploymentMode) { DebugLogger.Info(...); }
    /// <summary>
    /// Centralized flag management for clash zones.
    /// Eliminates redundant flag operations across multiple services.
    /// Uses GlobalIndexService for all Global XML operations.
    /// </summary>
    public class FlagManager
    {
        private readonly Document _document;
        
        public FlagManager(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }
        
        /// <summary>
        /// Syncs flags from Global XML (authoritative source) to Filter XML clash zones.
        /// Called during refresh to ensure Filter XML reflects the current state from Global XML.
        /// </summary>
        /// <param name="clashZones">List of clash zones to sync (from Filter XML)</param>
        /// <param name="category">MEP element category name</param>
        public void SyncFlagsFromGlobal(List<ClashZone> clashZones, string category)
        {
            if (clashZones == null || clashZones.Count == 0)
                return;
                
            if (string.IsNullOrWhiteSpace(category))
                return;
            
            try
            {
                var globalIndex = GlobalIndexService.LoadOrCreate(_document, category);
                
                foreach (var clashZone in clashZones)
                {
                    if (clashZone == null) continue;
                    
                    var globalEntry = globalIndex.Entries?.FirstOrDefault(e => 
                        string.Equals(e.Id, clashZone.Id.ToString(), StringComparison.OrdinalIgnoreCase));
                    
                    if (globalEntry != null)
                    {
                        // Sync flags FROM Global (authoritative) TO Filter XML
                        bool flagChanged = false;
                        
                        if (clashZone.IsResolved != globalEntry.IsResolved)
                        {
                            clashZone.IsResolved = globalEntry.IsResolved;
                            flagChanged = true;
                        }
                        
                        if (clashZone.IsClusterResolved != globalEntry.IsClusterResolved)
                        {
                            clashZone.IsClusterResolved = globalEntry.IsClusterResolved;
                            flagChanged = true;
                        }
                        
                        if (globalEntry.IsResolved && clashZone.SleeveInstanceId != globalEntry.SleeveInstanceId)
                        {
                            clashZone.SleeveInstanceId = globalEntry.SleeveInstanceId;
                            flagChanged = true;
                        }
                        
                        if (globalEntry.IsClusterResolved && clashZone.ClusterSleeveInstanceId != globalEntry.ClusterSleeveInstanceId)
                        {
                            clashZone.ClusterSleeveInstanceId = globalEntry.ClusterSleeveInstanceId;
                            flagChanged = true;
                        }
                        
                        if (flagChanged)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[FLAG-MANAGER] Synced ClashZone {clashZone.Id} from Global XML (IsResolved={globalEntry.IsResolved}, IsClusterResolved={globalEntry.IsClusterResolved})");
                        }
                    }
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[FLAG-MANAGER] Synced flags from Global XML for {clashZones.Count} clash zones in category '{category}'");
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[FLAG-MANAGER] Error syncing flags from Global XML for category '{category}': {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// Resets flags for clash zones where sleeves were deleted.
        /// Follows flag hierarchy: Cluster flags take precedence over individual flags.
        /// Checks Global XML first (authoritative), then checks Revit if needed.
        /// </summary>
        /// <param name="clashZones">List of clash zones to check</param>
        /// <param name="category">MEP element category name</param>
        public void ResetFlagsForDeletedSleeves(List<ClashZone> clashZones, string category)
        {
            if (clashZones == null || clashZones.Count == 0)
                return;
                
            if (string.IsNullOrWhiteSpace(category))
                return;
            
            try
            {
                var globalIndex = GlobalIndexService.LoadOrCreate(_document, category);
                var updates = new List<(Guid Id, bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterSleeveInstanceId)>();
                int resetCount = 0;
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[FLAG-MANAGER] ===== STARTING DELETED SLEEVE CHECK FOR {clashZones.Count} CLASH ZONES =====");
                
                foreach (var clashZone in clashZones)
                {
                    if (clashZone == null) continue;
                    
                    var globalEntry = globalIndex.Entries?.FirstOrDefault(e => 
                        string.Equals(e.Id, clashZone.Id.ToString(), StringComparison.OrdinalIgnoreCase));
                    
                    bool globalSaysClusterResolved = globalEntry?.IsClusterResolved ?? false;
                    bool globalSaysResolved = globalEntry?.IsResolved ?? false;
                    
                    // ✅ DEBUG: Log clash zone state before checking (deployment mode wrapped)
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[FLAG-MANAGER] Checking ClashZone {clashZone.Id}: IsClusterResolved={clashZone.IsClusterResolved}, IsResolved={clashZone.IsResolved}, ClusterSleeveId={clashZone.ClusterSleeveInstanceId}, SleeveId={clashZone.SleeveInstanceId}");
                        DebugLogger.Info($"[FLAG-MANAGER]   Global XML says: IsClusterResolved={globalSaysClusterResolved}, IsResolved={globalSaysResolved}, ClusterSleeveId={globalEntry?.ClusterSleeveInstanceId ?? -1}, SleeveId={globalEntry?.SleeveInstanceId ?? -1}");
                    }
                    
                    // Flag Hierarchy: Check cluster FIRST (cluster flags take precedence)
                    if (clashZone.IsClusterResolved)
                    {
                        // ✅ CRITICAL FIX: Always check Revit API FIRST (Global XML may be stale if sleeve was deleted)
                        // Don't trust Global XML blindly - verify the sleeve actually exists in Revit
                        // Use ClusterSleeveInstanceId from clashZone (synced from Global XML) or fallback to Global XML entry
                        int clusterSleeveId = clashZone.ClusterSleeveInstanceId > 0 
                            ? clashZone.ClusterSleeveInstanceId 
                            : (globalEntry?.ClusterSleeveInstanceId ?? -1);
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[FLAG-MANAGER]   → Checking cluster sleeve ID {clusterSleeveId} in Revit API...");
                        bool clusterSleeveExists = CheckClusterSleeveExists(clusterSleeveId);
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[FLAG-MANAGER]   → Revit API check result: {clusterSleeveExists}");
                        
                        if (!clusterSleeveExists)
                        {
                            // ✅ Cluster sleeve deleted → Reset ALL flags (regardless of Global XML state)
                            // Global XML may still say IsClusterResolved=true, but we verify Revit is authoritative
                            clashZone.IsClusterResolved = false;
                            clashZone.IsResolved = false;
                            clashZone.ClusterSleeveInstanceId = -1;
                            clashZone.SleeveInstanceId = -1;
                            clashZone.SleeveFamilyName = string.Empty;
                            clashZone.LastUpdated = DateTime.Now;
                            
                            updates.Add((clashZone.Id, false, false, -1, -1));
                            resetCount++;
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[FLAG-MANAGER] ✓ Reset ALL flags for ClashZone {clashZone.Id} - cluster sleeve {clusterSleeveId} NOT FOUND in Revit (Global XML was stale)");
                        }
                        else
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[FLAG-MANAGER] ✓ Verified: ClashZone {clashZone.Id} - cluster sleeve {clusterSleeveId} EXISTS in Revit");
                        }
                    }
                    // Then check individual (only if cluster flag is false)
                    else if (clashZone.IsResolved)
                    {
                        // ✅ CRITICAL FIX: Always check Revit API FIRST (Global XML may be stale if sleeve was deleted)
                        // Don't trust Global XML blindly - verify the sleeve actually exists in Revit
                        // Use SleeveInstanceId from clashZone (synced from Global XML) or fallback to Global XML entry
                        int individualSleeveId = clashZone.SleeveInstanceId > 0 
                            ? clashZone.SleeveInstanceId 
                            : (globalEntry?.SleeveInstanceId ?? -1);
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[FLAG-MANAGER]   → Checking individual sleeve ID {individualSleeveId} in Revit API...");
                        bool individualSleeveExists = CheckIndividualSleeveExists(individualSleeveId);
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[FLAG-MANAGER]   → Revit API check result: {individualSleeveExists}");
                        
                        if (!individualSleeveExists)
                        {
                            // ✅ Individual sleeve deleted → Reset individual flag (regardless of Global XML state)
                            // Global XML may still say IsResolved=true, but we verify Revit is authoritative
                            clashZone.IsResolved = false;
                            clashZone.SleeveInstanceId = -1;
                            clashZone.SleeveFamilyName = string.Empty;
                            clashZone.LastUpdated = DateTime.Now;
                            
                            updates.Add((clashZone.Id, false, clashZone.IsClusterResolved, -1, clashZone.ClusterSleeveInstanceId));
                            resetCount++;
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[FLAG-MANAGER] ✓ Reset individual flags for ClashZone {clashZone.Id} - individual sleeve {individualSleeveId} NOT FOUND in Revit (Global XML was stale)");
                        }
                        else
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[FLAG-MANAGER] ✓ Verified: ClashZone {clashZone.Id} - individual sleeve {individualSleeveId} EXISTS in Revit");
                        }
                    }
                    else
                    {
                        // Neither cluster nor individual flags are set - nothing to check
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[FLAG-MANAGER]   → Skipping ClashZone {clashZone.Id} - neither IsClusterResolved nor IsResolved is true");
                    }
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[FLAG-MANAGER] ===== COMPLETED DELETED SLEEVE CHECK: {resetCount} flags reset =====");
                
                // Save updated flags to Global XML using existing service
                if (updates.Count > 0)
                {
                    GlobalIndexService.UpsertFlagsWithIds(_document, category, updates);
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[FLAG-MANAGER] Updated Global XML for {updates.Count} clash zones with reset flags in category '{category}'");
                }
                
                if (resetCount > 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[FLAG-MANAGER] Reset flags for {resetCount} clash zones in category '{category}'");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[FLAG-MANAGER] Error resetting flags for deleted sleeves in category '{category}': {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// Updates flags after sleeve placement (individual or cluster).
        /// Updates both in-memory ClashZone object and Global XML.
        /// </summary>
        /// <param name="clashZone">The clash zone that had a sleeve placed</param>
        /// <param name="sleeveId">The Revit element ID of the placed sleeve</param>
        /// <param name="isCluster">True if this is a cluster sleeve, false for individual sleeve</param>
        /// <param name="category">MEP element category name</param>
        public void UpdateFlagsForPlacement(ClashZone clashZone, int sleeveId, bool isCluster, string category)
        {
            if (clashZone == null)
                throw new ArgumentNullException(nameof(clashZone));
                
            if (sleeveId <= 0)
                throw new ArgumentException("Sleeve ID must be greater than 0", nameof(sleeveId));
                
            if (string.IsNullOrWhiteSpace(category))
                throw new ArgumentException("Category cannot be null or empty", nameof(category));
            
            try
            {
                if (isCluster)
                {
                    // Cluster sleeve placed
                    clashZone.IsClusterResolved = true;
                    clashZone.ClusterSleeveInstanceId = sleeveId;
                    clashZone.IsResolved = true; // Keep individual flag true when clustered
                    clashZone.SleeveInstanceId = -1; // Clear individual ID (individual sleeve was deleted during clustering)
                }
                else
                {
                    // Individual sleeve placed
                    clashZone.IsResolved = true;
                    clashZone.SleeveInstanceId = sleeveId;
                    // Note: IsClusterResolved remains false for individual sleeves
                }
                
                // Update Global XML using existing service
                GlobalIndexService.UpsertFlagsWithIds(_document, category, new[] { 
                    (clashZone.Id, clashZone.IsResolved, clashZone.IsClusterResolved, 
                     clashZone.SleeveInstanceId, clashZone.ClusterSleeveInstanceId) 
                });
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[FLAG-MANAGER] Updated flags for ClashZone {clashZone.Id} after {(isCluster ? "cluster" : "individual")} sleeve placement (SleeveId={sleeveId})");
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[FLAG-MANAGER] Error updating flags for placement: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// Checks if a cluster sleeve exists in the active Revit document.
        /// </summary>
        /// <param name="sleeveId">The Revit element ID of the cluster sleeve</param>
        /// <returns>True if the sleeve exists, false otherwise</returns>
        private bool CheckClusterSleeveExists(int sleeveId)
        {
            if (sleeveId <= 0) return false;
            
            try
            {
                var element = _document.GetElement(new ElementId(sleeveId));
                return element != null;
            }
            catch
            {
                return false;
            }
        }
        
        /// <summary>
        /// Checks if an individual sleeve exists in the active Revit document.
        /// </summary>
        /// <param name="sleeveId">The Revit element ID of the individual sleeve</param>
        /// <returns>True if the sleeve exists, false otherwise</returns>
        private bool CheckIndividualSleeveExists(int sleeveId)
        {
            if (sleeveId <= 0) return false;
            
            try
            {
                var element = _document.GetElement(new ElementId(sleeveId));
                return element != null;
            }
            catch
            {
                return false;
            }
        }
    }
}

