using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    // ✅ DEPLOYMENT MODE: All logging wrapped - tested pattern for deployment
    // Pattern: if (!DeploymentConfiguration.DeploymentMode) { DebugLogger.Info(...); }
    /// <summary>
    /// Centralized flag management for clash zones.
    /// Eliminates redundant flag operations across multiple services.
    /// Uses SQLite Database for all flag operations.
    /// </summary>
    public class FlagManager
    {
        private readonly Document _document;
        
        // ✅ SESSION TRACKING: Track recently placed cluster sleeves to prevent deletion
        // This protects cluster sleeves that were just placed in the current session
        // from being deleted by DeleteSleeveForIntersectionPointChange() during refresh
        private static readonly HashSet<int> _recentlyPlacedClusterSleeveIds = new HashSet<int>();
        private static readonly object _recentlyPlacedLock = new object();
        
        public FlagManager(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }
        
        /// <summary>
        /// ✅ OOP HELPER: Gets the section box from the active 3D view.
        /// Returns null if no 3D view is active or no section box is enabled.
        /// </summary>
        /// <summary>
        /// ✅ SESSION PROTECTION: Register a cluster sleeve as recently placed.
        /// This prevents it from being deleted by DeleteSleeveForIntersectionPointChange() during refresh.
        /// </summary>
        /// <param name="clusterSleeveId">The Element ID (integer value) of the cluster sleeve</param>
        public static void RegisterRecentlyPlacedClusterSleeve(int clusterSleeveId)
        {
            if (clusterSleeveId <= 0) return;
            
            lock (_recentlyPlacedLock)
            {
                _recentlyPlacedClusterSleeveIds.Add(clusterSleeveId);
            }
            
            SafeFileLogger.SafeAppendText("cluster_debug.log",
                $"[{DateTime.Now:HH:mm:ss}] [FLAG-MANAGER] ✅✅✅ REGISTERED recently placed cluster sleeve ID={clusterSleeveId} (protected from deletion)\n");
            
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[FLAG-MANAGER] ✅ Registered recently placed cluster sleeve ID={clusterSleeveId} (protected from deletion)");
            }
        }
        
        /// <summary>
        /// ✅ SESSION PROTECTION: Clear the list of recently placed cluster sleeves.
        /// Call this at the start of a new session or when needed.
        /// </summary>
        public static void ClearRecentlyPlacedClusterSleeves()
        {
            lock (_recentlyPlacedLock)
            {
                _recentlyPlacedClusterSleeveIds.Clear();
            }
            
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[FLAG-MANAGER] Cleared recently placed cluster sleeves list");
            }
        }
        
        /// <summary>
        /// ✅ SESSION PROTECTION: Check if a cluster sleeve was recently placed.
        /// </summary>
        /// <param name="clusterSleeveId">The Element ID (integer value) of the cluster sleeve</param>
        /// <returns>True if the sleeve was recently placed, false otherwise</returns>
        private static bool IsRecentlyPlacedClusterSleeve(int clusterSleeveId)
        {
            if (clusterSleeveId <= 0) return false;
            
            lock (_recentlyPlacedLock)
            {
                return _recentlyPlacedClusterSleeveIds.Contains(clusterSleeveId);
            }
        }
        
        /// <summary>
        /// ✅ DATABASE-FIRST: Resets instance IDs for deleted sleeves (DB first, then XML).
        /// Sets SleeveInstanceId and ClusterSleeveInstanceId to -1 when sleeves are deleted.
        /// </summary>
        /// <param name="categories">List of categories to check</param>
        /// <param name="clashZonesByCategory">Optional dictionary of clash zones by category</param>
        /// <param name="refreshLogName">Optional refresh log file name</param>
        /// <returns>Number of instance IDs reset</returns>
        // ...existing code...
    }
}
