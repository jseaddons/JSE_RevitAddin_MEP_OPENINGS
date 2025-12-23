using System;
using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Services.Logging;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.FlagManagement
{
    /// <summary>
    /// Helper class to manage session-base protection for recently placed cluster sleeves.
    /// This replaces the static methods previously found in the legacy FlagManager.
    /// </summary>
    public static class FlagManagerProtectionHelper
    {
        // ✅ SESSION TRACKING: Track recently placed cluster sleeves to prevent deletion
        // This protects cluster sleeves that were just placed in the current session
        // from being deleted by DeleteSleeveForIntersectionPointChange() during refresh
        private static readonly HashSet<int> _recentlyPlacedClusterSleeveIds = new HashSet<int>();
        private static readonly object _recentlyPlacedLock = new object();

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
                $"[{DateTime.Now:HH:mm:ss}] [FLAG-MANAGER-HELPER] ✅✅✅ REGISTERED recently placed cluster sleeve ID={clusterSleeveId} (protected from deletion)\n");
            
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[FLAG-MANAGER-HELPER] ✅ Registered recently placed cluster sleeve ID={clusterSleeveId} (protected from deletion)");
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
                DebugLogger.Info($"[FLAG-MANAGER-HELPER] Cleared recently placed cluster sleeves list");
            }
        }
        
        /// <summary>
        /// ✅ SESSION PROTECTION: Check if a cluster sleeve was recently placed.
        /// </summary>
        /// <param name="clusterSleeveId">The Element ID (integer value) of the cluster sleeve</param>
        /// <returns>True if the sleeve was recently placed, false otherwise</returns>
        public static bool IsRecentlyPlacedClusterSleeve(int clusterSleeveId)
        {
            if (clusterSleeveId <= 0) return false;
            
            lock (_recentlyPlacedLock)
            {
                return _recentlyPlacedClusterSleeveIds.Contains(clusterSleeveId);
            }
        }
    }
}
