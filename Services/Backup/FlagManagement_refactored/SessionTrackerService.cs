using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Services.Logging;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.FlagManagement
{
    /// <summary>
    /// Team F: SOLID-compliant session tracker (removes static state).
    /// 
    /// This service replaces static HashSet in FlagManager with instance-based tracking.
    /// SOLID: Single Responsibility - session tracking only.
    /// 
    /// Purpose: Protects cluster sleeves that were just placed in the current session
    /// from being deleted by DeleteSleeveForIntersectionPointChange() during refresh.
    /// </summary>
    public class SessionTrackerService : ISessionTracker
    {
        private readonly HashSet<int> _recentlyPlacedClusterSleeveIds;
        private readonly object _lock;
        private readonly ILogger _logger;
        
        /// <summary>
        /// Creates a new session tracker instance.
        /// </summary>
        /// <param name="logger">Optional logger for tracking operations</param>
        public SessionTrackerService(ILogger logger = null)
        {
            _recentlyPlacedClusterSleeveIds = new HashSet<int>();
            _lock = new object();
            _logger = logger ?? LoggerAdapter.Default;
        }
        
        /// <summary>
        /// Register a cluster sleeve as recently placed (protects from deletion).
        /// </summary>
        public void RegisterRecentlyPlacedClusterSleeve(int clusterSleeveId)
        {
            if (clusterSleeveId <= 0)
            {
                _logger.Warning($"Attempted to register invalid cluster sleeve ID: {clusterSleeveId}", "SessionTracker");
                return;
            }
            
            lock (_lock)
            {
                _recentlyPlacedClusterSleeveIds.Add(clusterSleeveId);
                _logger.Debug($"Registered recently placed cluster sleeve: {clusterSleeveId}", "SessionTracker");
            }
        }
        
        /// <summary>
        /// Clear recently placed cluster sleeves list.
        /// </summary>
        public void ClearRecentlyPlacedClusterSleeves()
        {
            lock (_lock)
            {
                int count = _recentlyPlacedClusterSleeveIds.Count;
                _recentlyPlacedClusterSleeveIds.Clear();
                _logger.Debug($"Cleared {count} recently placed cluster sleeves", "SessionTracker");
            }
        }
        
        /// <summary>
        /// Check if a cluster sleeve was recently placed.
        /// </summary>
        public bool IsRecentlyPlacedClusterSleeve(int clusterSleeveId)
        {
            if (clusterSleeveId <= 0)
                return false;
            
            lock (_lock)
            {
                return _recentlyPlacedClusterSleeveIds.Contains(clusterSleeveId);
            }
        }
    }
}

