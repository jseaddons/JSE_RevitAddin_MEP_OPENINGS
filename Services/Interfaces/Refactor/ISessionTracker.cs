namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team F: Session tracking for recently placed cluster sleeves (removes static state).
    /// SOLID: Single Responsibility - session tracking only.
    /// 
    /// This interface abstracts session tracking, enabling:
    /// - Dependency injection for testability
    /// - SOLID compliance (SRP - single responsibility, removes static state)
    /// - Instance-based tracking (thread-safe, no static state)
    /// 
    /// Purpose: Protects cluster sleeves that were just placed in the current session
    /// from being deleted by DeleteSleeveForIntersectionPointChange() during refresh.
    /// </summary>
    public interface ISessionTracker
    {
        /// <summary>
        /// Register a cluster sleeve as recently placed (protects from deletion).
        /// This prevents the sleeve from being deleted during refresh operations.
        /// </summary>
        /// <param name="clusterSleeveId">The Element ID (integer value) of the cluster sleeve</param>
        void RegisterRecentlyPlacedClusterSleeve(int clusterSleeveId);
        
        /// <summary>
        /// Clear recently placed cluster sleeves list.
        /// Call this at the start of a new session or when needed.
        /// </summary>
        void ClearRecentlyPlacedClusterSleeves();
        
        /// <summary>
        /// Check if a cluster sleeve was recently placed.
        /// </summary>
        /// <param name="clusterSleeveId">The Element ID (integer value) of the cluster sleeve</param>
        /// <returns>True if the sleeve was recently placed, false otherwise</returns>
        bool IsRecentlyPlacedClusterSleeve(int clusterSleeveId);
    }
}

