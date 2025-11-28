using System;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.FlagManagement
{
    /// <summary>
    /// Thin adapter that forwards session tracking calls to the legacy static FlagManager helpers.
    /// </summary>
    internal sealed class LegacySessionTrackerAdapter : ISessionTracker
    {
        public void RegisterRecentlyPlacedClusterSleeve(int clusterSleeveId)
        {
            FlagManager.RegisterRecentlyPlacedClusterSleeve(clusterSleeveId);
        }

        public void ClearRecentlyPlacedClusterSleeves()
        {
            FlagManager.ClearRecentlyPlacedClusterSleeves();
        }

        public bool IsRecentlyPlacedClusterSleeve(int clusterSleeveId)
        {
            return FlagManager.IsRecentlyPlacedClusterSleeve(clusterSleeveId);
        }
    }
}


