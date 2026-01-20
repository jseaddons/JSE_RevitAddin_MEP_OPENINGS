using System;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service to store and retrieve current mark prefixes from ParameterServiceDialog
    /// This allows EmergencyMainDialog to access the prefixes set in ParameterServiceDialog
    /// </summary>
    public static class MarkPrefixService
    {
        private static MarkPrefixSettings _currentPrefixes = new MarkPrefixSettings();
        private static bool _hasPrefixes = false;

        /// <summary>
        /// Set the current mark prefixes from ParameterServiceDialog
        /// </summary>
        public static void SetCurrentPrefixes(MarkPrefixSettings prefixes)
        {
            _currentPrefixes = prefixes ?? new MarkPrefixSettings();
            _hasPrefixes = true;
            
                        if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[MarkPrefixService] Set prefixes - Project: '{_currentPrefixes.ProjectPrefix}', " +
                $"Duct: '{_currentPrefixes.DuctPrefix}', Pipe: '{_currentPrefixes.PipePrefix}', " +
                $"CableTray: '{_currentPrefixes.CableTrayPrefix}', Damper: '{_currentPrefixes.DamperPrefix}'");
        }

        /// <summary>
        /// Get the current mark prefixes for use by EmergencyMainDialog
        /// </summary>
        public static MarkPrefixSettings GetCurrentPrefixes()
        {
            if (!_hasPrefixes)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning("[MarkPrefixService] No prefixes set, returning defaults");
                return new MarkPrefixSettings();
            }

                        if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[MarkPrefixService] Retrieved prefixes - Project: '{_currentPrefixes.ProjectPrefix}', " +
                $"Duct: '{_currentPrefixes.DuctPrefix}', Pipe: '{_currentPrefixes.PipePrefix}', " +
                $"CableTray: '{_currentPrefixes.CableTrayPrefix}', Damper: '{_currentPrefixes.DamperPrefix}'");
            
            return _currentPrefixes;
        }

        /// <summary>
        /// Clear the stored prefixes (useful for cleanup)
        /// </summary>
        public static void ClearPrefixes()
        {
            _currentPrefixes = new MarkPrefixSettings();
            _hasPrefixes = false;
                        if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info("[MarkPrefixService] Cleared stored prefixes");
        }
    }
}
