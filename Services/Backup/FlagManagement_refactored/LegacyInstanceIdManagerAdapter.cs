using System;
using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.FlagManagement
{
    /// <summary>
    /// Adapter that exposes the legacy FlagManager instance-ID reset logic
    /// through the IInstanceIdManager abstraction expected by the refactored pipeline.
    /// </summary>
    internal sealed class LegacyInstanceIdManagerAdapter : IInstanceIdManager
    {
        private readonly FlagManager _legacyFlagManager;

        public LegacyInstanceIdManagerAdapter(FlagManager legacyFlagManager)
        {
            _legacyFlagManager = legacyFlagManager ?? throw new ArgumentNullException(nameof(legacyFlagManager));
        }

        public int ResetInstanceIdsForDeletedSleeves(
            List<string> categories,
            Dictionary<string, List<ClashZone>> clashZonesByCategory = null,
            string refreshLogName = null)
        {
            return _legacyFlagManager.ResetInstanceIdsForDeletedSleeves(categories, clashZonesByCategory, refreshLogName);
        }
    }
}


