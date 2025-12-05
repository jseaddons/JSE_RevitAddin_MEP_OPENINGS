using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Refactored
{
    /// <summary>
    /// ✅ SOLID REFACTORED: Service for accessing UI state (selected filters, files, host types).
    /// Extracted from UniversalSleevePlacementCommand to adhere to SRP and DIP.
    /// Wraps static FilterUiStateProvider calls.
    /// </summary>
    public class UiStateProviderService : IUiStateProvider
    {
        public List<string> GetSelectedHostCategories()
        {
            return FilterUiStateProvider.GetSelectedHostCategories?.Invoke() ?? new List<string>();
        }

        public List<string> GetSelectedReferenceFiles()
        {
            return FilterUiStateProvider.GetSelectedReferenceFiles?.Invoke() ?? new List<string>();
        }

        public List<string> GetSelectedHostFiles()
        {
            return FilterUiStateProvider.GetSelectedHostFiles?.Invoke() ?? new List<string>();
        }

        public List<string> GetSelectedFilterItems()
        {
            return FilterUiStateProvider.GetSelectedFilterItems?.Invoke() ?? new List<string>();
        }
    }
}

