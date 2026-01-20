using System.Collections.Generic;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces
{
    /// <summary>
    /// Interface for accessing UI state (selected filters, files, host types).
    /// Extracted from UniversalSleevePlacementCommand to adhere to SRP and DIP.
    /// </summary>
    public interface IUiStateProvider
    {
        /// <summary>
        /// Get selected host categories from UI.
        /// </summary>
        List<string> GetSelectedHostCategories();
        
        /// <summary>
        /// Get selected reference files from UI.
        /// </summary>
        List<string> GetSelectedReferenceFiles();
        
        /// <summary>
        /// Get selected host files from UI.
        /// </summary>
        List<string> GetSelectedHostFiles();
        
        /// <summary>
        /// Get selected filter items from UI.
        /// </summary>
        List<string> GetSelectedFilterItems();
    }
}

