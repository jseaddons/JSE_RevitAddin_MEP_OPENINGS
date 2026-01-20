using System.Collections.Generic;
using System.Windows.Forms;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team J: Filter UI orchestration (coordinates CRUD + Persistence).
    /// SOLID: Orchestrator pattern - coordinates focused services.
    /// </summary>
    public interface IFilterUiOrchestrator
    {
        /// <summary>
        /// Creates a new filter and saves it (database + XML).
        /// </summary>
        void CreateNewFilter(ListBox filterListBox);
        
        /// <summary>
        /// Copies a filter and saves it.
        /// </summary>
        void CopyFilter(ListBox filterListBox);
        
        /// <summary>
        /// Renames a filter and updates persistence.
        /// </summary>
        void RenameFilter(ListBox filterListBox);
        
        /// <summary>
        /// Deletes a filter from persistence and UI.
        /// </summary>
        void DeleteFilter(ListBox filterListBox);
        
        /// <summary>
        /// Saves filter to persistence (database + XML).
        /// </summary>
        void SaveFilter(ListBox filterListBox);
        
        /// <summary>
        /// Auto-saves filter (no user prompt).
        /// </summary>
        void SaveFilterAuto(ListBox filterListBox);
        
        /// <summary>
        /// Loads all saved filters into UI.
        /// </summary>
        void LoadAllSavedFilters(ListBox filterListBox);
        
        /// <summary>
        /// Seeds default filters into UI.
        /// </summary>
        void SeedDefaultFilters(ListBox filterListBox, IEnumerable<string> defaultNames);
    }
}

