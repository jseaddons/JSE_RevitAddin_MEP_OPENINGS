using System;
using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team I: Database operations for filters.
    /// SOLID: Single Responsibility - database operations only.
    /// </summary>
    public interface IFilterRepository
    {
        /// <summary>
        /// Ensures filter exists in database (creates if not exists).
        /// Returns FilterId if successful, -1 if failed.
        /// </summary>
        int EnsureFilter(string filterName, string category);
        
        /// <summary>
        /// Updates filter name in database.
        /// </summary>
        void UpdateFilterName(string oldName, string category, string newName);
        
        /// <summary>
        /// Deletes filter from database.
        /// </summary>
        void DeleteFilter(string filterName, string category);
        
        /// <summary>
        /// Gets all filters from database.
        /// </summary>
        List<FilterInfo> GetAllFilters();
        
        /// <summary>
        /// Gets filter ID by name and category.
        /// </summary>
        int GetFilterId(string filterName, string category);
        
        /// <summary>
        /// Saves filter UI state to database.
        /// </summary>
        void SaveFilterUIState(string filterName, string category, List<string> selectedHostCategories, OpeningSettings settings);
        
        /// <summary>
        /// Loads filter UI state from database.
        /// </summary>
        (List<string> selectedHostCategories, OpeningSettings settings) LoadFilterUIState(string filterName, string category);
    }
    
    /// <summary>
    /// Filter information from database.
    /// </summary>
    public class FilterInfo
    {
        public int FilterId { get; set; }
        public string FilterName { get; set; }
        public string Category { get; set; }
    }
}
