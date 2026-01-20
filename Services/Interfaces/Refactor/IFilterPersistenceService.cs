using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team I: Filter persistence operations (database + XML).
    /// SOLID: Single Responsibility - persistence operations only.
    /// </summary>
    public interface IFilterPersistenceService
    {
        /// <summary>
        /// Saves filter to database (primary storage).
        /// Returns FilterId if successful, -1 if failed.
        /// </summary>
        int SaveFilterToDatabase(string filterName, string category, OpeningFilter filter);
        
        /// <summary>
        /// Batch save multiple filters in a SINGLE TRANSACTION.
        /// 5-10x faster than saving individually.
        /// Returns array of FilterIds for successfully saved filters.
        /// </summary>
        int[] SaveFiltersBatch(List<(string name, string category, OpeningFilter filter)> filters);
        
        /// <summary>
        /// Loads filter from database.
        /// Returns null if not found.
        /// </summary>
        OpeningFilter LoadFilterFromDatabase(string filterName, string category);
        
        /// <summary>
        /// Saves filter to XML file (backward compatibility).
        /// </summary>
        void SaveFilterToXmlFile(OpeningFilter filter, string filePath);
        
        /// <summary>
        /// Loads filter from XML file (backward compatibility).
        /// Returns null if file not found.
        /// </summary>
        OpeningFilter LoadFilterFromXmlFile(string filePath);
        
        /// <summary>
        /// Checks if filter is saved (exists in database or XML).
        /// </summary>
        bool IsFilterSaved(string filterName, string category = null);
        
        /// <summary>
        /// Gets all saved filter names from database and XML.
        /// </summary>
        List<string> GetAllSavedFilterNames();
    }
}
