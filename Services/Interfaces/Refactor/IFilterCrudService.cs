using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team J: Filter CRUD operations.
    /// SOLID: Single Responsibility - CRUD operations only.
    /// </summary>
    public interface IFilterCrudService
    {
        /// <summary>
        /// Creates a new filter from current UI state.
        /// </summary>
        OpeningFilter CreateFilter(string filterName);
        
        /// <summary>
        /// Copies an existing filter.
        /// </summary>
        OpeningFilter CopyFilter(OpeningFilter sourceFilter, string newName);
        
        /// <summary>
        /// Validates filter before save.
        /// Returns (isValid, errorMessage) tuple.
        /// </summary>
        (bool isValid, string errorMessage) ValidateFilter(OpeningFilter filter);
    }
}

