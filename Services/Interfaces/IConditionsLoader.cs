using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces
{
    /// <summary>
    /// Interface for loading and saving opening conditions from XML files.
    /// Extracted from UniversalSleevePlacementCommand to adhere to SRP.
    /// </summary>
    public interface IConditionsLoader
    {
        /// <summary>
        /// Load conditions for a given filter name and category combination.
        /// </summary>
        OpeningConditions LoadConditions(string filterName, string category);
        
        /// <summary>
        /// Get the file path for conditions XML file.
        /// </summary>
        string GetConditionsFilePath(string key);
        
        /// <summary>
        /// Save conditions to XML file.
        /// </summary>
        void SaveConditions(OpeningConditions conditions, string key);
    }
}

