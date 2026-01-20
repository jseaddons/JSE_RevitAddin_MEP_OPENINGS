using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team F: Interface for collecting sleeves from Revit document.
    /// 
    /// This interface abstracts Revit API operations for sleeve collection,
    /// enabling unit testing without requiring actual Revit document.
    /// 
    /// SOLID: Dependency Inversion Principle - depends on abstraction, not concrete Revit API.
    /// </summary>
    public interface ISleeveCollector
    {
        /// <summary>
        /// Collect all sleeves from the document, grouped by MEP category.
        /// 
        /// ✅ OPTIMIZATION: Collects all sleeves in one API call (batch collection).
        /// Returns a dictionary where key is MEP category name and value is HashSet of sleeve IDs.
        /// </summary>
        /// <param name="document">Revit document</param>
        /// <returns>Dictionary of category name to HashSet of sleeve IDs</returns>
        Dictionary<string, HashSet<int>> CollectSleevesByCategory(Document document);
        
        /// <summary>
        /// Collect sleeves for a specific category.
        /// 
        /// Used when only one category needs to be processed (avoids batch collection overhead).
        /// </summary>
        /// <param name="document">Revit document</param>
        /// <param name="category">MEP category name (e.g., "Ducts", "Pipes")</param>
        /// <returns>HashSet of sleeve IDs for the specified category</returns>
        HashSet<int> CollectSleevesForCategory(Document document, string category);
        
        /// <summary>
        /// Check if a specific sleeve ID exists in the document.
        /// </summary>
        /// <param name="document">Revit document</param>
        /// <param name="sleeveId">Sleeve element ID (integer value)</param>
        /// <returns>True if sleeve exists, false otherwise</returns>
        bool SleeveExists(Document document, int sleeveId);
    }
}

