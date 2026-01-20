using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Electrical;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ClearanceProviders
{
    /// <summary>
    /// Factory for creating appropriate clearance providers based on MEP category
    /// Uses Factory pattern for cost-effective provider creation
    /// </summary>
    public class ClearanceProviderFactory
    {
        private readonly Dictionary<string, IClearanceProvider> _providers;
        
        public ClearanceProviderFactory()
        {
            _providers = new Dictionary<string, IClearanceProvider>
            {
                { "Ducts", new SleeveClearanceProvider() },
                { "Pipes", new SleeveClearanceProvider() },
                { "Cable Trays", new CableTrayClearanceProvider() },
                { "Fire Dampers", new FireDamperClearanceProvider() },
                { "Default", new SleeveClearanceProvider() }
            };
        }
        
        /// <summary>
        /// Get clearance provider for a specific category
        /// </summary>
        /// <param name="category">MEP category name</param>
        /// <returns>Appropriate clearance provider</returns>
        public IClearanceProvider GetProvider(string category)
        {
            return _providers.TryGetValue(category, out var provider) 
                ? provider 
                : _providers["Default"];
        }
        
        /// <summary>
        /// Get clearance provider for a MEP element by analyzing its type
        /// </summary>
        /// <param name="mepElement">The MEP element</param>
        /// <returns>Appropriate clearance provider</returns>
        public IClearanceProvider GetProvider(Element mepElement)
        {
            string category = GetCategoryFromElement(mepElement);
            return GetProvider(category);
        }
        
        /// <summary>
        /// Get all available categories
        /// </summary>
        /// <returns>List of supported categories</returns>
        public IEnumerable<string> GetAvailableCategories()
        {
            return _providers.Keys;
        }
        
        private string GetCategoryFromElement(Element mepElement)
        {
            return mepElement switch
            {
                Duct => "Ducts",
                Pipe => "Pipes",
                CableTray => "Cable Trays",
                FamilyInstance fi when fi.Symbol.Family.Name.Contains("Damper") => "Fire Dampers",
                _ => "Default"
            };
        }
    }
}
