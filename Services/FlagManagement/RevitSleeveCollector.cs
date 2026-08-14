using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.FlagManagement
{
    /// <summary>
    /// Team F: Revit implementation of ISleeveCollector.
    /// 
    /// Wraps Revit API operations for sleeve collection.
    /// This implementation uses FilteredElementCollector to collect sleeves from the document.
    /// </summary>
    public class RevitSleeveCollector : ISleeveCollector
    {
        /// <summary>
        /// Collect all sleeves from the document, grouped by MEP category.
        /// 
        /// ✅ OPTIMIZATION: Collects all sleeves in one API call (batch collection).
        /// </summary>
        public Dictionary<string, HashSet<int>> CollectSleevesByCategory(Document document)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));
            
            var sleevesByCategory = new Dictionary<string, HashSet<int>>(StringComparer.OrdinalIgnoreCase);
            
            var allSleeves = new FilteredElementCollector(document)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(s =>
                {
                    bool hasSleeveKeyword = s.Symbol?.FamilyName?.Contains("Sleeve", StringComparison.OrdinalIgnoreCase) == true ||
                                           s.Symbol?.FamilyName?.Contains("Opening", StringComparison.OrdinalIgnoreCase) == true;
                    string familyName = s.Symbol?.FamilyName ?? "";
                    bool isKnownFamily = familyName.Contains("CircularOpening", StringComparison.OrdinalIgnoreCase) ||
                                        familyName.Contains("RectangularOpening", StringComparison.OrdinalIgnoreCase);
                    return (s.Category?.Name == "Generic Models" || 
                            s.Category?.Name == "Structural Connections" ||
                            s.Category?.Name == "Duct Accessories" ||
                            s.Category?.Name == "Pipe Accessories" ||
                            s.Category?.Name == "Mechanical Equipment") &&
                           (hasSleeveKeyword || isKnownFamily);
                })
                .ToList();
            
            foreach (var sleeve in allSleeves)
            {
                var mepCategoryParam = sleeve.LookupParameter("MEP_Category");
                if (mepCategoryParam != null && !string.IsNullOrWhiteSpace(mepCategoryParam.AsString()))
                {
                    string sleeveCategory = mepCategoryParam.AsString().Trim();
                    if (!string.IsNullOrWhiteSpace(sleeveCategory))
                    {
                        if (!sleevesByCategory.ContainsKey(sleeveCategory))
                            sleevesByCategory[sleeveCategory] = new HashSet<int>();
                        
                        int sleeveId = sleeve.Id.GetIntegerValue();
                        sleevesByCategory[sleeveCategory].Add(sleeveId);
                    }
                }
            }
            
            return sleevesByCategory;
        }

        /// <summary>
        /// Collects all unique sleeve/opening family instance IDs in the document.
        /// ✅ PERFORMANCE: Returns a HashSet for O(1) existence checks during refresh.
        /// </summary>
        public HashSet<int> CollectAllSleeveIds(Document document)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            var allSleeveIds = new FilteredElementCollector(document)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(s =>
                {
                    bool hasSleeveKeyword = s.Symbol?.FamilyName?.Contains("Sleeve", StringComparison.OrdinalIgnoreCase) == true ||
                                           s.Symbol?.FamilyName?.Contains("Opening", StringComparison.OrdinalIgnoreCase) == true;
                    string familyName = s.Symbol?.FamilyName ?? "";
                    bool isKnownFamily = familyName.Contains("CircularOpening", StringComparison.OrdinalIgnoreCase) ||
                                        familyName.Contains("RectangularOpening", StringComparison.OrdinalIgnoreCase);
                    return (s.Category?.Name == "Generic Models" || 
                            s.Category?.Name == "Structural Connections" ||
                            s.Category?.Name == "Duct Accessories" ||
                            s.Category?.Name == "Pipe Accessories" ||
                            s.Category?.Name == "Mechanical Equipment") &&
                           (hasSleeveKeyword || isKnownFamily);
                })
                .Select(s => s.Id.GetIntegerValue())
                .ToList();

            return new HashSet<int>(allSleeveIds);
        }
        
        /// <summary>
        /// Collect sleeves for a specific category.
        /// </summary>
        public HashSet<int> CollectSleevesForCategory(Document document, string category)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));
            
            if (string.IsNullOrWhiteSpace(category))
                return new HashSet<int>();
            
            var categorySleeves = new FilteredElementCollector(document)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(s =>
                {
                    bool hasSleeveKeyword = s.Symbol?.FamilyName?.Contains("Sleeve", StringComparison.OrdinalIgnoreCase) == true ||
                                           s.Symbol?.FamilyName?.Contains("Opening", StringComparison.OrdinalIgnoreCase) == true;
                    string familyName = s.Symbol?.FamilyName ?? "";
                    bool isKnownFamily = familyName.Contains("CircularOpening", StringComparison.OrdinalIgnoreCase) ||
                                        familyName.Contains("RectangularOpening", StringComparison.OrdinalIgnoreCase);
                    if (!((s.Category?.Name == "Generic Models" || s.Category?.Name == "Structural Connections") &&
                           (hasSleeveKeyword || isKnownFamily)))
                        return false;
                    
                    var mepCategoryParam = s.LookupParameter("MEP_Category");
                    if (mepCategoryParam != null && !string.IsNullOrWhiteSpace(mepCategoryParam.AsString()))
                    {
                        string sleeveCategory = mepCategoryParam.AsString();
                        return string.Equals(sleeveCategory, category, StringComparison.OrdinalIgnoreCase);
                    }
                    return false;
                })
                .Select(s => s.Id.GetIntegerValue())
                .ToList();
            
            return new HashSet<int>(categorySleeves);
        }
        
        /// <summary>
        /// Check if a specific sleeve ID exists in the document.
        /// </summary>
        public bool SleeveExists(Document document, int sleeveId)
        {
            if (document == null || sleeveId <= 0)
                return false;
            
            try
            {
#if REVIT2024_OR_GREATER
                var element = document.GetElement(new ElementId((long)sleeveId));
#else
                var element = document.GetElement(new ElementId(sleeveId));
#endif
                if (element != null && element is FamilyInstance sleeve)
                {
                    bool isSleeve = (sleeve.Category?.Name == "Generic Models" || 
                                     sleeve.Category?.Name == "Structural Connections" ||
                                     sleeve.Category?.Name == "Duct Accessories" ||
                                     sleeve.Category?.Name == "Pipe Accessories" ||
                                     sleeve.Category?.Name == "Mechanical Equipment") &&
                                    (sleeve.Symbol?.FamilyName?.Contains("Sleeve", StringComparison.OrdinalIgnoreCase) == true ||
                                     sleeve.Symbol?.FamilyName?.Contains("Opening", StringComparison.OrdinalIgnoreCase) == true);
                    return isSleeve;
                }
            }
            catch
            {
                // Element doesn't exist or error occurred
            }
            
            return false;
        }
    }
}
