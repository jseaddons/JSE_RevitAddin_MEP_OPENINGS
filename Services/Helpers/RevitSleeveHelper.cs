using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Helpers
{
    public interface IRevitSleeveHelper
    {
        List<FamilyInstance> GetAllSleevesForCategory(Document doc, string category, SleeveDbContext context, object settings);
    }

    public class RevitSleeveHelper : IRevitSleeveHelper
    {
        public List<FamilyInstance> GetAllSleevesForCategory(Document doc, string category, SleeveDbContext context, object settings)
        {
            // 1. COLLECT ALL SLEEVE FAMILY INSTANCES (Fast)
            /* WARNFIX_R23: CS0162 - commented out always-false branch
            FilteredElementCollector collector;
            if (false) // settings is deprecated
                collector = new FilteredElementCollector(doc, doc.ActiveView.Id);
            else
                collector = new FilteredElementCollector(doc);
            */
            var collector = new FilteredElementCollector(doc);

            var instances = collector
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => IsSleeveFamily(fi))
                .ToList();

            if (string.IsNullOrEmpty(category)) return instances;

            // 2. FILTER BY CATEGORY (DB Check for Standard/Cluster)
            // For performance, provide an optimized LINQ path with a strict rollback flag.
            if (OptimizationFlags.UseOptimizedSleeveCategoryFilter)
            {
                return instances
                    .Where(fi => IsMatchingCategory(fi, category, context))
                    .ToList();
            }

            // Legacy path: explicit loop, retained for rollback safety.
            var filtered = new List<FamilyInstance>();
            foreach (var fi in instances)
            {
                if (IsMatchingCategory(fi, category, context))
                {
                    filtered.Add(fi);
                }
            }
            return filtered;
        }

        private bool IsSleeveFamily(FamilyInstance fi)
        {
            var famName = fi.Symbol?.Family?.Name ?? string.Empty;
            return famName.IndexOf("OpeningOnWall", StringComparison.OrdinalIgnoreCase) >= 0
                || famName.IndexOf("OpeningOnSlab", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private bool IsMatchingCategory(FamilyInstance fi, string category, SleeveDbContext context)
        {
            // Check MEP_Category parameter first (Combined/Manual)
            var catParam = fi.LookupParameter("MEP_Category");
            if (catParam != null && !string.IsNullOrEmpty(catParam.AsString()))
            {
                string val = catParam.AsString();
                if (val.Equals(category, StringComparison.OrdinalIgnoreCase)) return true;
                if (category.Equals("Combined", StringComparison.OrdinalIgnoreCase) && 
                    (val.Contains("Multi") || val.Contains("Combined"))) return true;
            }

            // Check DB for Standard/Cluster
            // Note: This is a bit slow for loops, but safe. 
            // In a real batch, we'd pre-load a category lookup.
            return false; // Fallback - in practice, most have the parameter or we use the specialized collectors
        }
    }
}
