using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Strategies;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public interface ICombinedSleeveMarkService
    {
        (int processedCount, int errorCount) ApplyMarks(Document doc, string projectPrefix, string numberFormat);
        List<FamilyInstance> GetAllCombinedSleeves(Document doc);
        Dictionary<ElementId, string> CalculateCombinedSleevePrefixes(Document doc, List<FamilyInstance> sleeves, string projectPrefix, bool remarkAll);
        (int processedCount, int errorCount) ApplyCombinedSleeveMarksBatch(Document doc, Dictionary<ElementId, string> prefixAssignments);
    }

    public class CombinedSleeveMarkService : ICombinedSleeveMarkService
    {
        public (int processedCount, int errorCount) ApplyMarks(Document doc, string projectPrefix, string numberFormat)
        {
            var combos = GetAllCombinedSleeves(doc);
            if (combos.Count == 0) return (0, 0);

            var prefixes = CalculateCombinedSleevePrefixes(doc, combos, projectPrefix, true);
            return ApplyCombinedSleeveMarksBatch(doc, prefixes);
        }

        public List<FamilyInstance> GetAllCombinedSleeves(Document doc)
        {
            try
            {
                using var context = new SleeveDbContext(doc);
                var repo = new CombinedSleeveRepository(context);
                var allCombos = repo.GetAllCombinedSleeves();

                return allCombos
                    .Select(c => doc.GetElement(new ElementId(c.CombinedInstanceId)) as FamilyInstance)
                    .Where(fi => fi != null)
                    .ToList();
            }
            catch { return new List<FamilyInstance>(); }
        }

        public Dictionary<ElementId, string> CalculateCombinedSleevePrefixes(Document doc, List<FamilyInstance> sleeves, string projectPrefix, bool remarkAll)
        {
            var assignments = new Dictionary<ElementId, string>();
            foreach (var sleeve in sleeves)
            {
                var p = sleeve.LookupParameter("MEP Mark") ?? sleeve.LookupParameter("Mark");
                if (!remarkAll && !string.IsNullOrEmpty(p?.AsString())) continue;

                string prefix = $"{projectPrefix}MEP"; 
                assignments[sleeve.Id] = prefix;
            }
            return assignments;
        }

        public (int processedCount, int errorCount) ApplyCombinedSleeveMarksBatch(Document doc, Dictionary<ElementId, string> prefixAssignments)
        {
            int processedCount = 0;
            int errorCount = 0;

            if (prefixAssignments.Count == 0) return (0, 0);

            try
            {
                // Note: Sequence logic for combined sleeves is often just "MEP1", "MEP2" etc.
                // For "Prefix Only" mode, it just sets the prefix.
                // For batch, we'll assign numbers if needed, but the command seems to handle numbering in Phase 2.
                
                foreach (var entry in prefixAssignments)
                {
                    try
                    {
                        var el = doc.GetElement(entry.Key);
                        var p = el?.LookupParameter("MEP Mark") ?? el?.LookupParameter("Mark");
                        if (p != null && !p.IsReadOnly)
                        {
                            p.Set(entry.Value);
                            processedCount++;
                        }
                    }
                    catch { errorCount++; }
                }
            }
            catch { errorCount++; }

            return (processedCount, errorCount);
        }
    }
}
