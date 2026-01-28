using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Parameters.Processing
{
    // Dummy CombinedSleeveMarkService if it was also missing
    public class CombinedSleeveMarkService
    {
        public List<FamilyInstance> GetAllCombinedSleeves(Document doc)
        {
            return new List<FamilyInstance>();
        }

        public Dictionary<int, string> CalculateCombinedSleevePrefixes(Document doc, List<FamilyInstance> sleeves, string projectPrefix, bool remarkAll)
        {
            return new Dictionary<int, string>();
        }

        public (int, int) ApplyCombinedSleeveMarksBatch(Document doc, Dictionary<int, string> assignments)
        {
            return (0, 0);
        }
    }
}
