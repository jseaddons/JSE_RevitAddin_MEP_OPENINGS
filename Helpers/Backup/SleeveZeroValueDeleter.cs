using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    public static class SleeveZeroValueDeleter
    {
        /// <summary>
        /// Deletes all FamilyInstances whose family name ends with 'OnWall' or 'OnSlab' and have zero width, height, or depth.
        /// Returns the number of deleted elements.
        /// </summary>
        public static int DeleteZeroValueSleeves(Document doc)
        {
            int deletedCount = 0;
            var sleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Family.Name.EndsWith("OnWall") || fi.Symbol.Family.Name.EndsWith("OnSlab"))
                .ToList();

            var toDelete = sleeves.Where(sleeve =>
            {
                double width = sleeve.LookupParameter("Width")?.AsDouble() ?? 0.0;
                double height = sleeve.LookupParameter("Height")?.AsDouble() ?? 0.0;
                double depth = sleeve.LookupParameter("Depth")?.AsDouble() ?? 0.0;
                return width == 0.0 || height == 0.0 || depth == 0.0;
            }).ToList();

            if (toDelete.Any())
            {
                using (var tx = new Transaction(doc, "Delete Zero Value Sleeves"))
                {
                    tx.Start();
                    foreach (var sleeve in toDelete)
                    {
                        doc.Delete(sleeve.Id);
                        deletedCount++;
                    }
                    tx.Commit();
                }
            }
            return deletedCount;
        }
    }
}
