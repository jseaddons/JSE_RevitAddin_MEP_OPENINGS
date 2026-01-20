using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Utils
{
    public static class SleeveCleanupHelper
    {
        /// <summary>
        /// Deletes all family instances whose family name ends with 'OpeningOnWall' or 'OpeningOnSlab'
        /// and have Width, Height, or Depth set to 0 (in internal units).
        /// </summary>
        public static int DeleteZeroSizeSleeves(Document doc)
        {
            var toDelete = new List<ElementId>();

            var collector = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .WhereElementIsNotElementType()
                .Cast<FamilyInstance>()
                .Where(fi =>
                    fi.Symbol != null &&
                    fi.Symbol.Family != null &&
                    (fi.Symbol.Family.Name.EndsWith("OpeningOnWall", StringComparison.OrdinalIgnoreCase) ||
                     fi.Symbol.Family.Name.EndsWith("OpeningOnSlab", StringComparison.OrdinalIgnoreCase)));

            foreach (var fi in collector)
            {
                double width = fi.LookupParameter("Width")?.AsDouble() ?? -1;
                double height = fi.LookupParameter("Height")?.AsDouble() ?? -1;
                double depth = fi.LookupParameter("Depth")?.AsDouble() ?? -1;

                if (width == 0 || height == 0 || depth == 0)
                {
                    toDelete.Add(fi.Id);
                }
            }

            if (toDelete.Count > 0)
            {
                using (var tx = new Transaction(doc, "Delete Zero-Size Sleeves"))
                {
                    tx.Start();
                    doc.Delete(toDelete);
                    tx.Commit();
                }
            }

            return toDelete.Count;
        }
    public static int DeletePipeSleeves(Document doc)
        {
            var toDelete = new List<ElementId>();

            var collector = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .WhereElementIsNotElementType()
                .Cast<FamilyInstance>()
                .Where(fi =>
                    fi.Symbol != null &&
                    fi.Symbol.Family != null &&
                    (fi.Symbol.Family.Name.Contains("OpeningOnWall", StringComparison.OrdinalIgnoreCase) ||
                     fi.Symbol.Family.Name.Contains("OpeningOnSlab", StringComparison.OrdinalIgnoreCase)) &&
                    fi.Symbol.Name.StartsWith("PS#", StringComparison.OrdinalIgnoreCase));

            foreach (var fi in collector)
            {
                toDelete.Add(fi.Id);
            }

            if (toDelete.Count > 0)
            {
                using (var tx = new Transaction(doc, "Delete Pipe Sleeves"))
                {
                    tx.Start();
                    doc.Delete(toDelete);
                    tx.Commit();
                }
            }

            return toDelete.Count;
        }
    }
}
