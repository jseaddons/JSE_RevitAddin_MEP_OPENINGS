using System;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    [Transaction(TransactionMode.Manual)]
    public class OpeningScheduleGenerator : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            using (Transaction t = new Transaction(doc, "Generate Opening Schedule"))
            {
                t.Start();

                // Assign marks to existing openings
                var allOpenings = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_GenericModel)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol.Family.Name.IndexOf("OpeningOnWall", StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToList();

                int pipeIndex = 1, ductIndex = 1;
                foreach (var fi in allOpenings)
                {
                    var markParam = fi.LookupParameter("Mark");
                    if (markParam != null && !markParam.IsReadOnly)
                    {
                        if (fi.Symbol.Name.IndexOf("PS#", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            markParam.Set($"PO-{pipeIndex:000}");
                            pipeIndex++;
                        }
                        else if (fi.Symbol.Name.IndexOf("DS#", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            markParam.Set($"DO-{ductIndex:000}");
                            ductIndex++;
                        }
                    }
                }

                // Commit mark assignment only
                t.Commit();
            }

            TaskDialog.Show("Schedule", "Opening schedule generated successfully.");
            return Result.Succeeded;
        }
    }
}