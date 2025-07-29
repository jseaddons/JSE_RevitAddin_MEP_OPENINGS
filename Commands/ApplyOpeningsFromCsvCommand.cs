using System;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class ApplyOpeningsFromCsvCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiDoc = commandData.Application.ActiveUIDocument;
            var doc = uiDoc.Document;

            const string csvPath = @"C:\Temp\opening_requirements.csv";
            if (!File.Exists(csvPath))
            {
                TaskDialog.Show("Error", $"CSV not found: {csvPath}");
                return Result.Failed;
            }

            // Auto-select opening families
            var symbols = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .ToList();
            var pipeSymbol = symbols.FirstOrDefault(sym =>
                sym.Family.Name.IndexOf("OpeningOnWall", StringComparison.OrdinalIgnoreCase) >= 0 &&
                sym.Name.IndexOf("PS#", StringComparison.OrdinalIgnoreCase) >= 0);
            var ductSymbol = symbols.FirstOrDefault(sym =>
                sym.Family.Name.IndexOf("OpeningOnWall", StringComparison.OrdinalIgnoreCase) >= 0 &&
                sym.Name.IndexOf("DS#", StringComparison.OrdinalIgnoreCase) >= 0);
            if (pipeSymbol == null || ductSymbol == null)
            {
                TaskDialog.Show("Error", "Required opening family symbols not loaded.");
                return Result.Failed;
            }

            string[] lines = File.ReadAllLines(csvPath);
            if (lines.Length < 2)
            {
                TaskDialog.Show("Error", "CSV contains no data.");
                return Result.Cancelled;
            }

            using (var tx = new Transaction(doc, "Apply Openings from CSV"))
            {
                tx.Start();

                var pipePlacer = new PipeSleevePlacer(doc);
                var ductPlacer = new DuctSleevePlacer(doc);

                foreach (var line in lines.Skip(1))
                {
                    var parts = line.Split(',');
                    if (parts.Length < 8) continue;
                    // Parse values
                    long wallId = (long)Convert.ToInt32(parts[0]);
                    double x = Convert.ToDouble(parts[1]);
                    double y = Convert.ToDouble(parts[2]);
                    double z = Convert.ToDouble(parts[3]);
                    double pipeDia = Convert.ToDouble(parts[4]);
                    double ductWidth = Convert.ToDouble(parts[5]);
                    double ductHeight = Convert.ToDouble(parts[6]);
                    string mepType = parts[7];

                    var wall = doc.GetElement(new ElementId((int)wallId)) as Wall;
                    if (wall == null) continue;
                    var intersection = new XYZ(x, y, z);
                    var direction = XYZ.BasisX; // default direction; adjust if needed

                    if (mepType.Equals("Pipe", StringComparison.OrdinalIgnoreCase))
                    {
                        // Default: do not rotate (X-axis wall logic)
                        bool shouldRotate = false;
                        int pipeElementId = 0; // No pipe element ID from CSV, so pass 0
                        pipePlacer.PlaceSleeve(intersection, pipeDia, direction, pipeSymbol, wall, shouldRotate, pipeElementId);
                    }
                    else if (mepType.Equals("Duct", StringComparison.OrdinalIgnoreCase))
                    {
                        // For wall, use wall.Orientation as hostNormal
                        ductPlacer.PlaceDuctSleeve(null, intersection, ductWidth, ductHeight, direction, ductSymbol, wall, wall.Orientation);
                    }
                }

                tx.Commit();
            }

            // TaskDialog.Show("Done", "Applied openings from CSV."); // Info dialog removed as requested
            return Result.Succeeded;
        }
    }
}
