using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;

using Autodesk.Revit.Attributes;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    public class ClusterMergeCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                // Filter selection for cluster families
                var selectedIds = uidoc.Selection.GetElementIds();
                if (selectedIds.Count != 2)
                {
                    message = "Please select exactly two cluster family instances.";
                    return Result.Failed;
                }

                var clusters = new List<ClusterFamilyData>();
                foreach (var id in selectedIds)
                {
                    var elem = doc.GetElement(id) as FamilyInstance;
                    if (elem == null || !elem.Symbol.Name.StartsWith("Cluster"))
                    {
                        message = "Selection must be cluster family instances.";
                        return Result.Failed;
                    }
                    clusters.Add(new ClusterFamilyData(elem));
                }

                // Check distance
                double dist = clusters[0].Position.DistanceTo(clusters[1].Position);
                if (dist > UnitUtils.ConvertToInternalUnits(100, UnitTypeId.Millimeters))
                {
                    message = "Clusters are more than 100mm apart.";
                    return Result.Failed;
                }


                // Detect scenario (horizontal, vertical, offset)
                var scenario = DetectMergeScenario(clusters[0], clusters[1]);
                var mergedParams = ComputeMergedParameters(clusters[0], clusters[1], scenario);

                // Place merged family (single parametric type, all parameters set as instance)
                using (Transaction t = new Transaction(doc, "Place Merged Cluster"))
                {
                    t.Start();
                    FamilySymbol mergedSymbol = FindMergedFamilySymbol(doc, scenario);
                    if (mergedSymbol == null)
                    {
                        message = $"Merged family for scenario '{scenario}' not loaded.";
                        t.RollBack();
                        return Result.Failed;
                    }
                    if (!mergedSymbol.IsActive) mergedSymbol.Activate();
                    var mergedInstance = doc.Create.NewFamilyInstance(mergedParams.Position, mergedSymbol, StructuralType.NonStructural);
                    SetMergedParameters(mergedInstance, mergedParams);
                    // Delete old clusters
                    doc.Delete(clusters[0].Id);
                    doc.Delete(clusters[1].Id);
                    t.Commit();
                }
                TaskDialog.Show("Cluster Merge", $"Merged cluster placed for scenario: {scenario}.");
                return Result.Succeeded;
            }

            private string DetectMergeScenario(ClusterFamilyData c1, ClusterFamilyData c2)
            {
                // Simple logic: horizontal if Y difference < X difference, else vertical
                double dx = Math.Abs(c1.Position.X - c2.Position.X);
                double dy = Math.Abs(c1.Position.Y - c2.Position.Y);
                if (dx > dy) return "Horizontal";
                else return "Vertical";
                // Extend for offset/L-shape as needed
            }

            private MergedClusterParams ComputeMergedParameters(ClusterFamilyData c1, ClusterFamilyData c2, string scenario)
            {
                // Compute parameters based on scenario
                var pos = c1.Position.Add(c2.Position).Divide(2);
                if (scenario == "Horizontal")
                {
                    double w1 = Math.Max(c1.Width, c2.Width);
                    double w2 = Math.Min(c1.Width, c2.Width);
                    double h1 = c1.Height;
                    double h2 = c2.Height;
                    return new MergedClusterParams { Position = pos, W1 = w1, W2 = w2, H1 = h1, H2 = h2 };
                }
                else // Vertical
                {
                    double w1 = c1.Width;
                    double w2 = c2.Width;
                    double h1 = Math.Max(c1.Height, c2.Height);
                    double h2 = Math.Min(c1.Height, c2.Height);
                    return new MergedClusterParams { Position = pos, W1 = w1, W2 = w2, H1 = h1, H2 = h2 };
                }
            }

            private FamilySymbol FindMergedFamilySymbol(Document doc, string scenario)
            {
                // Find the correct family symbol for the scenario
                string familyName = $"MergedCluster{scenario}";
                var collector = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol));
                foreach (FamilySymbol symbol in collector)
                {
                    if (symbol.Name == familyName)
                        return symbol;
                }
                return null;
            }

            private MergedClusterParams ComputeMergedParameters(ClusterFamilyData c1, ClusterFamilyData c2)
            {
                // TODO: Implement logic for all scenarios (horizontal, vertical, offset)
                // Example: horizontal merge
                var pos = c1.Position.Add(c2.Position).Divide(2);
                double w1 = Math.Max(c1.Width, c2.Width);
                double w2 = Math.Min(c1.Width, c2.Width);
                double h1 = c1.Height;
                double h2 = c2.Height;
                return new MergedClusterParams { Position = pos, W1 = w1, W2 = w2, H1 = h1, H2 = h2 };
            }

            private FamilySymbol FindMergedFamilySymbol(Document doc, MergedClusterParams mergedParams)
            {
                // TODO: Find the correct family symbol for the merged cluster
                var collector = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol));
                foreach (FamilySymbol symbol in collector)
                {
                    if (symbol.Name.StartsWith("MergedCluster"))
                        return symbol;
                }
                return null;
            }

            private void SetMergedParameters(FamilyInstance instance, MergedClusterParams mergedParams)
            {
                // TODO: Set all relevant parameters
                SetParam(instance, "W1", mergedParams.W1);
                SetParam(instance, "W2", mergedParams.W2);
                SetParam(instance, "H1", mergedParams.H1);
                SetParam(instance, "H2", mergedParams.H2);
            }

            private void SetParam(FamilyInstance instance, string paramName, double value)
            {
                var param = instance.LookupParameter(paramName);
                if (param != null && !param.IsReadOnly)
                    param.Set(value);
            }

            private class MergedClusterParams
            {
                public XYZ Position { get; set; }
                public double W1 { get; set; }
                public double W2 { get; set; }
                public double H1 { get; set; }
                public double H2 { get; set; }
                // Add more as needed (W3, H3, etc.)
            }
        }

    }
