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
                // Diagnostic start
                Serilog.Log.Logger?.Information("[ClusterMerge] Execute start");

                try
                {

                // Filter selection for cluster families
                var selectedIds = uidoc.Selection.GetElementIds();
                Serilog.Log.Logger?.Debug("[ClusterMerge] SelectedIds count={Count}", selectedIds.Count);
                if (selectedIds.Count != 2)
                {
                    message = "Please select exactly two cluster family instances.";
                    Serilog.Log.Logger?.Warning("[ClusterMerge] Invalid selection count: {Count}", selectedIds.Count);
                    return Result.Failed;
                }

                var clusters = new List<ClusterFamilyData>();
                foreach (var id in selectedIds)
                {
                    Serilog.Log.Logger?.Debug("[ClusterMerge] Inspecting selected id={Id}", id.IntegerValue);
                    var elem = doc.GetElement(id) as FamilyInstance;
                    if (elem == null)
                    {
                        message = "Selection must be cluster family instances.";
                        Serilog.Log.Logger?.Warning("[ClusterMerge] Selected element is not a FamilyInstance: {Id}", id.IntegerValue);
                        return Result.Failed;
                    }
                    var symName = elem.Symbol?.Name ?? "<null>";
                    Serilog.Log.Logger?.Debug("[ClusterMerge] FamilyInstance id={Id} symbol={Symbol}", id.IntegerValue, symName);
                    if (!symName.StartsWith("Cluster"))
                    {
                        message = "Selection must be cluster family instances.";
                        Serilog.Log.Logger?.Warning("[ClusterMerge] FamilyInstance symbol not a cluster: {Symbol}", symName);
                        return Result.Failed;
                    }
                    clusters.Add(new ClusterFamilyData(elem));
                }

                // Check distance
                double dist = clusters[0].Position.DistanceTo(clusters[1].Position);
                Serilog.Log.Logger?.Debug("[ClusterMerge] Distance between clusters={Dist}", dist);
                if (dist > UnitUtils.ConvertToInternalUnits(100, UnitTypeId.Millimeters))
                {
                    message = "Clusters are more than 100mm apart.";
                    Serilog.Log.Logger?.Warning("[ClusterMerge] Clusters too far apart: {Dist} > 100mm", dist);
                    return Result.Failed;
                }


                // Detect scenario (horizontal, vertical, offset)
                var scenario = DetectMergeScenario(clusters[0], clusters[1]);
                Serilog.Log.Logger?.Debug("[ClusterMerge] Detected scenario={Scenario}", scenario);
                var mergedParams = ComputeMergedParameters(clusters[0], clusters[1], scenario);

                // Place merged family (single parametric type, all parameters set as instance)
                using (Transaction t = new Transaction(doc, "Place Merged Cluster"))
                {
                    t.Start();
                    Serilog.Log.Logger?.Debug("[ClusterMerge] Transaction started");
                    FamilySymbol? mergedSymbol = FindMergedFamilySymbol(doc, scenario);
                    if (mergedSymbol == null)
                    {
                        message = $"Merged family for scenario '{scenario}' not loaded.";
                        Serilog.Log.Logger?.Error("[ClusterMerge] Merged family symbol not found for scenario={Scenario}", scenario);
                        t.RollBack();
                        return Result.Failed;
                    }
                    Serilog.Log.Logger?.Debug("[ClusterMerge] Found merged symbol={Symbol}", mergedSymbol.Name);
                    if (!mergedSymbol.IsActive) mergedSymbol.Activate();
                    var mergedInstance = doc.Create.NewFamilyInstance(mergedParams.Position, mergedSymbol, StructuralType.NonStructural);
                    SetMergedParameters(mergedInstance, mergedParams);
                    // Delete old clusters
                    Serilog.Log.Logger?.Debug("[ClusterMerge] Deleting old cluster ids {Id1} and {Id2}", clusters[0].Id.IntegerValue, clusters[1].Id.IntegerValue);
                    doc.Delete(clusters[0].Id);
                    doc.Delete(clusters[1].Id);
                    t.Commit();
                    Serilog.Log.Logger?.Information("[ClusterMerge] Transaction committed, merged instance created with id={Id}", mergedInstance.Id.IntegerValue);
                }
                TaskDialog.Show("Cluster Merge", $"Merged cluster placed for scenario: {scenario}.");
                Serilog.Log.Logger?.Information("[ClusterMerge] Execute completed successfully");
                return Result.Succeeded;
            }
            catch (System.Exception ex)
            {
                Serilog.Log.Logger?.Error(ex, "[ClusterMerge] Exception during merge");
                TaskDialog.Show("Cluster Merge Error", ex.ToString());
                message = ex.Message;
                return Result.Failed;
            }
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

            private FamilySymbol? FindMergedFamilySymbol(Document doc, string scenario)
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

            private FamilySymbol? FindMergedFamilySymbol(Document doc, MergedClusterParams mergedParams)
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
                public XYZ Position { get; set; } = new XYZ(0, 0, 0);
                public double W1 { get; set; }
                public double W2 { get; set; }
                public double H1 { get; set; }
                public double H2 { get; set; }
                // Add more as needed (W3, H3, etc.)
            }
        }

    }
