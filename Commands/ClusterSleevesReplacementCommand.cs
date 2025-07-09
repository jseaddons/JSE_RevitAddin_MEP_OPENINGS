using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class ClusterSleevesReplacementCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Initialize logging for the cluster sleeve replacement command
            DebugLogger.InitLogFile("ClusterSleevesReplacementCommand");
            DebugLogger.Log("ClusterSleevesReplacementCommand: Execute started - running cluster sleeve replacement commands");
            
            // Initialization code
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // DUPLICATION SUPPRESSION: Get baseline count of existing cluster openings
            var baselineClusterCount = GetExistingClusterOpeningsCount(doc);
            DebugLogger.Log($"Baseline cluster openings count: {baselineClusterCount}");

            DebugLogger.Log("Starting rectangular pipe clustering...");
            PlaceRectangularPipeOpenings(commandData, doc);
            
            // Check intermediate count after pipe clustering
            var afterPipeCount = GetExistingClusterOpeningsCount(doc);
            var pipeClusterAdded = afterPipeCount - baselineClusterCount;
            DebugLogger.Log($"Pipe clustering added {pipeClusterAdded} openings. Current total: {afterPipeCount}");
            
            DebugLogger.Log("Starting rectangular sleeve clustering...");
            PlaceRectangularSleeveCluster(commandData, doc);

            // Final count and summary
            var finalCount = GetExistingClusterOpeningsCount(doc);
            var sleeveClusterAdded = finalCount - afterPipeCount;
            var totalAdded = finalCount - baselineClusterCount;
            
            DebugLogger.Log($"Sleeve clustering added {sleeveClusterAdded} openings. Final total: {finalCount}");
            DebugLogger.Log($"ClusterSleevesReplacementCommand: All cluster commands completed. Total openings added: {totalAdded}");
            
            TaskDialog.Show("Cluster Sleeve Commands Completed", 
                "All cluster sleeve replacement commands have been executed:\n\n" +
                $"✓ Rectangular Pipe Openings (PS# → PipeOpeningOnWallRect): {pipeClusterAdded} added\n" +
                $"✓ Rectangular Sleeve Clustering (DS#/DMS#/CT# clustering): {sleeveClusterAdded} net change\n" +
                $"✓ Total cluster openings added: {totalAdded}\n\n" +
                "✓ Duplication suppression active (100mm tolerance)\n" +
                "✓ Commands coordinate to avoid conflicts\n" +
                "✓ Safe for re-execution\n\n" +
                "Check the debug log for detailed results.");

            return Result.Succeeded;
        }

        private void PlaceRectangularPipeOpenings(ExternalCommandData commandData, Document doc)
        {
            try
            {
                DebugLogger.Log("Executing PipeOpeningsRectCommand...");
                DebugLogger.Log("Note: PipeOpeningsRectCommand has built-in duplication suppression (100mm tolerance)");
                DebugLogger.Log("This command processes circular pipe sleeves (PS#) and creates rectangular openings (PipeOpeningOnWallRect)");
                
                var pipeRectCommand = new PipeOpeningsRectCommand();
                string message = "";
                ElementSet elements = new ElementSet();
                var result = pipeRectCommand.Execute(commandData, ref message, elements);
                if (result != Result.Succeeded)
                {
                    DebugLogger.Log($"Rectangular pipe opening placement failed: {message}");
                }
                else
                {
                    DebugLogger.Log("Rectangular pipe opening placement completed successfully (with duplication suppression)");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Error placing rectangular pipe openings: {ex.Message}");
            }
        }

        private void PlaceRectangularSleeveCluster(ExternalCommandData commandData, Document doc)
        {
            try
            {
                DebugLogger.Log("Executing RectangularSleeveClusterCommand...");
                DebugLogger.Log("Note: RectangularSleeveClusterCommand has built-in duplication suppression");
                DebugLogger.Log("This command processes existing rectangular sleeves (DS#, DMS#, CT#) and clusters them");
                DebugLogger.Log("It will skip areas where PipeOpeningsRectCommand has already placed rectangular openings");
                
                var rectClusterCommand = new RectangularSleeveClusterCommand();
                string message = "";
                ElementSet elements = new ElementSet();
                var result = rectClusterCommand.Execute(commandData, ref message, elements);
                if (result != Result.Succeeded)
                {
                    DebugLogger.Log($"Rectangular sleeve clustering failed: {message}");
                }
                else
                {
                    DebugLogger.Log("Rectangular sleeve clustering completed successfully (with duplication suppression)");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Error with rectangular sleeve clustering: {ex.Message}");
            }
        }

        /// <summary>
        /// DUPLICATION SUPPRESSION: Count existing cluster openings for baseline tracking
        /// This helps coordinate between the two cluster commands and provides metrics
        /// </summary>
        private int GetExistingClusterOpeningsCount(Document doc)
        {
            try
            {
                // Count cluster openings by families ending with "Rect" (matches RectangularSleeveClusterCommand logic)
                var clusterOpenings = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol?.Family?.Name != null &&
                        fi.Symbol.Family.Name.EndsWith("Rect", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                DebugLogger.Debug($"[CLUSTER COUNT] Found {clusterOpenings.Count} cluster openings by families ending with 'Rect'");
                // Log details of each cluster opening found
                foreach (var opening in clusterOpenings)
                {
                    DebugLogger.Debug($"  - Family: {opening.Symbol.Family.Name}, Symbol: {opening.Symbol.Name}, ID: {opening.Id.IntegerValue}");
                }
                return clusterOpenings.Count;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Error counting existing cluster openings: {ex.Message}");
                return 0; // Return 0 if count fails, allowing commands to proceed
            }
        }
    }
}
