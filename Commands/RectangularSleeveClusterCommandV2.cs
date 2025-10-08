using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class RectangularSleeveClusterCommandV2 : IExternalCommand
    {
        private readonly string _targetCategory;
        
        // Parameterless constructor for direct execution
        public RectangularSleeveClusterCommandV2() : this(null)
        {
        }
        
        // Constructor with category filter (for use by SleevePlacementExternalEvent)
        public RectangularSleeveClusterCommandV2(string targetCategory)
        {
            _targetCategory = targetCategory; // "Ducts", "Pipes", "Duct Accessories", "Cable Trays", or null for all
        }
        
        // Helper struct for grouping key
        struct SleeveGroupKey
        {
            public string hostType;
            public string systemType;
            public string orientation;
            public SleeveGroupKey(string hostType, string systemType, string orientation)
            {
                this.hostType = hostType;
                this.systemType = systemType;
                this.orientation = orientation;
            }
            // Override Equals and GetHashCode for dictionary/grouping
            public override bool Equals(object obj)
            {
                if (!(obj is SleeveGroupKey)) return false;
                var other = (SleeveGroupKey)obj;
                return hostType == other.hostType && systemType == other.systemType && orientation == other.orientation;
            }
            public override int GetHashCode()
            {
                return (hostType, systemType, orientation).GetHashCode();
            }
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string logFileName = $"RectangularSleeveClusterCommandV2_{timestamp}.log";
            try { DebugLogger.InitLogFile(logFileName); }
            catch (Exception ex) { TaskDialog.Show("Logging Error", $"Failed to initialize log file: {logFileName}\n{ex.Message}"); }
            DebugLogger.Log($"Assembly version: {Assembly.GetExecutingAssembly().GetName().Version}");
            DebugLogger.Log($"Assembly build timestamp: {System.IO.File.GetLastWriteTime(Assembly.GetExecutingAssembly().Location):o}");
            DebugLogger.Log("*** RectangularSleeveClusterCommandV2 START ***");
            
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // ⚠️ CRITICAL: Use cluster configuration from ClusterConfigurationManager
                double toleranceMm = ClusterConfigurationManager.Instance.JoinOpeningsDistance;
                DebugLogger.Log($"[RectangularCluster] ⚠️ Using JoinOpeningsDistance: {toleranceMm}mm (from {ClusterConfigurationManager.Instance.ConfigurationSource})");
                DebugLogger.Log($"[RectangularCluster] Configuration: {ClusterConfigurationManager.Instance.GetConfigurationSummary()}");
                DebugLogger.Log($"[RectangularCluster] Target category: {_targetCategory ?? "ALL"}");

                int placedCount = 0;
                int deletedCount = 0;

                // Use the UniversalClusterService to do all the work
                using (var tx = new Transaction(doc, "Place Clustered Rectangular Openings V2"))
                {
                    tx.Start();

                    var clusterService = new UniversalClusterService();
                    (placedCount, deletedCount) = clusterService.ClusterSleeves(doc, _targetCategory, uiDoc);

                    tx.Commit();
                }

                DebugLogger.Log($"All clustering processed. Summary: {placedCount} openings placed, {deletedCount} sleeves deleted.");
                
                // Show user feedback
                TaskDialog.Show("Clustering Complete", 
                    $"Successfully clustered sleeves:\n\n" +
                    $"• Clusters created: {placedCount}\n" +
                    $"• Individual sleeves replaced: {deletedCount}\n" +
                    $"• Join distance: {toleranceMm}mm\n" +
                    $"• Category: {_targetCategory ?? "All categories"}");
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[RectangularCluster] FATAL ERROR: {ex.Message}");
                DebugLogger.Error($"[RectangularCluster] Stack trace: {ex.StackTrace}");
                
                message = $"Clustering failed: {ex.Message}";
                TaskDialog.Show("Clustering Error", 
                    $"An error occurred during clustering:\n\n{ex.Message}\n\n" +
                    $"Check the log file for details:\n{logFileName}");
                
                return Result.Failed;
            }
        }
    }
}
