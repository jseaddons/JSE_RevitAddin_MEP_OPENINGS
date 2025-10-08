using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// Universal cluster command wrapper - now uses UniversalClusterService
    /// Can be called from any context (ICommand, ExternalEvent, etc.)
    /// </summary>
    public class UniversalClusterCommand : ICommand
    {
        private readonly string _targetCategory;
        
        public UniversalClusterCommand(string targetCategory)
        {
            _targetCategory = targetCategory ?? throw new ArgumentNullException(nameof(targetCategory));
        }
        
        public void Execute(UIApplication app)
        {
            try
            {
                DebugLogger.Info($"[UniversalClusterCommand] Starting clustering for {_targetCategory}");
                
                var doc = app.ActiveUIDocument.Document;
                var uiDoc = app.ActiveUIDocument;
                
                // ⚠️ CRITICAL: Transaction is required by UniversalClusterService
                using (var tx = new Transaction(doc, $"Cluster {_targetCategory} Openings"))
                {
                    tx.Start();
                    
                    var clusterService = new UniversalClusterService();
                    var (placedCount, deletedCount) = clusterService.ClusterSleeves(doc, _targetCategory, uiDoc);
                    
                    tx.Commit();
                    
                    DebugLogger.Info($"[UniversalClusterCommand] ✓ Clustering complete for {_targetCategory}: {placedCount} clusters placed, {deletedCount} individual sleeves deleted");
                    
                    // Optional: Show user feedback for manual testing
                    // TaskDialog.Show("Clustering Complete", $"{_targetCategory}: {placedCount} clusters placed, {deletedCount} sleeves deleted");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterCommand] Error clustering {_targetCategory}: {ex.Message}");
                DebugLogger.Error($"[UniversalClusterCommand] Stack trace: {ex.StackTrace}");
                // Don't show TaskDialog here - let orchestrator handle errors
                throw; // Re-throw to let caller handle
            }
        }
    }
}

