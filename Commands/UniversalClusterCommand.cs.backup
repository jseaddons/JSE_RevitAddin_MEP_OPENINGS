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
        private readonly string _xmlFilePath;

        public UniversalClusterCommand(string targetCategory, string xmlFilePath = null)
        {
            _targetCategory = targetCategory ?? throw new ArgumentNullException(nameof(targetCategory));
            _xmlFilePath = xmlFilePath; // Optional - if null, searches all XML files (backward compatibility)
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
                    var (placedCount, deletedCount) = clusterService.ClusterSleeves(doc, _targetCategory, uiDoc, _xmlFilePath);
                    
                    tx.Commit();
                    
                    DebugLogger.Info($"[UniversalClusterCommand] ✓ Clustering complete for {_targetCategory}: {placedCount} clusters placed, {deletedCount} individual sleeves deleted");
                    
                    // Show user feedback
                    if (placedCount > 0 || deletedCount > 0)
                    {
                        TaskDialog.Show("Clustering Complete", 
                            $"Category: {_targetCategory}\n\n" +
                            $"✓ {placedCount} cluster sleeve(s) created\n" +
                            $"✓ {deletedCount} individual sleeve(s) replaced\n\n" +
                            $"Join distance: {ClusterConfigurationManager.Instance.JoinOpeningsDistance}mm");
                    }
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

