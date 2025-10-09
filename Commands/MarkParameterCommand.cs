using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// Mark Parameter Command - Applies MEPMARK to cluster sleeves for a specific category
    /// Can be called from any context (ICommand, ExternalEvent, etc.)
    /// No UI dialogs - fully automated workflow using prefixes configured in main UI
    /// </summary>
    public class MarkParameterCommand : ICommand
    {
        private readonly string _targetCategory;
        private readonly string _projectPrefix;
        private readonly string _disciplinePrefix;
        private readonly bool _remarkAll;
        
        public MarkParameterCommand(string targetCategory, string projectPrefix, string disciplinePrefix, bool remarkAll = false)
        {
            _targetCategory = targetCategory ?? throw new ArgumentNullException(nameof(targetCategory));
            _projectPrefix = projectPrefix ?? throw new ArgumentNullException(nameof(projectPrefix));
            _disciplinePrefix = disciplinePrefix ?? throw new ArgumentNullException(nameof(disciplinePrefix));
            _remarkAll = remarkAll;
        }
        
        public void Execute(UIApplication app)
        {
            try
            {
                DebugLogger.Info($"[MarkParameterCommand] Starting MEPMARK application for {_targetCategory}");
                DebugLogger.Info($"[MarkParameterCommand] Project Prefix: '{_projectPrefix}', Discipline Prefix: '{_disciplinePrefix}'");
                
                var doc = app.ActiveUIDocument.Document;
                var uiDoc = app.ActiveUIDocument;
                
                // ⚠️ CRITICAL: Transaction management (follows existing pattern)
                using (var tx = new Transaction(doc, $"Mark {_targetCategory} Clusters"))
                {
                    tx.Start();
                    
                    var markService = new MarkParameterService();
                    var (processedCount, errorCount) = markService.ApplyMepMarkToClusters(
                        doc, _targetCategory, _projectPrefix, _disciplinePrefix, _remarkAll);
                    
                    tx.Commit();
                    
                    DebugLogger.Info($"[MarkParameterCommand] ✓ MEPMARK complete for {_targetCategory}: {processedCount} clusters processed, {errorCount} errors");
                    
                    // ✅ REMOVED: TaskDialog - not needed in automated workflow
                    // User already configured prefixes in UI, no need for confirmation dialog
                    // if (_showDialog && processedCount > 0)
                    // {
                    //     TaskDialog.Show("MEPMARK Applied", 
                    //         $"Category: {_targetCategory}\n\n" +
                    //         $"✓ {processedCount} cluster(s) marked\n" +
                    //         $"✓ Project Prefix: {_projectPrefix}\n" +
                    //         $"✓ Discipline Prefix: {_disciplinePrefix}");
                    // }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[MarkParameterCommand] Error applying MEPMARK to {_targetCategory}: {ex.Message}");
                DebugLogger.Error($"[MarkParameterCommand] Stack trace: {ex.StackTrace}");
                // Don't show TaskDialog here - let orchestrator handle errors
                throw; // Re-throw to let caller handle
            }
        }
    }
}
