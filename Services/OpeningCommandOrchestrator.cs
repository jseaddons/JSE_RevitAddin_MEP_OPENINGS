using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Data;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Streamlined Orchestrator that delegates all placement and clustering logic 
    /// to the specialized PlacementWorkflowOrchestrator.
    /// </summary>
    public class OpeningCommandOrchestrator
    {
        private readonly Document _document;
        private readonly UIDocument _uiDocument;
        private readonly Dictionary<string, double> _uiClearances;
        private IPerformanceMonitor _performanceMonitor;
        private readonly bool _forceDetectionMode;

        public OpeningCommandOrchestrator(Document document, UIDocument uiDocument, Dictionary<string, double> uiClearances = null, object markPrefixes = null, bool forceDetectionMode = false, IPerformanceMonitor performanceMonitor = null)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _uiDocument = uiDocument ?? throw new ArgumentNullException(nameof(uiDocument));
            _uiClearances = uiClearances ?? new Dictionary<string, double>();
            _forceDetectionMode = forceDetectionMode;
            _performanceMonitor = performanceMonitor;
        }

        /// <summary>
        /// Set UI clearance settings from the main dialog
        /// </summary>
        public void SetUIClearances(Dictionary<string, double> clearances)
        {
            _uiClearances.Clear();
            if (clearances != null)
            {
                foreach (var kvp in clearances)
                {
                    _uiClearances[kvp.Key] = kvp.Value;
                }
            }
        }

        /// <summary>
        /// Execute multiple filters with delegation to PlacementWorkflowOrchestrator.
        /// </summary>
        public void ExecuteMultipleFilters(List<OpeningFilter> filters, bool showProgress = true)
        {
            if (filters == null || filters.Count == 0) return;

            // Initialize performance monitoring
            if (_performanceMonitor == null)
            {
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                _performanceMonitor = new PlacementPerformanceMonitor($"SleevePlacement_Batch_{timestamp}.log");
            }

            try
            {
                // Delegation: The new PlacementWorkflowOrchestrator handles the optimized sequence
                var workflowOrchestrator = new PlacementWorkflowOrchestrator(
                    _document,
                    () => new SleeveDbContext(_document),
                    msg => { if (!DeploymentConfiguration.DeploymentMode) System.Diagnostics.Debug.WriteLine(msg); },
                    _performanceMonitor);

                using (var tracker = _performanceMonitor.TrackOperation("Complete Opening Placement Workflow"))
                {
                    var result = workflowOrchestrator.ExecuteOptimizedPlacementWorkflow(filters, tracker);

                    // Generate performance report
                    _performanceMonitor.GenerateReport(result.PlacedCount, 0);
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[OpeningCommandOrchestrator] Error: {ex.Message}");
                }
                throw;
            }
        }
    }
}
