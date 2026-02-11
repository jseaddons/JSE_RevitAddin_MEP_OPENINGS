using JSE_RevitAddin_MEP_OPENINGS.Services.MultiFloor;
using System.Linq;

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
        private IPerformanceMonitor? _performanceMonitor;
        private readonly bool _forceDetectionMode;

        public OpeningCommandOrchestrator(Document document, UIDocument uiDocument, Dictionary<string, double> uiClearances = null, object markPrefixes = null, bool forceDetectionMode = false, IPerformanceMonitor? performanceMonitor = null)
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

                    // Generate performance report using actual counts from result
                    _performanceMonitor.GenerateReport(result.PlacedCount, result.ClusterCount);
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

        /// <summary>
        /// Execute multi-floor placement with delegation to FloorBatchProcessor.
        /// </summary>
        public void ExecuteMultiFloorPlacement(List<Level> selectedLevels, OpeningFilter filter)
        {
            if (selectedLevels == null || selectedLevels.Count == 0) return;

            // Initialize performance monitoring
            if (_performanceMonitor == null)
            {
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                _performanceMonitor = new PlacementPerformanceMonitor($"SleevePlacement_MultiFloor_{timestamp}.log");
            }

            try
            {
                // Create multi-floor processor
                var processor = new FloorBatchProcessor(_document, _performanceMonitor);

                // Process all selected floors
                var result = processor.ProcessFloors(selectedLevels, filter, chunkSize: 5);

                // Report results
                SafeFileLogger.SafeAppendText("multifloor.log",
                    $"[{DateTime.Now}] ✅ Multi-floor processing complete:\n" +
                    $"   Successful: {result.SuccessfulFloors.Count} floors\n" +
                    $"   Failed: {result.FailedFloors.Count} floors\n" +
                    $"   Total sleeves placed: {result.TotalSleevesPlaced}\n" +
                    $"   Total clusters: {result.TotalClustersFormed}\n");

                if (result.FailedFloors.Any())
                {
                    TaskDialog.Show("Warning", 
                        $"Some floors failed:\n{string.Join("\n", result.FailedFloors)}");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[OpeningCommandOrchestrator] Multi-Floor Error: {ex.Message}");
                }
                throw;
            }
        }
    }
}

