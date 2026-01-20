using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Calculation;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Workflow
{
    /// <summary>
    /// Unified Orchestrator for V3 Batch Mode (Calculate -> Persist -> Place)
    /// Handles both Individual and Cluster sleeves.
    /// </summary>
    public class UnifiedSleeveWorkflow
    {
        private readonly SleeveCalculationService _individualCalculationService;
        private readonly BatchClusterCalculationService _clusterCalculationService;
        private readonly UnifiedPlacementService _placementService;
        private readonly IClashZoneRepository _repository;

        // Flags for operation mode
        public bool EnableClustering { get; set; } = true;
        public bool EnableIndividualPlacement { get; set; } = true;

        public UnifiedSleeveWorkflow(
            SleeveCalculationService individualCalculationService,
            BatchClusterCalculationService clusterCalculationService,
            UnifiedPlacementService placementService,
            IClashZoneRepository repository)
        {
            _individualCalculationService = individualCalculationService ?? throw new ArgumentNullException(nameof(individualCalculationService));
            _clusterCalculationService = clusterCalculationService ?? throw new ArgumentNullException(nameof(clusterCalculationService));
            _placementService = placementService ?? throw new ArgumentNullException(nameof(placementService));
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
        }

        public void Execute(Document doc, string targetCategory = null, bool isBulkMode = true)
        {
            string batchId = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            SafeFileLogger.SafeAppendText("unified_workflow.log", $"[{DateTime.Now}] Starting Unified Workflow Batch: {batchId}\n");

            // 1. Fetch Candidates (Unresolved Zones)
            // Ideally filtered by category if provided, or current selection/view scope.
            // For now, assume 'All Unresolved' or 'Unresolved by Category'.
            List<ClashZone> candidates = new List<ClashZone>();
            
            if (!string.IsNullOrEmpty(targetCategory))
            {
                 // Get by category
                 var zones = _repository.GetClashZonesByCategory(targetCategory);
                 candidates = zones.Where(z => 
                    !z.IsResolved && 
                    !z.IsClusterResolved && 
                    !z.IsCombinedResolved &&
                    (z.MarkedForClusterProcess != false)).ToList();
            }
            else
            {
                // Get all categories? Better to iterate distinct categories to avoid giant fetch.
                var cats = _repository.GetDistinctCategories();
                foreach(var cat in cats)
                {
                    var zones = _repository.GetClashZonesByCategory(cat);
                    // ✅ FIX: Filter candidates by MarkedForClusterProcess
                    // Only process zones that are flagged for clustering (or not yet flagged - assuming default is true/null)
                    // If MarkedForClusterProcess is FALSE, it should be skipped.
                    // Also filter out resolved zones.
                    candidates.AddRange(zones.Where(z => 
                        !z.IsResolved && 
                        !z.IsClusterResolved && 
                        !z.IsCombinedResolved &&
                        (z.MarkedForClusterProcess != false))); // Skip explicitly false
                }
            }

            SafeFileLogger.SafeAppendText("unified_workflow.log", $"[{DateTime.Now}] Fetched {candidates.Count} candidate zones.\n");
            
            if (candidates.Count == 0) return;

            // 2. Phase 1a: Calculate Individual Sleeves
            if (EnableIndividualPlacement)
            {
                SafeFileLogger.SafeAppendText("unified_workflow.log", $"[{DateTime.Now}] Calculating Individual Sleeves...\n");
                _individualCalculationService.CalculateAndSave(doc, candidates, batchId);
            }

            // 3. Phase 1b: Calculate Clusters
            if (EnableClustering)
            {
                SafeFileLogger.SafeAppendText("unified_workflow.log", $"[{DateTime.Now}] Calculating Clusters...\n");
                // Need to group by Filter/Combo? BatchClusterCalculationService expects 'comboId' and 'filterId'.
                // If we pass generic list, we might lose filter context.
                // However, the service logic uses group-by inside.
                // Actually `CalculateAndSave` takes `comboId` and `filterId`. This implies per-combo processing.
                // We should iterate unique combos in candidates.
                
                var groupedByCombo = candidates.GroupBy(z => new { z.ComboId, FilterId = GetFilterIdForZone(z) }); 
                // Wait, clashZone doesn't have FilterId property directly usually? 
                // MapClashZone doesn't show FilterId property (it has filter name in metadata).
                // I need GetComboAndFilterId from Repo?
                // Repo has `(int ComboId, int FilterId) GetComboAndFilterId(System.Guid clashZoneId);`
                
                // Optimized: Group in memory if possible or fetch efficiently.
                // Iterating candidates and calling GetComboAndFilterId for each might be slow.
                // But Phase 1 is decoupled, so maybe okay.
                // OR: rely on `z.ComboId` (it is in the class). FilterId might need lookup.
                // Actually, `ClusterCalculationService` needs it mainly for saving result.
                
                foreach (var grp in groupedByCombo)
                {
                     // We may not have FilterId easily.
                     // Helper:
                     var first = grp.First();
                     var (comboId, filterId) = _repository.GetComboAndFilterId(first.Id); // Get from DB for first
                     
                     _clusterCalculationService.CalculateAndSave(grp.ToList(), targetCategory ?? "Mixed", comboId, filterId, doc);
                }
            }

            // 4. Phase 2: Unified Placement
            SafeFileLogger.SafeAppendText("unified_workflow.log", $"[{DateTime.Now}] Executing Unified Placement (Mode: {(isBulkMode ? "Bulk" : "Sequential")})...\n");
            _placementService.PlaceAllPending(doc, useSingleTransaction: isBulkMode);

            SafeFileLogger.SafeAppendText("unified_workflow.log", $"[{DateTime.Now}] Workflow Complete.\n");
        }

        private int GetFilterIdForZone(ClashZone z)
        {
            // Helper if needed, but we do lookups inside the loop anyway.
            return 0; 
        }
    }
}
