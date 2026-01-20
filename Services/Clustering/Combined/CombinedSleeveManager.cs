using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Repository;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase3And4.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services; // For OptimizationFlags, DebugLogger // Assuming typical interface namespace

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined
{
    /// <summary>
    /// Orchestrates the Combined Sleeve workflow: Discovery -> Formation -> Persistence -> Creation.
    /// Replaces the embedded logic previously in OpeningCommandOrchestrator.
    /// </summary>
    public class CombinedSleeveManager
    {
        private readonly Document _document;
        private readonly Autodesk.Revit.UI.UIDocument _uiDocument;

        public CombinedSleeveManager(Document document, Autodesk.Revit.UI.UIDocument uiDocument)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _uiDocument = uiDocument ?? throw new ArgumentNullException(nameof(uiDocument));
        }

        public void Execute(bool showProgress)
        {
            try
            {
                // Feature Flag Check
                if (!OptimizationFlags.UseCombinedClustering || !OptimizationFlags.UseCombinedClusteringPhase1And2)
                {
                    return;
                }

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info("[CombinedSleeveManager] 🔥 Starting Combined Clustering Workflow 🔥");
                }

                // Include duct accessories/damper so they can be clustered as well
                var defaultCategories = new List<string> { "Ducts", "Pipes", "CableTrays", "Conduit", "Duct Accessories" };
                var filterName = "*"; // ✅ Wildcard: scan all filters for valid cluster sleeves

                using (var dbContext = new SleeveDbContext(_document))
                {
                    // 1. Setup Data Access
                    var clashZoneRepository = new ClashZoneRepository(dbContext, msg => DebugLogger.Info(msg));
                    var repo = new CombinedClusterRepository(clashZoneRepository);
                    
                    // 1b. Determine Scope: Whole Model or Section Box
                    BoundingBoxXYZ? sectionBox = null;
                    bool useSectionBox = false;

                    // CHECK: Is user in a 3D view? If so, offer Section Box option.
                    if (_document.ActiveView is View3D view3D)
                    {
                        var td = new TaskDialog("Combined Sleeve Scope");
                        td.MainInstruction = "Select Scope for Auto-Joining";
                        td.MainContent = "Do you want to process the entire model or only elements within the active 3D Section Box?";
                        td.CommonButtons = TaskDialogCommonButtons.Cancel;
                        
                        td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Entire Model", "Process all sleeves in the project.");
                        td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Active Section Box", "Process only sleeves inside the current Section Box.");

                        var result = td.Show();

                        if (result == TaskDialogResult.Cancel) return;

                        if (result == TaskDialogResult.CommandLink2)
                        {
                            if (view3D.IsSectionBoxActive)
                            {
                                sectionBox = view3D.GetSectionBox();
                                useSectionBox = true;
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[CombinedSleeveManager] 📦 Filtering by active section box: {sectionBox.Min} to {sectionBox.Max}");
                                }
                            }
                            else
                            {
                                TaskDialog.Show("Warning", "The Section Box is NOT active in the current 3D view.\n\nPlease enable the Section Box in the Properties panel and try again.");
                                return;
                            }
                        }
                    }
                    else
                    {
                        // Not in 3D view -> Default to entire model (or could prompt confirming that)
                        // For now, proceed with entire model to avoid blocking non-3D workflows unless explicitly requested.
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info("[CombinedSleeveManager] Not in 3D View - processing entire model.");
                    }

                    // 2. Discovery Phase
                    var discoveryService = new CombinedClusterDiscoveryService(repo);
                    // Explicitly qualify if needed, but using directives should handle it
                    var formationService = new JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase3And4.Services.CombinedClusterFormationService(repo);

                    var sleeves = discoveryService.Discover(_uiDocument, filterName, defaultCategories, sectionBox);
                    if (sleeves == null || sleeves.Count == 0)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info("[CombinedSleeveManager] No cluster sleeves discovered.");
                        }
                        return;
                    }

                    // 3. Formation Phase (Grouping)
                    // Note: .Result usage is acceptable here if called from a standard external command context 
                    // where async/await is not fully propagated up to IExternalCommand.
                    var combinedCandidates = formationService.FormCombinedClustersAsync(sleeves.ToList()).Result;

                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[CombinedSleeveManager] Phase 1-2 produced {combinedCandidates.Count} candidate groups.");
                    }

                    // 4. Persistence & Creation Phase (Phase 3 & 4)
                    if (!OptimizationFlags.UseCombinedClusteringPhase3And4)
                    {
                        return;
                    }

                    var combinedSleeveRepository = new CombinedSleeveRepository(dbContext); // ✅ Created repository
                    var paramAggregator = new ParameterAggregatorService();
                    var persistenceService = new CombinedClusterPersistenceService(clashZoneRepository, combinedSleeveRepository); // ✅ Passed to constructor

                    using (var tx = new Transaction(_document, "Create Combined Sleeves"))
                    {
                        if (tx.Start() != TransactionStatus.Started) return;

                        int createdCount = 0;
                        var failureOptions = tx.GetFailureHandlingOptions();
                        failureOptions.SetFailuresPreprocessor(new CombinedSleeveWarningSwallower());
                        tx.SetFailureHandlingOptions(failureOptions);

                        foreach (var candidate in combinedCandidates)
                        {
                            try
                            {
                                // A. Prepare Parameters & Geometry
                                var bbox = new BoundingBoxXYZ();
                                bbox.Min = new XYZ(candidate.CombinedBoundingBoxMinX, candidate.CombinedBoundingBoxMinY, candidate.CombinedBoundingBoxMinZ);
                                bbox.Max = new XYZ(candidate.CombinedBoundingBoxMaxX, candidate.CombinedBoundingBoxMaxY, candidate.CombinedBoundingBoxMaxZ);

                                var paramSet = paramAggregator.CreateCombinedSleeveParameterSet(candidate, bbox);

                                // B. Determine Family Name
                                // Always place opening families here; dampers are handled by dedicated placement services
                                // ✅ CLUSTER FIX: Use dynamic family selection instead of hardcoded Wall family
                                string familyName = JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Placement.ClusterPlacementService.GetFamilyName(
                                    candidate.HostType ?? "Wall", 
                                    candidate.CategoriesInvolved?.FirstOrDefault() ?? "Unknown", 
                                    Math.Max(candidate.CombinedWidth, candidate.CombinedHeight), 
                                    true); // isCluster = true
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[CombinedSleeveManager] Creating '{familyName}' for {string.Join(",", candidate.CategoriesInvolved)}");

                                // C. Load/Activate Symbol
                                FamilySymbol symbol = new FilteredElementCollector(_document)
                                    .OfClass(typeof(FamilySymbol))
                                    .Cast<FamilySymbol>()
                                    .FirstOrDefault(x => x.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase) ||
                                                         x.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase));

                                if (symbol == null)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Warning($"[CombinedSleeveManager] Family symbol '{familyName}' not found. Skipping.");
                                    continue;
                                }

                                if (!symbol.IsActive) symbol.Activate();

                                // D. Create Instance
                                var instance = _document.Create.NewFamilyInstance(bbox.Min, symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                int instanceId = instance.Id.IntegerValue;

                                // E. Persist Link
                                persistenceService.PersistCombinedCluster(candidate, instanceId);

                                createdCount++;
                            }
                            catch (Exception innerEx)
                            {
                                DebugLogger.Error($"[CombinedSleeveManager] Failed to create specific cluster: {innerEx.Message}");
                            }
                        }

                        tx.Commit();
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[CombinedSleeveManager] Success. Created {createdCount} combined sleeves.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[CombinedSleeveManager] Critical Failure: {ex.Message}\n{ex.StackTrace}");
                }
            }
        }
    }

    /// <summary>
    /// Dedicated failure preprocessor for combined sleeve creation warnings.
    /// </summary>
    public class CombinedSleeveWarningSwallower : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(FailuresAccessor fa)
        {
            var failures = fa.GetFailureMessages();
            foreach (var f in failures)
            {
                var severity = f.GetSeverity();
                if (severity == FailureSeverity.Warning)
                {
                    fa.DeleteWarning(f);
                }
                else if (severity == FailureSeverity.Error)
                {
                    // Check for harmless errors like "duplicate instances" which might happen in edge cases
                    var desc = f.GetDescriptionText();
                    if (desc.Contains("duplicate", StringComparison.OrdinalIgnoreCase))
                    {
                        fa.DeleteWarning(f); // Treating as warning to proceed
                    }
                }
            }
            return FailureProcessingResult.Continue;
        }
    }
}
