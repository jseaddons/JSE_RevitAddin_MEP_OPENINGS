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
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase3And4.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase3And4.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services; // For OptimizationFlags, DebugLogger // Assuming typical interface namespace
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

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
                        var td = new Autodesk.Revit.UI.TaskDialog("Combined Sleeve Scope");
                        td.MainInstruction = "Select Scope for Auto-Joining";
                        td.MainContent = "Do you want to process the entire model or only elements within the active 3D Section Box?";
                        td.CommonButtons = Autodesk.Revit.UI.TaskDialogCommonButtons.Cancel;
                        
                        td.AddCommandLink(Autodesk.Revit.UI.TaskDialogCommandLinkId.CommandLink1, "Entire Model", "Process all sleeves in the project.");
                        td.AddCommandLink(Autodesk.Revit.UI.TaskDialogCommandLinkId.CommandLink2, "Active Section Box", "Process only sleeves inside the current Section Box.");

                        var result = td.Show();

                        if (result == Autodesk.Revit.UI.TaskDialogResult.Cancel) return;

                        if (result == Autodesk.Revit.UI.TaskDialogResult.CommandLink2)
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
                                Autodesk.Revit.UI.TaskDialog.Show("Warning", "The Section Box is NOT active in the current 3D view.\n\nPlease enable the Section Box in the Properties panel and try again.");
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
                    var persistenceService = new CombinedClusterPersistenceService(clashZoneRepository, combinedSleeveRepository, _document); // ✅ Added _document
                    
                    // ✅ INSTANTIATE SNAPSHOT SERVICES
                    var snapshotTransfer = new ParameterSnapshotTransferService();
                    var paramService = new JSE_RevitAddin_MEP_OPENINGS.Services.Placement.SleeveParameterService(_document);
                    var newSleeves = new List<Element>();

                    int createdCount = 0;
                    
                    using (var tx = new Transaction(_document, "Create Combined Sleeves"))
                    {
                        if (tx.Start() != TransactionStatus.Started) return;
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
                                int instanceId = instance.Id.GetIntegerValue(); // Keeping as int if GetIntegerValue returns int, but CombinedInstanceId is long
                                long instanceIdLong = instance.Id.GetIdValue(); 

                                newSleeves.Add(instance); // ✅ Track for parameter transfer

                                // E. Persist Link
                                persistenceService.PersistCombinedCluster(candidate, instanceIdLong);

                                createdCount++;
                            }
                            catch (Exception innerEx)
                            {
                                DebugLogger.Error($"[CombinedSleeveManager] Failed to create specific cluster: {innerEx.Message}");
                            }
                        }

                        // ✅ TRANSFER SNAPSHOT PARAMETERS (if enabled)
                        if (OptimizationFlags.EnableSnapshotParameterTransfer && newSleeves.Count > 0)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[CombinedSleeveManager] 🔄 Transferring snapshot parameters for {newSleeves.Count} combined sleeves...");
                                
                            snapshotTransfer.TransferSnapshotParameters(_document, newSleeves.Select(e => e.Id), clashZoneRepository, paramService);
                            paramService.FlushDeferredParameters();
                        }

                        tx.Commit();
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[CombinedSleeveManager] Success. Created {createdCount} combined sleeves.");
                        }
                    }
                    
                    // ✅ CRITICAL FIX: Delete constituent sleeves after combined sleeves are created
                    // This must happen OUTSIDE the creation transaction to avoid conflicts
                    if (createdCount > 0 && combinedCandidates.Count > 0)
                    {
                        DeleteConstituentSleeves(combinedCandidates);
                        
                        // ✅ Save snapshots for parameter transfer
                        SaveCombinedSleeveSnapshots(combinedCandidates);
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
        
        /// <summary>
        /// ✅ CRITICAL FIX: Deletes constituent sleeves (cluster and individual) after combined sleeves are created.
        /// This prevents duplicate sleeves in the model - the combined sleeve replaces the constituents.
        /// </summary>
        private void DeleteConstituentSleeves(List<CombinedClusterCandidate> combinedCandidates)
        {
            if (combinedCandidates == null || combinedCandidates.Count == 0) return;
            
            var elementIdsToDelete = new HashSet<long>();
            
            // Collect all constituent element IDs from all combined candidates
            foreach (var candidate in combinedCandidates)
            {
                if (candidate?.MemberClusters == null) continue;
                
                foreach (var member in candidate.MemberClusters)
                {
                    if (member == null) continue;
                    
                    // Cluster sleeves have ClusterSleeveInstanceId
                    if (member.ClusterSleeveInstanceId > 0)
                    {
                        elementIdsToDelete.Add(member.ClusterSleeveInstanceId);
                    }
                    
                    // Individual sleeves - collect ClashZoneIds to resolve later
                    // (Individual sleeves are tracked via ClashZone -> SleeveInstanceId mapping)
                }
            }
            
            // Resolve individual sleeve IDs from database
            try
            {
                using (var dbContext = new SleeveDbContext(_document))
                {
                    // Get all clash zones that are part of combined sleeves
                    var allZoneGuids = combinedCandidates
                        .SelectMany(c => c.MemberClusters ?? new List<ClusterSleeveInfo>())
                        .Where(m => m != null && m.ClashZoneIds != null)
                        .SelectMany(m => m.ClashZoneIds)
                        .Distinct()
                        .ToList();
                    
                    if (allZoneGuids.Count > 0)
                    {
                        using (var cmd = dbContext.Connection.CreateCommand())
                        {
                            var guidPlaceholders = string.Join(",", allZoneGuids.Select((_, i) => $"@g{i}"));
                            cmd.CommandText = $@"
                                SELECT SleeveInstanceId 
                                FROM ClashZones 
                                WHERE SleeveInstanceId > 0 
                                AND ClashZoneGuid IN ({guidPlaceholders})";
                            
                            for (int i = 0; i < allZoneGuids.Count; i++)
                            {
                                cmd.Parameters.AddWithValue($"@g{i}", allZoneGuids[i].ToString());
                            }
                            
                            using (var reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    long sleeveId = reader.GetInt64(0);
                                    if (sleeveId > 0)
                                    {
                                        elementIdsToDelete.Add(sleeveId);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[CombinedSleeveManager] Failed to resolve individual sleeve IDs: {ex.Message}");
                }
            }
            
            if (elementIdsToDelete.Count == 0)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info("[CombinedSleeveManager] No constituent sleeves to delete.");
                }
                return;
            }
            
            // Delete the constituent elements in a new transaction
            try
            {
                using (var tx = new Transaction(_document, "Delete Combined Constituents"))
                {
                    if (tx.Start() == TransactionStatus.Started)
                    {
                        // Filter to only elements that still exist
                        var revitIds = elementIdsToDelete
                            .Where(id => id > 0)
#if REVIT2024_OR_GREATER
                            .Select(id => new ElementId(id))
#else
                            .Select(id => new ElementId((int)id))
#endif

                            .Where(id => _document.GetElement(id) != null)
                            .ToList();
                        
                        if (revitIds.Count > 0)
                        {
                            _document.Delete(revitIds);
                            tx.Commit();
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[CombinedSleeveManager] 🧹 Deleted {revitIds.Count} constituent sleeves after combining.");
                            }
                        }
                        else
                        {
                            tx.RollBack();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[CombinedSleeveManager] Failed to delete constituent sleeves: {ex.Message}");
                }
            }
        }
        
        /// <summary>
        /// ✅ CRITICAL FIX: Saves snapshots for combined sleeves to enable parameter transfer.
        /// This stores MEP parameters from constituent zones so they can be applied to combined sleeves.
        /// </summary>
        private void SaveCombinedSleeveSnapshots(List<CombinedClusterCandidate> combinedCandidates)
        {
            if (combinedCandidates == null || combinedCandidates.Count == 0) return;
            
            try
            {
                using (var dbContext = new SleeveDbContext(_document))
                using (var transaction = dbContext.Connection.BeginTransaction())
                {
                    foreach (var candidate in combinedCandidates)
                    {
                        if (candidate?.CombinedClusterInstanceId <= 0) continue;
                        if (candidate.MemberClusters == null || candidate.MemberClusters.Count == 0) continue;
                        
                        var mepElementIds = new List<long>();
                        var hostElementIds = new List<long>();
                        var mepParams = new Dictionary<string, object>();
                        var allZoneGuids = new List<Guid>();
                        
                        // Collect data from all member clusters
                        foreach (var member in candidate.MemberClusters)
                        {
                            if (member?.ClashZoneIds == null) continue;
                            allZoneGuids.AddRange(member.ClashZoneIds);
                            if (member.ClusterSleeveInstanceId > 0)
                            {
                                mepElementIds.Add(member.ClusterSleeveInstanceId);
                            }
                        }
                        
                        // Get MEP parameters from ClashZones
                        if (allZoneGuids.Count > 0)
                        {
                            try
                            {
                                using (var cmd = dbContext.Connection.CreateCommand())
                                {
                                    var guidPlaceholders = string.Join(",", allZoneGuids.Select((_, i) => $"@g{i}"));
                                    cmd.CommandText = $@"
                                        SELECT DISTINCT MepElementId, StructuralElementId, MepParameterValuesJson
                                        FROM ClashZones 
                                        WHERE ClashZoneGuid IN ({guidPlaceholders})";
                                    
                                    for (int i = 0; i < allZoneGuids.Count; i++)
                                    {
                                        cmd.Parameters.AddWithValue($"@g{i}", allZoneGuids[i].ToString());
                                    }
                                    
                                    using (var reader = cmd.ExecuteReader())
                                    {
                                        while (reader.Read())
                                        {
                                            if (!reader.IsDBNull(0))
                                                mepElementIds.Add(reader.GetInt64(0));
                                            if (!reader.IsDBNull(1))
                                                hostElementIds.Add(reader.GetInt64(1));
                                            
                                            // Parse MEP parameters from JSON
                                            if (!reader.IsDBNull(2))
                                            {
                                                try
                                                {
                                                    var json = reader.GetString(2);
                                                    if (!string.IsNullOrEmpty(json))
                                                    {
                                                        var values = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, object>>(json);
                                                        if (values != null)
                                                        {
                                                            foreach (var kvp in values)
                                                            {
                                                                if (!mepParams.ContainsKey(kvp.Key))
                                                                    mepParams[kvp.Key] = kvp.Value;
                                                            }
                                                        }
                                                    }
                                                }
                                                catch { /* Ignore JSON parse errors */ }
                                            }
                                        }
                                    }
                                }
                            }
                            catch { /* Ignore lookup errors */ }
                        }
                        
                        // Add combined sleeve metadata
                        mepParams["CombinedWidth"] = candidate.CombinedWidth;
                        mepParams["CombinedHeight"] = candidate.CombinedHeight;
                        mepParams["CombinedDepth"] = candidate.CombinedHeight; // Use Height as Depth for 2D approximation
                        mepParams["Categories"] = string.Join(",", candidate.CategoriesInvolved ?? new List<string>());
                        mepParams["ConstituentCount"] = candidate.MemberClusters?.Count ?? 0;
                        mepParams["IsCombinedSleeve"] = true;
                        mepParams["IsAutoCreated"] = true;
                        
                        // Insert snapshot
                        using (var cmd = dbContext.Connection.CreateCommand())
                        {
                            cmd.Transaction = transaction;
                            // ✅ FIX: Use INSERT OR REPLACE to prevent duplicate rows for same CombinedInstanceId
                            cmd.CommandText = @"
                                INSERT OR REPLACE INTO SleeveSnapshots (
                                    CombinedInstanceId, SourceType, ClashZoneGuid,
                                    MepElementIdsJson, HostElementIdsJson,
                                    MepParametersJson, HostParametersJson,
                                    FilterId, ComboId, CreatedAt, UpdatedAt
                                ) VALUES (
                                    @CombinedInstanceId, @SourceType, @ClashZoneGuid,
                                    @MepElementIdsJson, @HostElementIdsJson,
                                    @MepParametersJson, @HostParametersJson,
                                    @FilterId, @ComboId, 
                                    COALESCE((SELECT CreatedAt FROM SleeveSnapshots WHERE CombinedInstanceId = @CombinedInstanceId), CURRENT_TIMESTAMP),
                                    CURRENT_TIMESTAMP
                                )";
                            
                            cmd.Parameters.AddWithValue("@CombinedInstanceId", candidate.CombinedClusterInstanceId);
                            cmd.Parameters.AddWithValue("@SourceType", "Combined");
                            cmd.Parameters.AddWithValue("@ClashZoneGuid", $"COMBINED_AUTO_{candidate.CombinedClusterInstanceId}");
                            cmd.Parameters.AddWithValue("@MepElementIdsJson", System.Text.Json.JsonSerializer.Serialize(mepElementIds.Distinct()));
                            cmd.Parameters.AddWithValue("@HostElementIdsJson", System.Text.Json.JsonSerializer.Serialize(hostElementIds.Distinct()));
                            cmd.Parameters.AddWithValue("@MepParametersJson", System.Text.Json.JsonSerializer.Serialize(mepParams));
                            cmd.Parameters.AddWithValue("@HostParametersJson", "{}");
                            cmd.Parameters.AddWithValue("@FilterId", DBNull.Value); // Combined spans filters
                            cmd.Parameters.AddWithValue("@ComboId", DBNull.Value);  // Combined spans combos
                            
                            cmd.ExecuteNonQuery();
                        }
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[CombinedSleeveManager] ✅ Saved snapshot for combined sleeve {candidate.CombinedClusterInstanceId}");
                        }
                    }
                    
                    transaction.Commit();
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[CombinedSleeveManager] Failed to save combined sleeve snapshots: {ex.Message}");
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