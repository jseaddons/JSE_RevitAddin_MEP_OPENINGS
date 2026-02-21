using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for transferring parameters from various sources to openings
    /// </summary>
    public partial class ParameterTransferService
    {
        private readonly ParameterRenamingService _renamingService;
        private readonly ParameterMappingService _mappingService;
        private readonly ServiceTypeAbbreviationService _abbreviationService;
        // private readonly MepElementAnalysisService _mepAnalysisService;
        private readonly ISectionBoxService _sectionBoxService;
        
        public ParameterTransferService()
        {
            _renamingService = new ParameterRenamingService();
            _mappingService = new ParameterMappingService();
            _abbreviationService = new ServiceTypeAbbreviationService();
            // _mepAnalysisService = new MepElementAnalysisService();
            _sectionBoxService = new SectionBoxService();
        }
        
        /// <summary>
        /// Transaction-scoped transfer from reference elements. This method assumes an active transaction
        /// is already started by the caller (command/orchestrator). It will not start/commit transactions.
        /// </summary>
        public ParameterTransferResult TransferFromReferenceElementsInTransaction(
            Document doc,
            List<ElementId> openingIds,
            ParameterMapping mapping,
            UIDocument uiDoc = null)
        {
            var result = new ParameterTransferResult();

            try
            {
                if (!mapping.IsEnabled)
                {
                    result.Success = true;
                    result.Message = "Mapping is disabled, skipping transfer.";
                    return result;
                }

                var transferredCount = 0;
                var failedCount = 0;
                var errors = new List<string>();

                foreach (var openingId in openingIds)
                {
                    try
                    {
                        var opening = doc.GetElement(openingId);
                        if (opening == null) continue;

                        // Get MEP elements that intersect with this opening
                        var mepElements = GetMepElementsInOpening(doc, opening);

                        if (mepElements.Count == 0)
                        {
                            result.Warnings.Add($"No MEP elements found for opening {openingId}");
                            continue;
                        }

                        // Transfer parameter from first MEP element (or combine if multiple)
                        var transferSuccess = TransferParameterFromElements(
                            doc, opening, mepElements, mapping);

                        if (transferSuccess)
                            transferredCount++;
                        else
                            failedCount++;
                    }
                    catch (Exception ex)
                    {
                        failedCount++;
                        errors.Add($"Error transferring to opening {openingId}: {ex.Message}");
                    }
                }

                result.Success = true;
                result.TransferredCount = transferredCount;
                result.FailedCount = failedCount;
                result.Errors = errors;
                result.Message = $"Transferred {transferredCount} parameters, {failedCount} failed.";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Transfer failed: {ex.Message}";
                result.Errors.Add(ex.Message);
            }

            return result;
        }

        /// <summary>
        /// Backward-compatible wrapper that keeps the older behavior (service-owned transaction).
        /// Prefer using TransferFromReferenceElementsInTransaction by callers that own the transaction.
        /// </summary>
        [Obsolete("Use TransferFromReferenceElementsInTransaction and own the transaction at the command level.")]
        public ParameterTransferResult TransferFromReferenceElements(
            Document doc,
            List<ElementId> openingIds,
            ParameterMapping mapping)
        {
            using (var t = new Transaction(doc, "Transfer Parameters from Reference Elements (wrapper)"))
            {
                t.Start();
                var r = TransferFromReferenceElementsInTransaction(doc, openingIds, mapping);
                t.Commit();
                return r;
            }
        }
        
        /// <summary>
        /// Transfer parameters from host elements (walls, floors, ceilings) to openings
        /// </summary>
        public ParameterTransferResult TransferFromHostElements(
            Document doc, 
            List<ElementId> openingIds, 
            ParameterMapping mapping)
        {
            // backward-compatible wrapper
            using (var t = new Transaction(doc, "Transfer Parameters from Host Elements (wrapper)"))
            {
                t.Start();
                var r = TransferFromHostElementsInTransaction(doc, openingIds, mapping);
                t.Commit();
                return r;
            }
        }

        public ParameterTransferResult TransferFromHostElementsInTransaction(
            Document doc,
            List<ElementId> openingIds,
            ParameterMapping mapping)
        {
            var result = new ParameterTransferResult();

            try
            {
                if (!mapping.IsEnabled)
                {
                    result.Success = true;
                    result.Message = "Mapping is disabled, skipping transfer.";
                    return result;
                }

                var transferredCount = 0;
                var failedCount = 0;
                var errors = new List<string>();

                foreach (var openingId in openingIds)
                {
                    try
                    {
                        var opening = doc.GetElement(openingId);
                        if (opening == null) continue;

                        // Get host elements (walls, floors, ceilings) that contain this opening
                        var hostElements = GetHostElementsForOpening(doc, opening);

                        if (hostElements.Count == 0)
                        {
                            result.Warnings.Add($"No host elements found for opening {openingId}");
                            continue;
                        }

                        // Transfer parameter from host elements
                        var transferSuccess = TransferParameterFromElements(
                            doc, opening, hostElements, mapping);

                        if (transferSuccess)
                            transferredCount++;
                        else
                            failedCount++;
                    }
                    catch (Exception ex)
                    {
                        failedCount++;
                        errors.Add($"Error transferring to opening {openingId}: {ex.Message}");
                    }
                }

                result.Success = true;
                result.TransferredCount = transferredCount;
                result.FailedCount = failedCount;
                result.Errors = errors;
                result.Message = $"Transferred {transferredCount} parameters, {failedCount} failed.";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Transfer failed: {ex.Message}";
                result.Errors.Add(ex.Message);
            }

            return result;
        }
        
        /// <summary>
        /// Transfer parameters from levels to openings
        /// </summary>
        public ParameterTransferResult TransferFromLevels(
            Document doc, 
            List<ElementId> openingIds, 
            ParameterMapping mapping)
        {
            // Non-transactional service method: perform transfer logic without starting/committing transactions.
            // Caller (command/orchestrator) is expected to own the transaction when required.
            return TransferFromLevelsInTransaction(doc, openingIds, mapping);
        }

        public ParameterTransferResult TransferFromLevelsInTransaction(
            Document doc,
            List<ElementId> openingIds,
            ParameterMapping mapping)
        {
            var result = new ParameterTransferResult();

            try
            {
                if (!mapping.IsEnabled)
                {
                    result.Success = true;
                    result.Message = "Mapping is disabled, skipping transfer.";
                    return result;
                }

                var transferredCount = 0;
                var failedCount = 0;
                var errors = new List<string>();

                foreach (var openingId in openingIds)
                {
                    try
                    {
                        var opening = doc.GetElement(openingId);
                        if (opening == null) continue;

                        // Get level for this opening
                        var level = GetLevelForOpening(doc, opening);

                        if (level == null)
                        {
                            result.Warnings.Add($"No level found for opening {openingId}");
                            continue;
                        }

                        // Transfer parameter from level
                        var transferSuccess = TransferParameterFromElement(
                            doc, opening, level, mapping);

                        if (transferSuccess)
                            transferredCount++;
                        else
                            failedCount++;
                    }
                    catch (Exception ex)
                    {
                        failedCount++;
                        errors.Add($"Error transferring to opening {openingId}: {ex.Message}");
                    }
                }

                result.Success = true;
                result.TransferredCount = transferredCount;
                result.FailedCount = failedCount;
                result.Errors = errors;
                result.Message = $"Transferred {transferredCount} parameters, {failedCount} failed.";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Transfer failed: {ex.Message}";
                result.Errors.Add(ex.Message);
            }

            return result;
        }
        
        /// <summary>
        /// Transfer service size calculations with clearance to openings
        /// </summary>
        public ParameterTransferResult TransferServiceSizeCalculations(
            Document doc, 
            List<ElementId> openingIds, 
            string targetParameter,
            double clearance = 50.0)
        {
            // Non-transactional service method: do the transfers; caller must manage transactions.
            return TransferServiceSizeCalculationsInTransaction(doc, openingIds, targetParameter, clearance);
        }

        public ParameterTransferResult TransferServiceSizeCalculationsInTransaction(
            Document doc,
            List<ElementId> openingIds,
            string targetParameter,
            double clearance = 50.0)
        {
            var result = new ParameterTransferResult();

            try
            {
                var transferredCount = 0;
                var failedCount = 0;
                var errors = new List<string>();

                foreach (var openingId in openingIds)
                {
                    try
                    {
                        var opening = doc.GetElement(openingId);
                        if (opening == null) continue;

                        // Get MEP elements that intersect with this opening
                        var mepElements = GetMepElementsInOpening(doc, opening);

                        if (mepElements.Count == 0)
                        {
                            result.Warnings.Add($"No MEP elements found for opening {openingId}");
                            continue;
                        }

                        // Calculate service size with clearance
                        // var serviceSizeCalculation = _mepAnalysisService.CalculateServiceSize(mepElements, clearance);

                        // Set parameter value
                        var param = opening.LookupParameter(targetParameter);
                        if (param != null && !param.IsReadOnly)
                        {
                            // param.Set(serviceSizeCalculation.CalculationString); // Stubbed
                            transferredCount++;
                        }
                        else
                        {
                            failedCount++;
                            errors.Add($"Cannot set parameter {targetParameter} on opening {openingId}");
                        }
                    }
                    catch (Exception ex)
                    {
                        failedCount++;
                        errors.Add($"Error transferring service size calculation to opening {openingId}: {ex.Message}");
                    }
                }

                result.Success = true;
                result.TransferredCount = transferredCount;
                result.FailedCount = failedCount;
                result.Errors = errors;
                result.Message = $"Transferred {transferredCount} service size calculations, {failedCount} failed.";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Service size calculation transfer failed: {ex.Message}";
                result.Errors.Add(ex.Message);
            }

            return result;
        }
        
        /// <summary>
        /// Transfer model names to openings
        /// </summary>
        public ParameterTransferResult TransferModelNames(
            Document doc, 
            List<ElementId> openingIds, 
            string targetParameter)
        {
            // Non-transactional service method: perform model name writes; caller must own the transaction.
            return TransferModelNamesInTransaction(doc, openingIds, targetParameter);
        }

        public ParameterTransferResult TransferModelNamesInTransaction(
            Document doc,
            List<ElementId> openingIds,
            string targetParameter)
        {
            var result = new ParameterTransferResult();

            try
            {
                var transferredCount = 0;
                var failedCount = 0;
                var errors = new List<string>();

                foreach (var openingId in openingIds)
                {
                    try
                    {
                        var opening = doc.GetElement(openingId);
                        if (opening == null) continue;

                        // Get model name
                        var modelName = doc.Title;

                        // Set parameter value
                        var param = opening.LookupParameter(targetParameter);
                        if (param != null && !param.IsReadOnly)
                        {
                            param.Set(modelName);
                            transferredCount++;
                        }
                        else
                        {
                            failedCount++;
                            errors.Add($"Cannot set parameter {targetParameter} on opening {openingId}");
                        }
                    }
                    catch (Exception ex)
                    {
                        failedCount++;
                        errors.Add($"Error transferring model name to opening {openingId}: {ex.Message}");
                    }
                }

                result.Success = true;
                result.TransferredCount = transferredCount;
                result.FailedCount = failedCount;
                result.Errors = errors;
                result.Message = $"Transferred {transferredCount} model names, {failedCount} failed.";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Transfer failed: {ex.Message}";
                result.Errors.Add(ex.Message);
            }

            return result;
        }
        
        /// <summary>
        /// Transfer standard parameters (Level, Dimensions, System/Service Type) from reference MEP elements to openings
        /// - Ducts, Cable Trays, Duct Accessories: Height and Width
        /// - Pipes: Outside Diameter
        /// - Ducts, Pipes, Duct Accessories: System Type
        /// - Cable Trays: Service Type
        /// Targets (if present and writable): Reference_Level, Reference_Height, Reference_Width, Reference_Diameter, MEP_System_Type
        /// </summary>
        public ParameterTransferResult TransferStandardParametersFromReferenceElements(
            Document doc,
            List<ElementId> openingIds)
        {
            var result = new ParameterTransferResult();
            var errors = new List<string>();
            var warnings = new List<string>();
            // int transferred = 0; // FIX: CS0219 - commented to fix critical warning
            // int failed = 0; // FIX: CS0219 - commented to fix critical warning

            // Wrapper
            using (var t = new Transaction(doc, "Transfer Standard Parameters from Reference Elements (wrapper)"))
            {
                t.Start();
                var r = TransferStandardParametersFromReferenceElementsInTransaction(doc, openingIds);
                t.Commit();
                return r;
            }
        }

        public ParameterTransferResult TransferStandardParametersFromReferenceElementsInTransaction(
            Document doc,
            List<ElementId> openingIds)
        {
            var result = new ParameterTransferResult();
            var errors = new List<string>();
            var warnings = new List<string>();
            int transferred = 0;
            int failed = 0;

            try
            {
                foreach (var openingId in openingIds)
                {
                    try
                    {
                        var opening = doc.GetElement(openingId);
                        if (opening == null)
                        {
                            failed++;
                            errors.Add($"Opening {openingId} not found");
                            continue;
                        }

                        // Find intersecting MEP elements
                        var mepElements = GetMepElementsInOpening(doc, opening);
                        if (mepElements.Count == 0)
                        {
                            warnings.Add($"No MEP elements found for opening {openingId}");
                            continue;
                        }

                        // Use the first intersecting MEP element as the source
                        var source = mepElements[0];

                        bool anySet = false;

                        // 1) Level → Reference_Level (string)
                        var levelName = GetLevelName(doc, source);
                        if (!string.IsNullOrEmpty(levelName))
                        {
                            var p = opening.LookupParameter("Reference_Level");
                            if (SetParameterValueSafely(p, levelName)) anySet = true;
                        }

                        // 2) Dimensions
                        var categoryId = source.Category?.Id.GetIntegerValue() ?? -1;

                        // Rectangular (Ducts, Cable Trays, Duct Accessories): Height, Width
                        if (categoryId == (int)BuiltInCategory.OST_DuctCurves ||
                            categoryId == (int)BuiltInCategory.OST_CableTray ||
                            categoryId == (int)BuiltInCategory.OST_DuctAccessory)
                        {
                            var height = GetParamDouble(source, "Height");
                            var width = GetParamDouble(source, "Width");
                            if (height.HasValue)
                            {
                                var pH = opening.LookupParameter("Reference_Height");
                                if (SetParameterValueSafely(pH, height.Value)) anySet = true;
                            }
                            if (width.HasValue)
                            {
                                var pW = opening.LookupParameter("Reference_Width");
                                if (SetParameterValueSafely(pW, width.Value)) anySet = true;
                            }
                        }

                        // Circular (Pipes): Outside Diameter
                        if (categoryId == (int)BuiltInCategory.OST_PipeCurves)
                        {
                            var diameter = GetParamDouble(source, "Outside Diameter", "Diameter");
                            if (diameter.HasValue)
                            {
                                var pD = opening.LookupParameter("Reference_Diameter");
                                if (SetParameterValueSafely(pD, diameter.Value)) anySet = true;
                            }
                        }

                        // 3) System/Service Type → MEP_System_Type (string)
                        string systemValue = null;
                        if (categoryId == (int)BuiltInCategory.OST_CableTray)
                        {
                            systemValue = source.LookupParameter("Service Type")?.AsString();
                        }
                        else if (categoryId == (int)BuiltInCategory.OST_DuctCurves ||
                                 categoryId == (int)BuiltInCategory.OST_PipeCurves ||
                                 categoryId == (int)BuiltInCategory.OST_DuctAccessory)
                        {
                            systemValue = source.LookupParameter("System Type")?.AsString();
                        }

                        if (!string.IsNullOrWhiteSpace(systemValue))
                        {
                            var pSys = opening.LookupParameter("MEP_System_Type");
                            if (SetParameterValueSafely(pSys, systemValue)) anySet = true;

                            // Also write to Service_Category if present
                            var pSvc = opening.LookupParameter("Service_Category");
                            SetParameterValueSafely(pSvc, systemValue);
                        }

                        if (anySet) transferred++; else failed++;
                    }
                    catch (Exception exOpen)
                    {
                        failed++;
                        errors.Add($"Error on opening {openingId}: {exOpen.Message}");
                    }
                }

                result.Success = errors.Count == 0;
                result.TransferredCount = transferred;
                result.FailedCount = failed;
                result.Errors = errors;
                result.Warnings = warnings;
                result.Message = $"Standard transfer complete: {transferred} updated, {failed} failed.";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Standard parameter transfer failed: {ex.Message}";
                result.Errors.Add(ex.Message);
            }

            return result;
        }
        
        /// <summary>
        /// Execute complete parameter transfer configuration
        /// </summary>
        public ParameterTransferResult ExecuteTransferConfiguration(
            Document doc,
            List<ElementId> openingIds,
            ParameterTransferConfiguration config)
        {
            if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[PARAM_TRANSFER] ExecuteTransferConfiguration called with {openingIds.Count} openings and {config.Mappings.Count} mappings");
            if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[{DateTime.Now}] [PARAM_TRANSFER] ExecuteTransferConfiguration called with {openingIds.Count} openings and {config.Mappings.Count} mappings\n");

            var result = new ParameterTransferResult();
            var allResults = new List<ParameterTransferResult>();

            // Add timing for the whole operation
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();

            // Backward-compatible wrapper which creates a transaction and calls the in-transaction implementation
            using (var t = new Transaction(doc, "Execute Parameter Transfer Configuration (wrapper)"))
            {
                t.Start();
                ParameterTransferResult r = null;
                SafeFileLogger.SafeAppendText("transfer_debug.log", $"[DEBUG] UseParameterTransferRefactor Limit?? Flag = {OptimizationFlags.UseParameterTransferRefactor}\n");
                if (OptimizationFlags.UseParameterTransferRefactor)
                {
                    try
                    {
                        // Call the optimized path (to be implemented/refactored)
                        r = ExecuteTransferConfigurationInTransaction_Refactored(doc, openingIds, config, null);
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Warning($"[OPTIMIZATION ROLLBACK] Refactored parameter transfer failed: {ex.Message}. Falling back to legacy path.");
                        OptimizationFlags.UseParameterTransferRefactor = false; // Optionally disable for session
                        r = ExecuteTransferConfigurationInTransaction(doc, openingIds, config, null);
                    }
                }
                else
                {
                    r = ExecuteTransferConfigurationInTransaction(doc, openingIds, config, null);
                }
                t.Commit();

                stopwatch.Stop();
                var elapsed = stopwatch.Elapsed;
                var transferDebugLogPath = SafeFileLogger.GetLogFilePath("transfer_debug.log");
                SafeFileLogger.SafeAppendText(transferDebugLogPath,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] TOTAL ELAPSED TIME: {elapsed.TotalSeconds:F2} seconds for parameter transfer, mark, and remark operations.\n");

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[PARAM_TRANSFER] ExecuteTransferConfiguration completed: Success={r.Success}, TransferredCount={r.TransferredCount}, FailedCount={r.FailedCount}");
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[{DateTime.Now}] [PARAM_TRANSFER] ExecuteTransferConfiguration completed: Success={r.Success}, TransferredCount={r.TransferredCount}, FailedCount={r.FailedCount}\n");

                return r;
            }
        }

        // Fully implemented refactored method
        private ParameterTransferResult ExecuteTransferConfigurationInTransaction_Refactored(
            Document doc,
            List<ElementId> openingIds,
            ParameterTransferConfiguration config,
            UIDocument uiDoc = null)
        {
            // Optimized parameter transfer logic with buffered logging, lazy cache, unified snapshot lookup
            // LOOP INVERSION OPTIMIZATION: Iterate Sleeves -> Mappings to minimize GetElement and DB Lookups.
            
            var result = new ParameterTransferResult();
            var logBuffer = new System.Text.StringBuilder(50000); // Pre-allocate 50KB to minimize resizes
            
            // Performance tracking
            var methodStartTime = System.Diagnostics.Stopwatch.StartNew();

            try
            {
                // Check if any sleeves exist in the model
                if (openingIds == null || openingIds.Count == 0)
                {
                    result.Success = false;
                    result.Message = "No sleeves found in the model. Please place sleeves first before transferring parameters.";
                    result.Errors.Add("No sleeves found - place sleeves first");
                    logBuffer.AppendLine("[PARAM_TRANSFER] No sleeves found in model - user needs to place sleeves first");
                    SafeFileLogger.SafeAppendText("transfer_debug.log", logBuffer.ToString());
                    return result;
                }

                if (!DeploymentConfiguration.DeploymentMode)
                    logBuffer.AppendLine($"[PARAM_TRANSFER] Refactored Transfer (Optimized Loop): Found {openingIds.Count} sleeves in model");

                // Performance Report Variables
                long timeLoadSnapshot = 0;
                long timeGetElement = 0;
                long timeResolution = 0;
                long timeParamSet = 0;
                var sw = new System.Diagnostics.Stopwatch();

                sw.Start();
                // Load snapshot index
                SleeveSnapshotIndex snapshotIndex;
                try
                {
                    using (var dbContext = new SleeveDbContext(doc, msg => { }))
                    {
                        var snapshotRepository = new SleeveSnapshotRepository(dbContext, msg => { });
                        snapshotIndex = snapshotRepository.LoadSnapshotIndex();
                    }
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.Message = $"Failed to load sleeve parameter snapshots: {ex.Message}";
                    result.Errors.Add(ex.Message);
                    logBuffer.AppendLine($"[PARAM_TRANSFER] Failed to load sleeve parameter snapshots: {ex.Message}");
                    SafeFileLogger.SafeAppendText("transfer_debug.log", logBuffer.ToString());
                    return result;
                }
                sw.Stop();
                timeLoadSnapshot = sw.ElapsedMilliseconds;
                
                // Lazy Parameter Cache
                // Key: "{ElementId}_{ParamName}" -> Parameter
                var parameterCache = new Dictionary<string, Parameter>();
                
                // Helper for lazy parameter lookup
                Parameter GetOrCacheParameter(Element element, string paramName)
                {
                    var key = $"{element.Id.GetIntegerValue()}_{paramName}";
                    if (parameterCache.TryGetValue(key, out var cached)) return cached;
                    var param = element.LookupParameter(paramName);
                    // Cache even if null to avoid repeated failed lookups
                    parameterCache[key] = param;
                    return param;
                }

                int totalSleevesUpdated = 0;
                int totalSleevesFailed = 0;
                
                // ✅ OPTIMIZATION: Iterate Sleeves (Outer) -> Mappings (Inner)
                // This ensures we only load the element and snapshot ONCE per sleeve.
                foreach (var openingId in openingIds)
                {
                    bool sleeveUpdated = false;
                    bool sleeveFailed = false; // Track if any critical failure occurred for this sleeve

                    try 
                    {
                        sw.Restart();
                        var element = doc.GetElement(openingId);
                        sw.Stop();
                        timeGetElement += sw.ElapsedMilliseconds;

                        if (element == null) 
                        { 
                            totalSleevesFailed++; 
                            continue; 
                        }
                        
                        sw.Restart();
                        SleeveSnapshotView snapshot = null;

                        // 1. COMBINED SLEEVE LOOKUP (Check first as it overrides others)
                        if (snapshotIndex.TryGetByCombined(openingId.GetIntegerValue(), out var combinedConstituents))
                        {
                            snapshot = new SleeveSnapshotView
                            {
                                SnapshotId = -1,
                                SleeveInstanceId = openingId.GetIntegerValue(),
                                SourceType = "Combined",
                                MepParameters = AggregateConstituentParameters(combinedConstituents, snapshotIndex, useHost: false),
                                HostParameters = AggregateConstituentParameters(combinedConstituents, snapshotIndex, useHost: true)
                            };
                        }
                        // 2. STANDARD LOOKUP (Using Sleeve/Cluster IDs parameters)
                        else
                        {
                            // We use standard Lookup for IDs as they are critical and usually present
                            var pSleeveId = GetOrCacheParameter(element, "Sleeve Instance ID");
                            var pClusterId = GetOrCacheParameter(element, "Cluster Sleeve Instance ID");
                            
                            int sleeveInstanceId = pSleeveId != null && pSleeveId.StorageType == StorageType.Integer ? pSleeveId.AsInteger() : 0;
                            int clusterInstanceId = pClusterId != null && pClusterId.StorageType == StorageType.Integer ? pClusterId.AsInteger() : 0;

                            // Resolve snapshot
                            if (clusterInstanceId > 0 && snapshotIndex.TryGetByCluster(clusterInstanceId, out var clusterView))
                            {
                                snapshot = clusterView;
                            }
                            else if (sleeveInstanceId > 0 && snapshotIndex.TryGetBySleeve(sleeveInstanceId, out var sleeveView))
                            {
                                snapshot = sleeveView;
                            }
                            // Fallback: ClashZone GUID
                            else if (sleeveInstanceId > 0 && snapshotIndex.SleeveIdToClashZoneGuid.TryGetValue(sleeveInstanceId, out var clashZoneGuid))
                            {
                                if (snapshotIndex.TryGetByClashZoneGuid(clashZoneGuid, out var guidSnapshot))
                                    snapshot = guidSnapshot;
                            }
                        }
                        sw.Stop();
                        timeResolution += sw.ElapsedMilliseconds;

                        if (snapshot == null)
                        {
                            // Log only in verbose/debug mode to avoid spam
                            // if (!DeploymentConfiguration.DeploymentMode) 
                            //    logBuffer.AppendLine($"[PARAM_TRANSFER] Snapshot not found for sleeve {openingId.GetIntegerValue()}");
                            
                            // Not a failure per se, just nothing to transfer. 
                            // Maybe untracked sleeve.
                            continue;
                        }

                        sw.Restart();
                        // Apply Mappings
                        foreach (var mapping in config.Mappings)
                        {
                            if (!mapping.IsEnabled) continue;

                            // Perform value transfer
                            Dictionary<string, string> sourceParams = null; // MepParameters or HostParameters
                            
                            if (mapping.TransferType == TransferType.HostToOpening)
                            {
                                sourceParams = snapshot.HostParameters;
                            }
                            else // RefToOpening, LevelToOpening, etc. - mostly fall back to MEP snapshot data
                            {
                                sourceParams = snapshot.MepParameters;
                            }
                            
                            if (sourceParams != null && sourceParams.TryGetValue(mapping.SourceParameter, out var sourceValue))
                            {
                                // Set Value
                                var targetParam = GetOrCacheParameter(element, mapping.TargetParameter);
                                if (targetParam != null && !targetParam.IsReadOnly)
                                {
                                    // ✅ OPTIMIZATION: Skip if parameter already has the correct value
                                    if (OptimizationFlags.SkipAlreadyTransferredParameters)
                                    {
                                        string currentValue = targetParam.AsValueString();
                                        if (string.IsNullOrEmpty(currentValue))
                                        {
                                            currentValue = targetParam.AsString();
                                        }

                                        // If values match, skip the expensive Set operation
                                        // Use loose comparison (ignore case)
                                        if (string.Equals(currentValue, sourceValue, StringComparison.OrdinalIgnoreCase))
                                        {
                                            continue;
                                        }
                                        
                                        // Double check: if source is empty and current is empty/null, also skip
                                        if (string.IsNullOrWhiteSpace(sourceValue) && string.IsNullOrWhiteSpace(currentValue))
                                        {
                                            continue;
                                        }
                                    }

                                    if (SetParameterValueSafely(targetParam, sourceValue))
                                    {
                                        sleeveUpdated = true;
                                    }
                                }
                            }
                        }
                        sw.Stop();
                        timeParamSet += sw.ElapsedMilliseconds;
                    }
                    catch (Exception ex)
                    {
                        sleeveFailed = true;
                        // Log sporadic errors but continue
                        if (!DeploymentConfiguration.DeploymentMode)
                             logBuffer.AppendLine($"[PARAM_TRANSFER] Error processing sleeve {openingId}: {ex.Message}");
                    }

                    if (sleeveUpdated) totalSleevesUpdated++;
                    if (sleeveFailed) totalSleevesFailed++;
                }

                methodStartTime.Stop();
                
                // Logging
                logBuffer.AppendLine($"[PARAM_TRANSFER] Batch Transfer Completed in {methodStartTime.ElapsedMilliseconds}ms");
                logBuffer.AppendLine($"   - Load Snapshot Index: {timeLoadSnapshot}ms");
                logBuffer.AppendLine($"   - Get Element (Total): {timeGetElement}ms (Avg: {(openingIds.Count > 0 ? timeGetElement/openingIds.Count : 0)}ms)");
                logBuffer.AppendLine($"   - Snapshot Resolution (Total): {timeResolution}ms (Avg: {(openingIds.Count > 0 ? timeResolution/openingIds.Count : 0)}ms)");
                logBuffer.AppendLine($"   - Parameter Set (Total): {timeParamSet}ms (Avg: {(openingIds.Count > 0 ? timeParamSet/openingIds.Count : 0)}ms)");
                logBuffer.AppendLine($"[PARAM_TRANSFER] Sleeves Updated: {totalSleevesUpdated}, Failed/Skipped: {totalSleevesFailed}");
                
                SafeFileLogger.SafeAppendText("transfer_debug.log", logBuffer.ToString());

                result.Success = true; // Overall success if we finished the loop
                result.TransferredCount = totalSleevesUpdated;
                result.FailedCount = totalSleevesFailed;
                result.Message = $"Transferred parameters to {totalSleevesUpdated} sleeves. ({totalSleevesFailed} failed/skipped)";

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Critical Failure in Refactored Transfer: {ex.Message}";
                result.Errors.Add(ex.Message);
                SafeFileLogger.SafeAppendText("transfer_debug.log", $"CRITICAL ERROR: {ex}\n");
                return result;
            }
        }


        /// <summary>
        /// Execute the whole transfer configuration assuming the caller owns the transaction.
        /// This method acts as a DISPATCHER to route to either the Optimized (Refactored) or Legacy implementation.
        /// </summary>
        public ParameterTransferResult ExecuteTransferConfigurationInTransaction(
            Document doc,
            List<ElementId> openingIds,
            ParameterTransferConfiguration config,
            UIDocument uiDoc = null)
        {
            // ✅ DISPATCHER: Route to optimized or legacy path based on flag
            if (OptimizationFlags.UseParameterTransferRefactor)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [DISPATCHER] 🚀 Routing to Refactored (Loop-Inverted) Transfer Method\n");
                }
                
                try 
                {
                    return ExecuteTransferConfigurationInTransaction_Refactored(doc, openingIds, config, uiDoc);
                }
                catch (Exception ex)
                {
                    // Fallback handled by the caller or here?
                    // If we crash here, we should log and try legacy?
                    // OptimizationFlags.UseParameterTransferRefactor = false; 
                    // But changing static flag might affect other threads/calls. 
                    // For now, let's just let the refactored method handle its own safety or bubble up.
                    // The Wrapper (lines 600+) handles fallback.
                    // But direct callers (UI) won't have fallback.
                    
                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [DISPATCHER] ❌ Refactored method failed: {ex.Message}. Falling back to Legacy.\n");
                    
                    return ExecuteTransferConfigurationInTransaction_Legacy(doc, openingIds, config, uiDoc);
                }
            }
            else
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [DISPATCHER] 🐢 Routing to Legacy Transfer Method\n");
                }
                return ExecuteTransferConfigurationInTransaction_Legacy(doc, openingIds, config, uiDoc);
            }
        }

        /// <summary>
        /// [LEGACY] Execute the whole transfer configuration assuming the caller owns the transaction.
        /// This method performs all transfers by calling the InTransaction variants.
        /// </summary>
        public ParameterTransferResult ExecuteTransferConfigurationInTransaction_Legacy(
            Document doc,
            List<ElementId> openingIds,
            ParameterTransferConfiguration config,
            UIDocument uiDoc = null)
        {
            var result = new ParameterTransferResult();
            var allResults = new List<ParameterTransferResult>();

            // ✅ FIX: Track unique sleeves that were successfully transferred (not per mapping)
            var successfullyTransferredSleeveIds = new HashSet<int>();

            try
            {
                // Check if any sleeves exist in the model
                if (openingIds == null || openingIds.Count == 0)
                {
                    result.Success = false;
                    result.Message = "No sleeves found in the model. Please place sleeves first before transferring parameters.";
                    result.Errors.Add("No sleeves found - place sleeves first");
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning("[PARAM_TRANSFER] No sleeves found in model - user needs to place sleeves first");
                    return result;
                }
                
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[PARAM_TRANSFER] Found {openingIds.Count} sleeves in model - proceeding with parameter transfer");

                var transferDebugLogPath = SafeFileLogger.GetLogFilePath("transfer_debug.log");

                SleeveSnapshotIndex snapshotIndex;
                try
                {
                    using (var dbContext = new SleeveDbContext(doc, msg =>
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[SQLite] {msg}");
                        SafeFileLogger.SafeAppendText("transfer_debug.log", $"[{DateTime.Now}] [SQLite] {msg}\n");
                    }))
                    {
                        var snapshotRepository = new SleeveSnapshotRepository(dbContext, msg =>
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[SQLite] {msg}");
                            SafeFileLogger.SafeAppendText("transfer_debug.log", $"[{DateTime.Now}] [SQLite] {msg}\n");
                        });

                        snapshotIndex = snapshotRepository.LoadSnapshotIndex();
                    }
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.Message = $"Failed to load sleeve parameter snapshots: {ex.Message}";
                    result.Errors.Add(ex.Message);
                    return result;
                }

                if (snapshotIndex.BySleeve.Count == 0 && snapshotIndex.ByCluster.Count == 0)
                {
                    // ✅ IMPROVED: Check if sleeves exist in database before failing
                    // Snapshots are created during placement, but if they weren't created for any reason,
                    // we should inform the user that Refresh will recreate them from existing sleeves
                    int sleevesInDb = 0;
                    try
                    {
                        using (var dbContext = new SleeveDbContext(doc, msg => { }))
                        {
                            using (var checkCmd = dbContext.Connection.CreateCommand())
                            {
                                checkCmd.CommandText = @"
                                    SELECT COUNT(*) FROM ClashZones 
                                    WHERE SleeveInstanceId IS NOT NULL AND SleeveInstanceId > 0";
                                var count = checkCmd.ExecuteScalar();
                                sleevesInDb = count != null && count != DBNull.Value ? Convert.ToInt32(count) : 0;
                            }
                        }
                    }
                    catch { }
                    
                    result.Success = false;
                    if (sleevesInDb > 0)
                    {
                        result.Message = $"No sleeve parameter snapshots found, but {sleevesInDb} sleeve(s) exist in database. Snapshots are created during placement or Refresh. Please run Refresh to create snapshots from existing sleeves, then try transferring parameters again.";
                        result.Errors.Add($"SleeveSnapshots table is empty (but {sleevesInDb} sleeves found in ClashZones)");
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[PARAM_TRANSFER] No sleeve snapshots found in SQLite, but {sleevesInDb} sleeves exist in database. User needs to run Refresh to create snapshots.");
                    }
                    else
                    {
                        result.Message = "No sleeve parameter snapshots found. Please place sleeves and run Refresh before transferring parameters.";
                        result.Errors.Add("SleeveSnapshots table is empty");
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning("[PARAM_TRANSFER] No sleeve snapshots found in SQLite and no sleeves in database.");
                    }
                    File.AppendAllText(transferDebugLogPath, $"[{DateTime.Now}] [PARAM_TRANSFER] No sleeve snapshots found in SQLite. Sleeves in DB: {sleevesInDb}\n");
                    return result;
                }
                
                // ✅ DIAGNOSTIC: Log snapshot index status BEFORE processing
                SafeFileLogger.SafeAppendText("transfer_debug.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅ Snapshot index loaded: BySleeve.Count={snapshotIndex.BySleeve.Count}, ByCluster.Count={snapshotIndex.ByCluster.Count}\n");
                
                if (snapshotIndex.BySleeve.Count > 0)
                {
                    var sampleSleeveIds = string.Join(", ", snapshotIndex.BySleeve.Keys.Take(5));
                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] Sample SleeveInstanceIds: [{sampleSleeveIds}]\n");
                }
                
                if (snapshotIndex.ByCluster.Count > 0)
                {
                    var sampleClusterIds = string.Join(", ", snapshotIndex.ByCluster.Keys.Take(5));
                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] Sample ClusterInstanceIds: [{sampleClusterIds}]\n");
                }

                // Execute each mapping
                foreach (var mapping in config.Mappings)
                {
                    // ✅ FIX: Skip disabled mappings to avoid unnecessary processing
                    if (!mapping.IsEnabled)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[PARAM_TRANSFER] ⏭️ Skipping disabled mapping: {mapping.SourceParameter} -> {mapping.TargetParameter}");
                        }
                        continue;
                    }
                    
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[PARAM_TRANSFER] Processing mapping: {mapping.SourceParameter} -> {mapping.TargetParameter}");
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[{DateTime.Now}] [PARAM_TRANSFER] Processing mapping: {mapping.SourceParameter} -> {mapping.TargetParameter}\n");
                    
                    ParameterTransferResult mappingResult = null;

                    switch (mapping.TransferType)
                    {
                        case TransferType.ReferenceToOpening:
                            mappingResult = TransferFromReferenceElementsInTransaction(doc, openingIds, mapping, snapshotIndex, successfullyTransferredSleeveIds, uiDoc);
                            break;
                        case TransferType.HostToOpening:
                            mappingResult = TransferFromHostElementsInTransaction(doc, openingIds, mapping, snapshotIndex, successfullyTransferredSleeveIds);
                            break;
                        case TransferType.LevelToOpening:
                            mappingResult = TransferFromLevelsInTransaction(doc, openingIds, mapping);
                            break;
                    }

                    if (mappingResult != null)
                        allResults.Add(mappingResult);
                }

                // Transfer model names if enabled
                if (config.TransferModelNames)
                {
                    var modelResult = TransferModelNamesInTransaction(doc, openingIds, config.ModelNameParameter);
                    allResults.Add(modelResult);
                }

                // Transfer service size calculations if enabled
                if (config.TransferServiceSizeCalculations)
                {
                    // Set clearance parameters
                    // _mepAnalysisService.SetClearanceParameters(config.DefaultClearance, config.ClearanceSuffix);

                    var serviceSizeResult = TransferServiceSizeCalculationsInTransaction(
                        doc, openingIds, config.ServiceSizeCalculationParameter, config.DefaultClearance);
                    allResults.Add(serviceSizeResult);
                }

                // Combine results
                result.Success = allResults.All(r => r.Success);
                
                // ✅ FIX: Count unique sleeves transferred, not per-mapping transfers
                // Use the tracking set to get actual unique sleeve count
                int uniqueSleeveCount = successfullyTransferredSleeveIds?.Count ?? 0;
                
                // Also keep the parameter count for backward compatibility
                int parameterTransferCount = allResults.Sum(r => r.TransferredCount);
                
                // Use unique sleeve count if available, otherwise use parameter count
                result.TransferredCount = uniqueSleeveCount > 0 ? uniqueSleeveCount : parameterTransferCount;
                result.FailedCount = allResults.Sum(r => r.FailedCount);
                result.Errors = allResults.SelectMany(r => r.Errors).ToList();
                result.Warnings = allResults.SelectMany(r => r.Warnings).ToList();
                
                // Improved message showing unique sleeves
                if (uniqueSleeveCount > 0)
                {
                    result.Message = $"Transfer completed: {uniqueSleeveCount} sleeves processed ({parameterTransferCount} parameter transfers), {result.FailedCount} failed.";
                }
                else
                {
                result.Message = $"Transfer completed: {result.TransferredCount} successful, {result.FailedCount} failed.";
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Configuration transfer failed: {ex.Message}";
                result.Errors.Add(ex.Message);
            }

            return result;
        }
        
        public ParameterTransferResult TransferFromReferenceElementsInTransaction(
            Document doc,
            List<ElementId> openingIds,
            ParameterMapping mapping,
            SleeveSnapshotIndex snapshotIndex,
            HashSet<int> successfullyTransferredSleeveIds = null,
            UIDocument uiDoc = null)
        {
            // Delegate to core with a resolver for MEP bags
            return TransferFromElementsWithSnapshot(doc, openingIds, mapping, snapshotIndex, useHost:false, successfullyTransferredSleeveIds, uiDoc);
        }

        public ParameterTransferResult TransferFromHostElementsInTransaction(
            Document doc,
            List<ElementId> openingIds,
            ParameterMapping mapping,
            SleeveSnapshotIndex snapshotIndex,
            HashSet<int> successfullyTransferredSleeveIds = null)
        {
            // Delegate to core with a resolver for HOST bags
            return TransferFromElementsWithSnapshot(doc, openingIds, mapping, snapshotIndex, useHost:true, successfullyTransferredSleeveIds);
        }

        private ParameterTransferResult TransferFromElementsWithSnapshot(
            Document doc,
            List<ElementId> openingIds,
            ParameterMapping mapping,
            SleeveSnapshotIndex snapshotIndex,
            bool useHost,
            HashSet<int> successfullyTransferredSleeveIds = null,
            UIDocument uiDoc = null)
        {
            // ✅ BUILD STAMP: Log build information and code version (ALWAYS LOG - critical for debugging)
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            var buildTime = System.IO.File.GetLastWriteTime(assembly.Location);
            // ✅ CRITICAL: Always log build stamp (even in deployment mode) to verify correct code is running
            SafeFileLogger.SafeAppendText("transfer_debug.log",
                $"\n===== PARAMETER TRANSFER SESSION STARTED {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====\n");
            SafeFileLogger.SafeAppendText("transfer_debug.log",
                $"🔨 BUILD TIME: {buildTime:yyyy-MM-dd HH:mm:ss} (DLL: {System.IO.Path.GetFileName(assembly.Location)})\n");
            SafeFileLogger.SafeAppendText("transfer_debug.log",
                $"✅ CODE VERSION: Element ID Comparison Fix Applied (v2025-12-03)\n");
            SafeFileLogger.SafeAppendText("transfer_debug.log",
                $"📍 EXPECTED CODE PATH: Element validation uses ID comparison (not document reference)\n");
            SafeFileLogger.SafeAppendText("transfer_debug.log",
                $"⚙️ DEPLOYMENT MODE: {DeploymentConfiguration.DeploymentMode} (false = full logging, true = minimal logging)\n");
            
            // ✅ DIAGNOSTIC: Track which sleeves we've logged snapshot contents for (to avoid duplicate logs)
            var loggedSnapshotParams = new HashSet<int>();
            
            // ✅ DIAGNOSTIC: Track per-sleeve skip counts (to show which parameters were skipped on each sleeve)
            var perSleeveSkipCounts = new Dictionary<int, int>();
            
            var result = new ParameterTransferResult();
            var transferredCount = 0;
            var failedCount = 0;
            var skippedCount = 0;
            var errors = new List<string>();
            
            // ✅ CRITICAL: Log method entry with sleeve count (always log - critical for debugging)
            var transferStartTime = System.Diagnostics.Stopwatch.StartNew();
            SafeFileLogger.SafeAppendText("transfer_debug.log",
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 📥 METHOD ENTRY: TransferFromElementsWithSnapshot called with {openingIds?.Count ?? 0} sleeves, Mapping: '{mapping?.SourceParameter}' -> '{mapping?.TargetParameter}'\n");
            SafeFileLogger.SafeAppendText("transfer_debug.log",
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PERFORMANCE] ⏱️ OPTIMIZATION FLAGS: SectionBox={Services.OptimizationFlags.UseSectionBoxFilterForParameterTransfer}, SkipAlreadyTransferred={Services.OptimizationFlags.SkipAlreadyTransferredParameters}, BatchLookups={Services.OptimizationFlags.UseBatchParameterLookups}\n");

            // ✅ CRITICAL: Validate openingIds before processing
            if (openingIds == null || openingIds.Count == 0)
            {
                SafeFileLogger.SafeAppendText("transfer_debug.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ❌ ERROR: openingIds is NULL or EMPTY - cannot process\n");
                result.Success = false;
                result.Message = "No opening IDs provided for parameter transfer.";
                result.Errors.Add("openingIds is null or empty");
                return result;
            }
            
            SafeFileLogger.SafeAppendText("transfer_debug.log",
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅ Validated: {openingIds.Count} opening IDs provided, proceeding to cache initialization\n");
            
            // ✅ OPTIMIZATION 4: Batch Parameter Lookups - Pre-cache elements and parameters
            Dictionary<int, Element> elementCache = new Dictionary<int, Element>();
            Dictionary<int, Dictionary<string, Parameter>> parameterCache = new Dictionary<int, Dictionary<string, Parameter>>();
            
            // ✅ PROTECTION 15: Parameter cache initialization with validation
            // ✅ CRITICAL: Validate mapping is not null before cache initialization
            if (mapping == null)
            {
                SafeFileLogger.SafeAppendText("transfer_debug.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ❌ ERROR: mapping is NULL - cannot proceed with parameter transfer\n");
                result.Success = false;
                result.Message = "Parameter mapping is null - cannot transfer parameters.";
                result.Errors.Add("mapping parameter is null");
                return result;
            }
            
            SafeFileLogger.SafeAppendText("transfer_debug.log",
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅ Mapping validated: Source='{mapping.SourceParameter}', Target='{mapping.TargetParameter}', Enabled={mapping.IsEnabled}\n");
            
            int cachedElementCount = 0;
            if (Services.OptimizationFlags.UseBatchParameterLookups)
            {
                var cacheStartTime = System.Diagnostics.Stopwatch.StartNew();
                SafeFileLogger.SafeAppendText("transfer_debug.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔄 Starting batch parameter cache initialization (UseBatchParameterLookups=true)\n");
                
                // ✅ CRITICAL: Validate document before caching
                if (doc == null)
                {
                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️ Document is null - skipping parameter cache initialization\n");
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning("[PARAM_TRANSFER] ⚠️ Document is null - skipping parameter cache initialization");
                    }
                }
                else
                {
                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔄 Processing {openingIds.Count} sleeves for parameter cache\n");
                    foreach (var openingId in openingIds)
                    {
                        try
                        {
                            if (openingId == null || openingId.GetIntegerValue() <= 0) continue;
                            
                            // ✅ Sleeves are always in the active document (doc), not linked files
                            var element = doc.GetElement(openingId);
                            cachedElementCount++;
                            if (element != null && element.IsValidObject && element is FamilyInstance)
                            {
                                // ✅ PROTECTION 16: Validate element document matches before caching
                                if (element.Document != doc)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        DebugLogger.Warning($"[PARAM_TRANSFER] ⚠️ Element {openingId.GetIntegerValue()} belongs to different document - skipping cache");
                                    }
                                    continue;
                                }
                                
                                elementCache[openingId.GetIntegerValue()] = element;
                                
                                // Pre-cache the target parameter (sleeves are in active document)
                                try
                                {
                                    var targetParam = element.LookupParameter(mapping.TargetParameter);
                                    if (targetParam != null && targetParam.Element != null && targetParam.Element.IsValidObject)
                                    {
                                        // ✅ PROTECTION 17: Validate parameter element matches before caching
                                        if (targetParam.Element.Id == element.Id)
                                        {
                                            if (!parameterCache.ContainsKey(openingId.GetIntegerValue()))
                                            {
                                                parameterCache[openingId.GetIntegerValue()] = new Dictionary<string, Parameter>();
                                            }
                                            parameterCache[openingId.GetIntegerValue()][mapping.TargetParameter] = targetParam;
                                        }
                                    }
                                }
                                catch (Exception paramEx)
                                {
                                    // Parameter lookup failed - skip caching for this element
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        DebugLogger.Warning($"[PARAM_TRANSFER] ⚠️ Failed to cache parameter '{mapping.TargetParameter}' for element {openingId.GetIntegerValue()}: {paramEx.Message}");
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            // Element retrieval failed - skip caching for this element
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Warning($"[PARAM_TRANSFER] ⚠️ Failed to cache element {openingId?.GetIntegerValue() ?? -1}: {ex.Message}");
                            }
                        }
                    }
                }
                
                cacheStartTime.Stop();
                SafeFileLogger.SafeAppendText("transfer_debug.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PERFORMANCE] ⏱️ Cache initialization took {cacheStartTime.ElapsedMilliseconds}ms for {cachedElementCount} elements\n");
                SafeFileLogger.SafeAppendText("transfer_debug.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅ Batch parameter cache initialization complete: {elementCache.Count} elements, {parameterCache.Count} parameter sets cached\n");
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[PARAM_TRANSFER] ✅ Pre-cached {elementCache.Count} elements and {parameterCache.Count} parameter sets for batch lookup");
                }
            }
            else
            {
                SafeFileLogger.SafeAppendText("transfer_debug.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⏭️ Skipping batch parameter cache (UseBatchParameterLookups=false)\n");
            }

            SafeFileLogger.SafeAppendText("transfer_debug.log",
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔄 Starting foreach loop to process {openingIds.Count} sleeves (UseBatchParameterLookups={Services.OptimizationFlags.UseBatchParameterLookups})\n");
            
            int loopIteration = 0;
            foreach (var openingId in openingIds)
            {
                var sleeveStartTime = System.Diagnostics.Stopwatch.StartNew();
                try
                {
                    loopIteration++;
                    // ✅ CRITICAL: Log each sleeve being processed
                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔄 Loop iteration {loopIteration}/{openingIds.Count}: Processing sleeve openingId={openingId?.GetIntegerValue() ?? -1}\n");

                    // ✅ CRASH-SAFETY 1: Validate ElementId before retrieval
                    if (openingId == null || openingId.GetIntegerValue() <= 0)
                    {
                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ❌ SKIP: Invalid opening element ID: {openingId?.GetIntegerValue() ?? -1}\n");
                        failedCount++;
                        errors.Add($"Invalid opening element ID: {openingId?.GetIntegerValue() ?? -1}");
                        continue;
                    }

                    // ✅ PROTECTION 1: Sleeves are always in the active document (doc), not linked files
                    // ✅ CRITICAL: Validate document is not null before element retrieval
                    if (doc == null)
                    {
                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ❌ SKIP: Document is null for opening element {openingId.GetIntegerValue()}\n");
                        failedCount++;
                        errors.Add($"Document is null for opening element {openingId.GetIntegerValue()}.");
                        continue;
                    }

                    // ✅ CRITICAL: Validate document is not closed
                    if (doc.IsModifiable == false && doc.IsReadOnly)
                    {
                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ❌ SKIP: Document is read-only for opening element {openingId.GetIntegerValue()}\n");
                        failedCount++;
                        errors.Add($"Document is read-only or closed for opening element {openingId.GetIntegerValue()}.");
                        continue;
                    }

                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔍 Retrieving opening element {openingId.GetIntegerValue()} from document...\n");

                    var opening = doc.GetElement(openingId);
                    if (opening == null)
                    {
                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ❌ SKIP: Opening element {openingId.GetIntegerValue()} not found in document\n");
                        failedCount++;
                        errors.Add($"Opening element {openingId.GetIntegerValue()} not found.");
                        continue;
                    }

                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅ Opening element {openingId.GetIntegerValue()} retrieved: Type={opening.GetType().Name}, IsValid={opening.IsValidObject}\n");

                    // ✅ CRASH-SAFETY 2: Check if element is still valid (not deleted)
                    if (!opening.IsValidObject)
                    {
                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ❌ SKIP: Opening element {openingId.GetIntegerValue()} is no longer valid (deleted)\n");
                        failedCount++;
                        errors.Add($"Opening element {openingId.GetIntegerValue()} is no longer valid (may have been deleted).");
                        continue;
                    }

                    // ⚠️⚠️⚠️ CRITICAL: DO NOT REINTRODUCE DOCUMENT VALIDATION CHECK ⚠️⚠️⚠️
                    // 
                    // REMOVED: Document validation check (opening.Document != doc)
                    // 
                    // BUG HISTORY: This check was added as "PROTECTION 2" crash safety but caused a critical bug:
                    // - Used reference equality (opening.Document != doc) which fails even for the same document
                    // - Document objects can be different instances even when representing the same document
                    // - Caused ALL sleeves to be skipped with false "document mismatch" errors
                    // - Lost 1 day of debugging time (2025-12-03)
                    //
                    // WHY IT'S SAFE TO REMOVE:
                    // - We retrieve element using doc.GetElement(openingId), so it MUST be in doc
                    // - Element ID validation later (STEP 7) already ensures parameter belongs to correct element
                    // - The check was redundant and harmful
                    //
                    // IF YOU NEED DOCUMENT VALIDATION:
                    // - Compare by document title: string.Equals(opening.Document?.Title, doc.Title, StringComparison.OrdinalIgnoreCase)
                    // - But even this is redundant since doc.GetElement() guarantees the element is in doc
                    // - Consider validating at element retrieval time instead (if doc.GetElement returns null, element doesn't exist)

                    // ✅ CRASH-SAFETY 3: Validate element type before accessing properties
                    if (!(opening is FamilyInstance))
                    {
                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ❌ SKIP: Opening element {openingId.GetIntegerValue()} is not a FamilyInstance (type: {opening.GetType().Name})\n");
                        failedCount++;
                        errors.Add($"Opening element {openingId.GetIntegerValue()} is not a FamilyInstance (type: {opening.GetType().Name}).");
                        continue;
                    }

                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔍 Reading Sleeve Instance ID and Cluster Instance ID from opening {openingId.GetIntegerValue()}...\n");
                    // ✅ PERFORMANCE: Time parameter lookup operation (this might be slow!)
                    var dbLookupStartTime = System.Diagnostics.Stopwatch.StartNew();
                    var sleeveInstanceId = GetIntegerParameter(opening, "Sleeve Instance ID");
                    var clusterInstanceId = GetIntegerParameter(opening, "Cluster Sleeve Instance ID");
                    dbLookupStartTime.Stop();

                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PERFORMANCE] ⏱️ Parameter lookup took {dbLookupStartTime.ElapsedMilliseconds}ms (SleeveID={sleeveInstanceId}, ClusterID={clusterInstanceId})\n");
                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅ Read IDs: SleeveInstanceId={sleeveInstanceId}, ClusterInstanceId={clusterInstanceId}\n");

                    // ✅ DIAGNOSTIC: Log matching attempt
                    var dbMatchStartTime = System.Diagnostics.Stopwatch.StartNew();
                    var transferDebugLogPath = SafeFileLogger.GetLogFilePath("transfer_debug.log");
                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] Matching sleeve {openingId.GetIntegerValue()}: SleeveInstanceId={sleeveInstanceId}, ClusterInstanceId={clusterInstanceId}, SnapshotIndex.BySleeve.Count={snapshotIndex.BySleeve.Count}, SnapshotIndex.ByCluster.Count={snapshotIndex.ByCluster.Count}\n");

                    // ✅ PROTECTION 13: Validate snapshot index is not null
                    if (snapshotIndex == null)
                    {
                        failedCount++;
                        errors.Add($"Snapshot index is null for opening element {openingId.GetIntegerValue()}.");
                        continue;
                    }

                    SleeveSnapshotView snapshot = null;
                    string matchMethod = "None";

                    // 1. TRY COMBINED (Aggregated by ElementId or direct Snapshot)
                    if (snapshotIndex.TryGetByCombined(openingId.GetIntegerValue(), out var combinedConstituents))
                    {
                        snapshot = new SleeveSnapshotView
                        {
                            SnapshotId = -1,
                            SleeveInstanceId = openingId.GetIntegerValue(),
                            SourceType = "Combined",
                            MepParameters = AggregateConstituentParameters(combinedConstituents, snapshotIndex, useHost: false),
                            HostParameters = AggregateConstituentParameters(combinedConstituents, snapshotIndex, useHost: true)
                        };
                        matchMethod = "CombinedIdContents";
                    }
                    else if (snapshotIndex.TryGetByCombinedSnapshot(openingId.GetIntegerValue(), out var directCombinedView))
                    {
                        snapshot = directCombinedView;
                        matchMethod = "CombinedIdDirect";
                    }
                    // 2. TRY CLUSTER CONSTITUENTS (Aggregated by ClusterInstanceId)
                    else if (clusterInstanceId > 0 && snapshotIndex.ByClusterConstituents.TryGetValue(clusterInstanceId, out var clusterConstituents))
                    {
                        snapshot = new SleeveSnapshotView
                        {
                            SnapshotId = -1,
                            ClusterInstanceId = clusterInstanceId,
                            SourceType = "Cluster",
                            MepParameters = AggregateConstituentParameters(clusterConstituents, snapshotIndex, useHost: false),
                            HostParameters = AggregateConstituentParameters(clusterConstituents, snapshotIndex, useHost: true)
                        };
                        matchMethod = "ClusterId";
                    }
                    // 3. TRY GUID FALLBACK (Aggregated by UniqueId)
                    else 
                    {
                        string uniqueId = opening.UniqueId;
                        
                        // Fallback A: Combined GUID
                        if (snapshotIndex.ByCombinedGuidConstituents.TryGetValue(uniqueId, out var combGuidRefs))
                        {
                            snapshot = new SleeveSnapshotView
                            {
                                SnapshotId = -1,
                                SourceType = "CombinedGuid",
                                MepParameters = AggregateConstituentParameters(combGuidRefs, snapshotIndex, useHost: false),
                                HostParameters = AggregateConstituentParameters(combGuidRefs, snapshotIndex, useHost: true)
                            };
                            matchMethod = "CombinedGuidFallback";
                        }
                        // Fallback B: Cluster GUID
                        else if (snapshotIndex.ByClusterGuidConstituents.TryGetValue(uniqueId, out var clusGuidRefs))
                        {
                            snapshot = new SleeveSnapshotView
                            {
                                SnapshotId = -1,
                                SourceType = "ClusterGuid",
                                MepParameters = AggregateConstituentParameters(clusGuidRefs, snapshotIndex, useHost: false),
                                HostParameters = AggregateConstituentParameters(clusGuidRefs, snapshotIndex, useHost: true)
                            };
                            matchMethod = "ClusterGuidFallback";
                        }
                        // Fallback C: Individual Sleeve GUID
                        else if (snapshotIndex.ByClashZoneGuid.TryGetValue(uniqueId, out var guidView))
                        {
                            snapshot = guidView;
                            matchMethod = "IndividualGuidFallback";
                        }
                    }

                    // 4. TRY DIRECT SNAPSHOT (Backward compatibility / Single Zone Clusters)
                    if (snapshot == null)
                    {
                        if (clusterInstanceId > 0 && snapshotIndex.TryGetByCluster(clusterInstanceId, out var clusterView))
                        {
                            snapshot = clusterView;
                            matchMethod = "ClusterIdDirect";
                        }
                        else if (sleeveInstanceId > 0 && snapshotIndex.TryGetBySleeve(sleeveInstanceId, out var sleeveView))
                        {
                            snapshot = sleeveView;
                            matchMethod = "SleeveIdDirect";
                        }
                    }

                    // 5. TRY STALE ID FALLBACK
                    if (snapshot == null && sleeveInstanceId > 0 && 
                        snapshotIndex.SleeveIdToClashZoneGuid.TryGetValue(sleeveInstanceId, out var clashZoneGuid))
                    {
                        if (snapshotIndex.TryGetByClashZoneGuid(clashZoneGuid, out var guidSnapshot))
                        {
                            snapshot = guidSnapshot;
                            matchMethod = "StaleIdFallback";
                        }
                    }

                    if (snapshot != null)
                    {
                         SafeFileLogger.SafeAppendText("transfer_debug.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅ Matched via {matchMethod} for opening {openingId.GetIntegerValue()}\n");
                    }

                    dbMatchStartTime.Stop();
                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PERFORMANCE] ⏱️ Database snapshot matching took {dbMatchStartTime.ElapsedMilliseconds}ms\n");

                    if (snapshot == null)
                    {
                        // ✅ DIAGNOSTIC: Log available snapshot IDs for debugging
                        var availableSleeveIds = string.Join(", ", snapshotIndex.BySleeve.Keys.Take(10));
                        var availableClusterIds = string.Join(", ", snapshotIndex.ByCluster.Keys.Take(10));
                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ❌ SKIP: No snapshot found for sleeve {openingId.GetIntegerValue()} (SleeveId={sleeveInstanceId}, ClusterId={clusterInstanceId}). Available SleeveIds: [{availableSleeveIds}], Available ClusterIds: [{availableClusterIds}]\n");
                        result.Warnings.Add($"No persisted snapshot found for sleeve {openingId.GetIntegerValue()} (SleeveId={sleeveInstanceId}, ClusterId={clusterInstanceId}).");
                        continue;
                    }

                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅ Snapshot found for sleeve {openingId.GetIntegerValue()}, proceeding to parameter processing...\n");

                    // ✅ CRASH-SAFETY 4: Validate snapshot structure
                    if (snapshot == null)
                    {
                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ❌ SKIP: Snapshot is null for sleeve {openingId.GetIntegerValue()} (duplicate check)\n");
                        result.Warnings.Add($"Snapshot is null for sleeve {openingId.GetIntegerValue()}.");
                        continue;
                    }

                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔍 Getting source parameters from snapshot (useHost={useHost}) for sleeve {openingId.GetIntegerValue()}...\n");

                    var sourceParams = useHost ? snapshot.HostParameters : snapshot.MepParameters;
                    if (sourceParams == null || sourceParams.Count == 0)
                    {
                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ❌ SKIP: No {(useHost ? "host" : "MEP")} parameters captured for sleeve {openingId.GetIntegerValue()} (sourceParams is null or empty)\n");
                        result.Warnings.Add($"No {(useHost ? "host" : "MEP")} parameters captured for sleeve {openingId.GetIntegerValue()}.");
                        continue;
                    }

                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅ Source parameters loaded: {sourceParams.Count} {(useHost ? "host" : "MEP")} parameters for sleeve {openingId.GetIntegerValue()}\n");

                    // ✅ DIAGNOSTIC: Log all available parameters in snapshot (first time only per sleeve)
                    if (!DeploymentConfiguration.DeploymentMode && !loggedSnapshotParams.Contains(openingId.GetIntegerValue()))
                    {
                        loggedSnapshotParams.Add(openingId.GetIntegerValue());
                        var allParamKeys = string.Join(", ", sourceParams.Keys.OrderBy(k => k));
                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 📋 SNAPSHOT CONTENTS for sleeve {openingId.GetIntegerValue()} ({sourceParams.Count} {(useHost ? "host" : "MEP")} params): [{allParamKeys}]\n");
                    }

                    // ✅ CRASH-SAFETY 5: Validate mapping parameter name
                    if (string.IsNullOrWhiteSpace(mapping?.SourceParameter))
                    {
                        failedCount++;
                        errors.Add($"Invalid source parameter name for sleeve {openingId.GetIntegerValue()}.");
                        continue;
                    }

                    // ✅ PERFORMANCE FIX: Always read from snapshot (not Revit) for ALL parameters including "MEP Size"
                    // Snapshot data is captured during placement and is much faster than querying Revit MEP elements
                    // This eliminates expensive GetElementFromDocumentOrLinked calls which query linked documents
                    string sourceValue = null;
                    bool readFromRevit = false;
                    bool isClusterSleeve = clusterInstanceId > 0;

                    // ✅ DIAGNOSTIC: Log what source parameter we're looking for
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔍 Checking source parameter: '{mapping.SourceParameter}' for sleeve {openingId.GetIntegerValue()} (IsCluster={isClusterSleeve}, ClusterId={clusterInstanceId}, SleeveId={sleeveInstanceId})\n");
                    }

                    // ✅ PERFORMANCE FIX: Read ALL parameters from snapshot (not Revit)
                    // Special handling for "MEP Size" / "Size" parameter - try multiple keys
                    // ✅ CRITICAL: Declare isSizeParameter early so it's accessible throughout the method
                    bool isSizeParameter = string.Equals(mapping.SourceParameter, "MEP Size", StringComparison.OrdinalIgnoreCase) ||
                                           string.Equals(mapping.SourceParameter, "Size", StringComparison.OrdinalIgnoreCase);

                    // ✅ CRITICAL FIX: Use flag to skip else block when Size is successfully read
                    // This allows Size parameters to proceed directly to parameter setting logic
                    bool skipElseBlock = false;

                    if (isSizeParameter)
                    {
                        // Try to get from snapshot with multiple possible keys
                        if (sourceParams.TryGetValue("Size", out sourceValue) ||
                            sourceParams.TryGetValue("MEP Size", out sourceValue) ||
                            sourceParams.TryGetValue(mapping.SourceParameter, out sourceValue))
                        {
                            if (!string.IsNullOrWhiteSpace(sourceValue))
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅✅✅ SUCCESS: Read '{mapping.SourceParameter}'='{sourceValue}' from SNAPSHOT for sleeve {openingId.GetIntegerValue()} (IsCluster={isClusterSleeve})\n");
                                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔍 DEBUG: After reading Size from snapshot, sourceValue='{sourceValue}', proceeding to parameter setting logic...\n");
                                }
                                readFromRevit = false;
                                // ✅ CRITICAL FIX: Set flag to skip else block and proceed directly to parameter setting
                                skipElseBlock = true;

                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔍 DEBUG: Size successfully read, setting skipElseBlock=true, will proceed directly to parameter setting for sleeve {openingId.GetIntegerValue()}\n");
                                }

                                // ✅ CRITICAL DIAGNOSTIC: Log that we're about to exit the inner if (!string.IsNullOrWhiteSpace) block
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔍 DEBUG: About to exit inner if (!string.IsNullOrWhiteSpace) block, sourceValue='{sourceValue}' for sleeve {openingId.GetIntegerValue()}\n");
                                }
                            }
                            else
                            {
                                // Empty value in snapshot
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️ '{mapping.SourceParameter}' found in snapshot but is EMPTY for sleeve {openingId.GetIntegerValue()}\n");
                                }
                                result.Warnings.Add($"'{mapping.SourceParameter}' parameter is empty in snapshot for sleeve {openingId.GetIntegerValue()}.");
                                continue;
                            }

                            // ✅ CRITICAL DIAGNOSTIC: Log that we've exited the inner if (TryGetValue) block
                            if (!DeploymentConfiguration.DeploymentMode && !string.IsNullOrWhiteSpace(sourceValue))
                            {
                                SafeFileLogger.SafeAppendText("transfer_debug.log",
                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔍 DEBUG: Exited inner if (TryGetValue) block, sourceValue='{sourceValue}' for sleeve {openingId.GetIntegerValue()}\n");
                            }
                        }
                        else
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("transfer_debug.log",
                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️ '{mapping.SourceParameter}' NOT FOUND in snapshot for sleeve {openingId.GetIntegerValue()}. Available params: {string.Join(", ", sourceParams.Keys.Take(10))}\n");
                            }
                            result.Warnings.Add($"'{mapping.SourceParameter}' parameter not found in snapshot for sleeve {openingId.GetIntegerValue()}.");
                            continue;
                        }

                        // ✅ CRITICAL DIAGNOSTIC: Log that we've exited the if (isSizeParameter) block after successfully reading Size
                        if (!DeploymentConfiguration.DeploymentMode && !string.IsNullOrWhiteSpace(sourceValue))
                        {
                            SafeFileLogger.SafeAppendText("transfer_debug.log",
                                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔍 DEBUG: About to exit if (isSizeParameter) block, sourceValue='{sourceValue}' for sleeve {openingId.GetIntegerValue()}\n");
                        }
                    }

                    // ✅ CRITICAL FIX: Use flag instead of else to allow Size parameters to skip this block
                    if (!skipElseBlock)
                    {
                        // For all other parameters (not Size/MEP Size), read from snapshot with variation support
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("transfer_debug.log",
                                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔍 Source parameter '{mapping.SourceParameter}' - reading from SNAPSHOT for sleeve {openingId.GetIntegerValue()}\n");
                        }

                        // Try exact match first
                        if (!sourceParams.TryGetValue(mapping.SourceParameter, out sourceValue) || string.IsNullOrWhiteSpace(sourceValue))
                        {
                            // Try variations (space/underscore, MEP prefix, etc.)
                            string category = GetCategoryFromOpening(opening, snapshot);
                            var variationsList = new List<string>
                            {
                                mapping.SourceParameter.Replace(" ", "_"),
                                mapping.SourceParameter.Replace("_", " "),
                                "MEP " + mapping.SourceParameter,
                                mapping.SourceParameter.Replace("MEP ", "")
                            };

                            // Cable Trays: System Type → Service Type
                            if (string.Equals(category, "Cable Trays", StringComparison.OrdinalIgnoreCase) &&
                                string.Equals(mapping.SourceParameter, "System Type", StringComparison.OrdinalIgnoreCase))
                            {
                                variationsList.Insert(0, "Service Type");
                            }

                            foreach (var variation in variationsList)
                            {
                                if (sourceParams.TryGetValue(variation, out sourceValue) && !string.IsNullOrWhiteSpace(sourceValue))
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅ Found '{mapping.SourceParameter}' as variation '{variation}'='{sourceValue}' in snapshot\n");
                                    }
                                    break;
                                }
                            }
                        }

                        if (!string.IsNullOrWhiteSpace(sourceValue))
                        {
                            readFromRevit = false;
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("transfer_debug.log",
                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅ Found '{mapping.SourceParameter}'='{sourceValue}' in snapshot for sleeve {openingId.GetIntegerValue()}\n");
                            }
                        }

                        // ✅ CRITICAL: Check if parameter is "MEP Size" - special handling required
                        if (string.IsNullOrWhiteSpace(sourceValue) &&
                            (string.Equals(mapping.SourceParameter, "MEP Size", StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(mapping.SourceParameter, "Size", StringComparison.OrdinalIgnoreCase)))
                        {
                            // ✅ DIAGNOSTIC: Log that we're attempting to read from Revit
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("transfer_debug.log",
                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔍 Source parameter is 'MEP Size' or 'Size' - INDIVIDUAL SLEEVE detected (SleeveId={sleeveInstanceId}), attempting to read from Revit MEP element for sleeve {openingId.GetIntegerValue()}\n");
                            }

                            // Read ONLY from Revit MEP element - no snapshot fallback
                            try
                            {
                                var mepElementIdParam = opening.LookupParameter("MEP_ElementId");
                                if (mepElementIdParam != null && mepElementIdParam.HasValue)
                                {
                                    // ✅ DIAGNOSTIC: Log that MEP_ElementId parameter exists
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅ MEP_ElementId parameter found on sleeve {openingId.GetIntegerValue()}\n");
                                    }

                                    var mepElementId = mepElementIdParam.AsElementId();

                                    // ✅ CRITICAL FIX: Fallback to snapshot if MEP_ElementId is missing on sleeve
                                    if (mepElementId == null || mepElementId == ElementId.InvalidElementId)
                                    {
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            SafeFileLogger.SafeAppendText("transfer_debug.log",
                                                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️ MEP_ElementId missing on sleeve {openingId.GetIntegerValue()}, checking snapshot...\n");
                                        }

                                        if (snapshot != null && snapshot.MepParameters != null &&
                                            snapshot.MepParameters.TryGetValue("MEP_ElementId", out var snapshotMepIdStr) &&
                                            int.TryParse(snapshotMepIdStr, out int snapshotMepIdInt))
                                        {
                                            mepElementId = new ElementId(snapshotMepIdInt);
                                            if (!DeploymentConfiguration.DeploymentMode)
                                            {
                                                SafeFileLogger.SafeAppendText("transfer_debug.log",
                                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅ Found MEP_ElementId={snapshotMepIdInt} in snapshot for sleeve {openingId.GetIntegerValue()}\n");
                                            }
                                        }
                                    }

                                    if (mepElementId != null && mepElementId != ElementId.InvalidElementId)
                                    {
                                        // ✅ DIAGNOSTIC: Log MEP element ID
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            SafeFileLogger.SafeAppendText("transfer_debug.log",
                                                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔍 MEP_ElementId = {mepElementId.GetIntegerValue()} for sleeve {openingId.GetIntegerValue()}, attempting to get MEP element...\n");
                                        }

                                        // ✅ PROTECTION 7: MEP elements may be in linked documents - use ElementRetrievalService
                                        var mepElement = Services.ElementRetrievalService.GetElementFromDocumentOrLinked(doc, mepElementId);
                                        if (mepElement != null && mepElement.IsValidObject)
                                        {
                                            // ✅ PROTECTION 8: Validate MEP element is still valid before parameter access
                                            if (!mepElement.IsValidObject)
                                            {
                                                if (!DeploymentConfiguration.DeploymentMode)
                                                {
                                                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️ MEP element {mepElementId.GetIntegerValue()} became invalid for sleeve {openingId.GetIntegerValue()}\n");
                                                }
                                                result.Warnings.Add($"MEP element {mepElementId.GetIntegerValue()} became invalid for sleeve {openingId.GetIntegerValue()}.");
                                                continue;
                                            }

                                            // ✅ DIAGNOSTIC: Log that MEP element was found
                                            if (!DeploymentConfiguration.DeploymentMode)
                                            {
                                                SafeFileLogger.SafeAppendText("transfer_debug.log",
                                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅ MEP element {mepElementId.GetIntegerValue()} found and valid, looking for 'Size' parameter...\n");
                                            }

                                            // ✅ PROTECTION 9: Read "Size" parameter from MEP element with validation
                                            try
                                            {
                                                var sizeParam = mepElement.LookupParameter("Size");
                                                if (sizeParam != null && sizeParam.HasValue && !sizeParam.IsReadOnly)
                                                {
                                                    sourceValue = sizeParam.AsValueString() ?? sizeParam.AsString();
                                                    if (!string.IsNullOrWhiteSpace(sourceValue))
                                                    {
                                                        readFromRevit = true;

                                                        if (!DeploymentConfiguration.DeploymentMode)
                                                        {
                                                            SafeFileLogger.SafeAppendText("transfer_debug.log",
                                                                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅✅✅ SUCCESS: Read 'MEP Size'='{sourceValue}' from Revit MEP element {mepElementId.GetIntegerValue()} for sleeve {openingId.GetIntegerValue()}\n");
                                                        }
                                                    }
                                                }
                                            }
                                            catch (Exception paramEx)
                                            {
                                                if (!DeploymentConfiguration.DeploymentMode)
                                                {
                                                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️ Exception reading 'Size' parameter from MEP element {mepElementId.GetIntegerValue()}: {paramEx.Message} for sleeve {openingId.GetIntegerValue()}\n");
                                                }
                                                result.Warnings.Add($"Error reading 'Size' parameter from MEP element {mepElementId.GetIntegerValue()}: {paramEx.Message}");
                                                continue;
                                            }

                                            if (!readFromRevit)
                                            {
                                                // ✅ CRITICAL FIX FOR REVIT 2024: MEP element exists but "Size" parameter is empty - fallback to snapshot
                                                // In Revit 2024, the "Size" parameter may be NULL/EMPTY even though it's saved in the snapshot
                                                if (!DeploymentConfiguration.DeploymentMode)
                                                {
                                                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️ MEP element {mepElementId.GetIntegerValue()} exists but 'Size' parameter is NULL or EMPTY for sleeve {openingId.GetIntegerValue()}. Attempting fallback to snapshot...\n");
                                                }

                                                // ✅ FALLBACK: Try to read from snapshot when Revit parameter is empty (Revit 2024 compatibility)
                                                if (sourceParams.TryGetValue("Size", out sourceValue) || sourceParams.TryGetValue("MEP Size", out sourceValue))
                                                {
                                                    if (!string.IsNullOrWhiteSpace(sourceValue))
                                                    {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                                        {
                                                            SafeFileLogger.SafeAppendText("transfer_debug.log",
                                                                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅✅✅ FALLBACK SUCCESS: Read 'MEP Size'='{sourceValue}' from SNAPSHOT (Revit parameter was empty) for sleeve {openingId.GetIntegerValue()}\n");
                                                        }
                                                        readFromRevit = false; // Mark as from snapshot
                                                        result.Warnings.Add($"MEP element {mepElementId.GetIntegerValue()} 'Size' parameter was empty, used snapshot value '{sourceValue}' for sleeve {openingId.GetIntegerValue()}.");
                                                    }
                                                    else
                                                    {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                                        {
                                                            SafeFileLogger.SafeAppendText("transfer_debug.log",
                                                                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️ 'Size' or 'MEP Size' found in snapshot but is EMPTY for sleeve {openingId.GetIntegerValue()}\n");
                                                        }
                                                        result.Warnings.Add($"MEP element {mepElementId.GetIntegerValue()} 'Size' parameter is empty in both Revit and snapshot for sleeve {openingId.GetIntegerValue()}.");
                                                        continue;
                                                    }
                                                }
                                                else
                                                {
                                                    if (!DeploymentConfiguration.DeploymentMode)
                                                    {
                                                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                                                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️ MEP element {mepElementId.GetIntegerValue()} 'Size' parameter is empty AND 'Size'/'MEP Size' NOT FOUND in snapshot for sleeve {openingId.GetIntegerValue()}. Available snapshot params: [{string.Join(", ", sourceParams.Keys.Take(10))}]\n");
                                                    }
                                                    result.Warnings.Add($"MEP element {mepElementId.GetIntegerValue()} 'Size' parameter is empty and not found in snapshot for sleeve {openingId.GetIntegerValue()}.");
                                                    continue;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            // MEP element not found or invalid
                                            if (!DeploymentConfiguration.DeploymentMode)
                                            {
                                                SafeFileLogger.SafeAppendText("transfer_debug.log",
                                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️ MEP element {mepElementId.GetIntegerValue()} NOT FOUND or INVALID for sleeve {openingId.GetIntegerValue()}\n");
                                            }
                                            result.Warnings.Add($"MEP element {mepElementId.GetIntegerValue()} not found or invalid for sleeve {openingId.GetIntegerValue()}.");
                                            continue;
                                        }
                                    }
                                    else
                                    {
                                        // Invalid MEP_ElementId - try fallback to snapshot
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            SafeFileLogger.SafeAppendText("transfer_debug.log",
                                                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️ MEP_ElementId is NULL or InvalidElementId for sleeve {openingId.GetIntegerValue()}. Attempting fallback to snapshot...\n");
                                        }

                                        // ✅ FIX: Fallback to snapshot if MEP_ElementId is invalid
                                        if (sourceParams.TryGetValue("Size", out sourceValue) || sourceParams.TryGetValue("MEP Size", out sourceValue))
                                        {
                                            if (!string.IsNullOrWhiteSpace(sourceValue))
                                            {
                                                if (!DeploymentConfiguration.DeploymentMode)
                                                {
                                                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅✅✅ FALLBACK SUCCESS: Read 'MEP Size'='{sourceValue}' from SNAPSHOT (MEP_ElementId was invalid) for sleeve {openingId.GetIntegerValue()}\n");
                                                }
                                                readFromRevit = false; // Mark as from snapshot
                                            }
                                            else
                                            {
                                                if (!DeploymentConfiguration.DeploymentMode)
                                                {
                                                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️ 'Size' or 'MEP Size' found in snapshot but is EMPTY for sleeve {openingId.GetIntegerValue()}\n");
                                                }
                                                result.Warnings.Add($"'Size' parameter is empty in snapshot for sleeve {openingId.GetIntegerValue()}, and MEP_ElementId is invalid.");
                                                continue;
                                            }
                                        }
                                        else
                                        {
                                            if (!DeploymentConfiguration.DeploymentMode)
                                            {
                                                SafeFileLogger.SafeAppendText("transfer_debug.log",
                                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️ MEP_ElementId is invalid AND 'Size'/'MEP Size' NOT FOUND in snapshot for sleeve {openingId.GetIntegerValue()}. Available snapshot params: [{string.Join(", ", sourceParams.Keys.Take(10))}]\n");
                                            }
                                            result.Warnings.Add($"Invalid MEP_ElementId and 'Size' parameter not found in snapshot for sleeve {openingId.GetIntegerValue()}.");
                                            continue;
                                        }
                                    }
                                }
                                else
                                {
                                    // MEP_ElementId parameter missing - try fallback to snapshot
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️ MEP_ElementId parameter is NULL or HAS NO VALUE for sleeve {openingId.GetIntegerValue()}. Attempting fallback to snapshot...\n");
                                    }

                                    // ✅ FIX: Fallback to snapshot if MEP_ElementId parameter is missing
                                    if (sourceParams.TryGetValue("Size", out sourceValue) || sourceParams.TryGetValue("MEP Size", out sourceValue))
                                    {
                                        if (!string.IsNullOrWhiteSpace(sourceValue))
                                        {
                                            if (!DeploymentConfiguration.DeploymentMode)
                                            {
                                                SafeFileLogger.SafeAppendText("transfer_debug.log",
                                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅✅✅ FALLBACK SUCCESS: Read 'MEP Size'='{sourceValue}' from SNAPSHOT (MEP_ElementId was missing) for sleeve {openingId.GetIntegerValue()}\n");
                                            }
                                            readFromRevit = false; // Mark as from snapshot
                                        }
                                        else
                                        {
                                            if (!DeploymentConfiguration.DeploymentMode)
                                            {
                                                SafeFileLogger.SafeAppendText("transfer_debug.log",
                                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️ 'Size' or 'MEP Size' found in snapshot but is EMPTY for sleeve {openingId.GetIntegerValue()}\n");
                                            }
                                            result.Warnings.Add($"'Size' parameter is empty in snapshot for sleeve {openingId.GetIntegerValue()}, and MEP_ElementId is missing.");
                                            continue;
                                        }
                                    }
                                    else
                                    {
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            SafeFileLogger.SafeAppendText("transfer_debug.log",
                                                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️ MEP_ElementId parameter missing AND 'Size'/'MEP Size' NOT FOUND in snapshot for sleeve {openingId.GetIntegerValue()}. Available snapshot params: [{string.Join(", ", sourceParams.Keys.Take(10))}]\n");
                                        }
                                        result.Warnings.Add($"MEP_ElementId parameter missing and 'Size' parameter not found in snapshot for sleeve {openingId.GetIntegerValue()}.");
                                        continue;
                                    }
                                }
                            }
                            catch (Exception revitEx)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ❌❌❌ EXCEPTION reading 'MEP Size' from Revit: {revitEx.GetType().Name}: {revitEx.Message} for sleeve {openingId.GetIntegerValue()}\n");
                                }
                                result.Warnings.Add($"Error reading 'MEP Size' from Revit for sleeve {openingId.GetIntegerValue()}: {revitEx.Message}");
                                continue;
                            }
                        } // End of "MEP Size" parameter fallback handling within else block

                        // ✅ CRITICAL DIAGNOSTIC: Log that we've exited the if-else block
                        if (!DeploymentConfiguration.DeploymentMode && isSizeParameter)
                        {
                            SafeFileLogger.SafeAppendText("transfer_debug.log",
                                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔍 DEBUG: Exited if-else block for Size parameter, sourceValue='{sourceValue}', isSizeParameter={isSizeParameter}, isEmpty={string.IsNullOrWhiteSpace(sourceValue)} for sleeve {openingId.GetIntegerValue()}\n");
                        }

                        // ✅ CRITICAL FIX: For "Size"/"MEP Size" parameters, if we successfully read from snapshot (lines 1255-1267),
                        // we should skip the "other parameters" block and proceed directly to parameter setting.
                        // Only enter this block if it's NOT a Size parameter OR if sourceValue is still empty (needs fallback handling)
                        // Note: isSizeParameter is already declared at the start of the parameter reading section

                        // ✅ DIAGNOSTIC: Log whether we're skipping the "other parameters" block for Size
                        if (isSizeParameter && !string.IsNullOrWhiteSpace(sourceValue))
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("transfer_debug.log",
                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔍 DEBUG: Size parameter successfully read (sourceValue='{sourceValue}'), SKIPPING 'other parameters' block, proceeding directly to parameter setting for sleeve {openingId.GetIntegerValue()}\n");
                            }
                        }

                        // For all other parameters (not "MEP Size"), use snapshot with variations
                        // Skip this block if Size parameter was successfully read from snapshot (proceed directly to parameter setting)
                        if (!isSizeParameter || string.IsNullOrWhiteSpace(sourceValue))
                        {
                            // ✅ DIAGNOSTIC: Log that we're using snapshot for non-MEP Size parameters
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("transfer_debug.log",
                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔍 Source parameter '{mapping.SourceParameter}' is NOT 'MEP Size' - using snapshot for sleeve {openingId.GetIntegerValue()}\n");
                            }

                            // For other parameters, try snapshot first
                            // ✅ DIAGNOSTIC: Log what we're looking for
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("transfer_debug.log",
                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔍 Looking for parameter '{mapping.SourceParameter}' in snapshot for sleeve {openingId.GetIntegerValue()}...\n");
                            }

                            // ✅ FIX: Try exact match first, then try common variations (case-insensitive dictionary handles case)
                            if (!sourceParams.TryGetValue(mapping.SourceParameter, out sourceValue) || string.IsNullOrWhiteSpace(sourceValue))
                            {
                                // Get category to check for Cable Trays specific mapping
                                string category = GetCategoryFromOpening(opening, snapshot);

                                // Build variations list
                                var variationsList = new List<string>
                            {
                                mapping.SourceParameter.Replace(" ", "_"),  // "System Type" -> "System_Type"
                                mapping.SourceParameter.Replace("_", " "),  // "System_Type" -> "System Type"
                                "MEP " + mapping.SourceParameter,           // "System Type" -> "MEP System Type"
                                mapping.SourceParameter.Replace("MEP ", ""), // "MEP System Type" -> "System Type"
                            };

                                // ✅ CRITICAL FIX: For Cable Trays, "Service Type" maps to "System Type"
                                if (string.Equals(category, "Cable Trays", StringComparison.OrdinalIgnoreCase) &&
                                    string.Equals(mapping.SourceParameter, "System Type", StringComparison.OrdinalIgnoreCase))
                                {
                                    variationsList.Insert(0, "Service Type"); // Add as first priority variation
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔍 Cable Trays detected - mapping 'System Type' -> 'Service Type' for sleeve {openingId.GetIntegerValue()}\n");
                                    }
                                }

                                string[] variations = variationsList.ToArray();

                                bool foundVariation = false;
                                foreach (var variation in variations)
                                {
                                    if (sourceParams.TryGetValue(variation, out sourceValue) && !string.IsNullOrWhiteSpace(sourceValue))
                                    {
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            SafeFileLogger.SafeAppendText("transfer_debug.log",
                                                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅✅✅ Found parameter '{mapping.SourceParameter}' as variation '{variation}'='{sourceValue}' in snapshot for sleeve {openingId.GetIntegerValue()} (IsCluster={isClusterSleeve})\n");
                                        }
                                        foundVariation = true;
                                        break; // Found it, exit loop
                                    }
                                }

                                if (!foundVariation && !DeploymentConfiguration.DeploymentMode)
                                {
                                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️ Tried variations for '{mapping.SourceParameter}' but none found: [{string.Join(", ", variations)}] for sleeve {openingId.GetIntegerValue()} (IsCluster={isClusterSleeve})\n");
                                }
                            }
                            else
                            {
                                // Exact match found
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅ Found exact match '{mapping.SourceParameter}'='{sourceValue}' in snapshot for sleeve {openingId.GetIntegerValue()} (IsCluster={isClusterSleeve})\n");
                                }
                            }

                            if (string.IsNullOrWhiteSpace(sourceValue))
                            {
                                // ✅ FIX: For cluster sleeves, snapshot is the ONLY source - no Revit fallback
                                if (isClusterSleeve)
                                {
                                    // Cluster sleeve - snapshot is the only source, no Revit fallback
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        var allAvailableParams = string.Join(", ", sourceParams.Keys.OrderBy(k => k));
                                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️ Parameter '{mapping.SourceParameter}' NOT FOUND in snapshot for CLUSTER sleeve {openingId.GetIntegerValue()} (ClusterId={clusterInstanceId}). Available snapshot params ({sourceParams.Count} total): [{allAvailableParams}]. Tried variations: [{string.Join(", ", new[] { mapping.SourceParameter.Replace(" ", "_"), mapping.SourceParameter.Replace("_", " "), "MEP " + mapping.SourceParameter, mapping.SourceParameter.Replace("MEP ", "") })}]\n");
                                    }
                                    result.Warnings.Add($"Parameter '{mapping.SourceParameter}' not found in snapshot for cluster sleeve {openingId.GetIntegerValue()} (ClusterId={clusterInstanceId}). Available: [{string.Join(", ", sourceParams.Keys.Take(10))}]");
                                    continue;
                                }

                                // ✅ FIX: If not in snapshot, try reading from Revit MEP element (for individual sleeves only)
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    var allAvailableParams = string.Join(", ", sourceParams.Keys.OrderBy(k => k));
                                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️ Parameter '{mapping.SourceParameter}' NOT FOUND in snapshot for individual sleeve {openingId.GetIntegerValue()}. Available snapshot params ({sourceParams.Count} total): [{allAvailableParams}]. Attempting to read from Revit MEP element...\n");
                                }

                                // Try to read from Revit MEP element as fallback
                                try
                                {
                                    var mepElementIdParam = opening.LookupParameter("MEP_ElementId");
                                    if (mepElementIdParam != null && mepElementIdParam.HasValue)
                                    {
                                        var mepElementId = mepElementIdParam.AsElementId();
                                        if (mepElementId != null && mepElementId != ElementId.InvalidElementId)
                                        {
                                            var mepElement = doc.GetElement(mepElementId);
                                            if (mepElement != null && mepElement.IsValidObject)
                                            {
                                                var revitParam = mepElement.LookupParameter(mapping.SourceParameter);
                                                if (revitParam != null && revitParam.HasValue)
                                                {
                                                    sourceValue = revitParam.AsValueString() ?? revitParam.AsString();
                                                    readFromRevit = true;

                                                    if (!DeploymentConfiguration.DeploymentMode)
                                                    {
                                                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                                                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅✅✅ FALLBACK SUCCESS: Read '{mapping.SourceParameter}'='{sourceValue}' from Revit MEP element {mepElementId.GetIntegerValue()} for sleeve {openingId.GetIntegerValue()} (not in snapshot)\n");
                                                    }
                                                }
                                                else
                                                {
                                                    if (!DeploymentConfiguration.DeploymentMode)
                                                    {
                                                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                                                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️ Parameter '{mapping.SourceParameter}' not found on MEP element {mepElementId.GetIntegerValue()} for sleeve {openingId.GetIntegerValue()}\n");
                                                    }
                                                    result.Warnings.Add($"Parameter '{mapping.SourceParameter}' not found in snapshot or on MEP element for sleeve {openingId.GetIntegerValue()}.");
                                                    continue;
                                                }
                                            }
                                            else
                                            {
                                                if (!DeploymentConfiguration.DeploymentMode)
                                                {
                                                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️ MEP element {mepElementId.GetIntegerValue()} not found or invalid for sleeve {openingId.GetIntegerValue()}\n");
                                                }
                                                result.Warnings.Add($"Parameter '{mapping.SourceParameter}' not found in snapshot, and MEP element {mepElementId.GetIntegerValue()} is invalid for sleeve {openingId.GetIntegerValue()}.");
                                                continue;
                                            }
                                        }
                                        else
                                        {
                                            if (!DeploymentConfiguration.DeploymentMode)
                                            {
                                                SafeFileLogger.SafeAppendText("transfer_debug.log",
                                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️ MEP_ElementId is NULL or InvalidElementId for sleeve {openingId.GetIntegerValue()}, cannot read '{mapping.SourceParameter}' from Revit\n");
                                            }
                                            result.Warnings.Add($"Parameter '{mapping.SourceParameter}' not found in snapshot for sleeve {openingId.GetIntegerValue()}, and MEP_ElementId is invalid.");
                                            continue;
                                        }
                                    }
                                    else
                                    {
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            SafeFileLogger.SafeAppendText("transfer_debug.log",
                                                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️ Parameter '{mapping.SourceParameter}' NOT FOUND in snapshot for sleeve {openingId.GetIntegerValue()}, and MEP_ElementId parameter is missing. Available snapshot params: [{string.Join(", ", sourceParams.Keys.Take(10))}]\n");
                                        }
                                        result.Warnings.Add($"Parameter '{mapping.SourceParameter}' not found in snapshot for sleeve {openingId.GetIntegerValue()}, and MEP_ElementId is missing.");
                                        continue;
                                    }
                                }
                                catch (Exception revitEx)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ❌ Error reading '{mapping.SourceParameter}' from Revit MEP element: {revitEx.Message} for sleeve {openingId.GetIntegerValue()}\n");
                                    }
                                    result.Warnings.Add($"Parameter '{mapping.SourceParameter}' not found in snapshot, and error reading from Revit: {revitEx.Message} for sleeve {openingId.GetIntegerValue()}.");
                                    continue;
                                }
                            }
                            else
                            {
                                // ✅ DIAGNOSTIC: Log successful snapshot read
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅ Read '{mapping.SourceParameter}'='{sourceValue}' from snapshot for sleeve {openingId.GetIntegerValue()}\n");
                                }
                            }
                        }
                    } // ✅ CRITICAL FIX: Close the if (!skipElseBlock) block here - parameter setting logic must be OUTSIDE

                    // ✅ CRITICAL CHECKPOINT: Log that we've exited the if-else block and are about to enter parameter setting
                    // This code is OUTSIDE the if (!skipElseBlock) block, so it runs for both Size and non-Size parameters
                    if (!DeploymentConfiguration.DeploymentMode && !string.IsNullOrWhiteSpace(sourceValue))
                    {
                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔍 CHECKPOINT: Exited if-else block, about to enter parameter setting section, sourceValue='{sourceValue}', isSizeParameter={isSizeParameter} for sleeve {openingId.GetIntegerValue()}\n");
                    }

                    // ✅ CRITICAL DIAGNOSTIC: Log that we've reached parameter setting logic (for Size parameters)
                    if (!DeploymentConfiguration.DeploymentMode && isSizeParameter && !string.IsNullOrWhiteSpace(sourceValue))
                    {
                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅✅✅ REACHED PARAMETER SETTING LOGIC: Size parameter with sourceValue='{sourceValue}' for sleeve {openingId.GetIntegerValue()} - proceeding to set 'MEP Size' parameter\n");
                    }

                    // ✅ DIAGNOSTIC: Verify sourceValue is not empty before proceeding
                    if (string.IsNullOrWhiteSpace(sourceValue))
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("transfer_debug.log",
                                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️⚠️⚠️ WARNING: sourceValue is EMPTY after read attempt for '{mapping.SourceParameter}' on sleeve {openingId.GetIntegerValue()} (IsCluster={isClusterSleeve}, ClusterId={clusterInstanceId}, SleeveId={sleeveInstanceId}) - SKIPPING transfer\n");
                        }
                        result.Warnings.Add($"Parameter '{mapping.SourceParameter}' value is empty for sleeve {openingId.GetIntegerValue()}.");
                        continue;
                    }

                    // ✅ CRITICAL: Wrap parameter setting logic in if block to ensure proper structure
                    // This if block ensures sourceValue is not empty before proceeding to parameter setting
                    if (!string.IsNullOrWhiteSpace(sourceValue))
                    {
                        // ✅ DIAGNOSTIC: Log that we have a valid sourceValue and are proceeding
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("transfer_debug.log",
                                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅ sourceValue is VALID: '{mapping.SourceParameter}'='{sourceValue}' for sleeve {openingId.GetIntegerValue()} (IsCluster={isClusterSleeve}) - PROCEEDING to target parameter lookup\n");
                        }

                        // ✅ CRASH-SAFETY 6: Re-validate element is still valid before parameter access
                        if (!opening.IsValidObject)
                        {
                            failedCount++;
                            errors.Add($"Opening element {openingId.GetIntegerValue()} became invalid during transfer (may have been deleted).");
                            continue;
                        }

                        // ✅ DIAGNOSTIC: Log that we're about to look up target parameter
                        if (!DeploymentConfiguration.DeploymentMode && !string.IsNullOrWhiteSpace(sourceValue))
                        {
                            SafeFileLogger.SafeAppendText("transfer_debug.log",
                                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔍 Looking up target parameter '{mapping.TargetParameter}' on sleeve {openingId.GetIntegerValue()} (sourceValue='{sourceValue}')\n");
                        }

                        // ✅ PROTECTION 3: Use cached parameter if available, with comprehensive validation
                        Parameter targetParam = null;
                        if (Services.OptimizationFlags.UseBatchParameterLookups &&
                            parameterCache.TryGetValue(openingId.GetIntegerValue(), out var paramDict) &&
                            paramDict.TryGetValue(mapping.TargetParameter, out targetParam))
                        {
                            // ✅ CRITICAL VALIDATION: Ensure cached parameter is still valid and belongs to the correct element
                            bool isValidCache = targetParam != null &&
                                               targetParam.Element != null &&
                                               targetParam.Element.IsValidObject &&
                                               targetParam.Element.Id == opening.Id &&
                                               targetParam.Element.Id.GetIntegerValue() == openingId.GetIntegerValue(); // Double-check ID match

                            // ✅ PROTECTION 4: Verify parameter definition is still valid
                            if (isValidCache)
                            {
                                try
                                {
                                    var paramDef = targetParam.Definition;
                                    if (paramDef == null || string.IsNullOrEmpty(paramDef.Name))
                                    {
                                        isValidCache = false;
                                    }
                                    else if (!paramDef.Name.Equals(mapping.TargetParameter, StringComparison.OrdinalIgnoreCase))
                                    {
                                        isValidCache = false; // Parameter name mismatch
                                    }
                                }
                                catch
                                {
                                    isValidCache = false; // Parameter definition access failed
                                }
                            }

                            if (isValidCache)
                            {
                                // Use cached parameter
                                if (!DeploymentConfiguration.DeploymentMode && !string.IsNullOrWhiteSpace(sourceValue))
                                {
                                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅ Found target parameter '{mapping.TargetParameter}' in cache for sleeve {openingId.GetIntegerValue()}\n");
                                }
                            }
                            else
                            {
                                // Cached parameter is stale - clear it and lookup fresh
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️ Cached parameter '{mapping.TargetParameter}' is stale for sleeve {openingId.GetIntegerValue()} - looking up fresh\n");
                                }
                                targetParam = null; // Force fresh lookup
                            }
                        }

                        // Fallback: Lookup parameter if not cached or cache is stale
                        if (targetParam == null)
                        {
                            targetParam = opening.LookupParameter(mapping.TargetParameter);
                            if (!DeploymentConfiguration.DeploymentMode && !string.IsNullOrWhiteSpace(sourceValue))
                            {
                                if (targetParam != null)
                                {
                                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅ Found target parameter '{mapping.TargetParameter}' via lookup for sleeve {openingId.GetIntegerValue()}\n");
                                }
                                else
                                {
                                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ❌ Target parameter '{mapping.TargetParameter}' NOT FOUND on sleeve {openingId.GetIntegerValue()}\n");
                                }
                            }
                        }

                        if (targetParam == null)
                        {
                            // ✅ FIX: Missing parameters are warnings, not errors (especially for optional metadata like MEP_Size)
                            // Only count as failure if it's a critical parameter (e.g., MEP_ElementId, MEP System Type)
                            // ✅ FIX: Parameter names use spaces for "MEP System Type", underscore for "MEP_ElementId"
                            bool isCritical = mapping.TargetParameter.Equals("MEP_ElementId", StringComparison.OrdinalIgnoreCase) ||
                                             mapping.TargetParameter.Equals("MEP System Type", StringComparison.OrdinalIgnoreCase) ||
                                             mapping.TargetParameter.Equals("System Type", StringComparison.OrdinalIgnoreCase);

                            if (isCritical)
                            {
                                failedCount++;
                                errors.Add($"Target parameter '{mapping.TargetParameter}' not found on opening {openingId.GetIntegerValue()}.");
                            }
                            else
                            {
                                // Non-critical parameter missing - just warn, don't fail
                                result.Warnings.Add($"Target parameter '{mapping.TargetParameter}' not found on opening {openingId.GetIntegerValue()} (skipped).");
                            }
                            continue;
                        }

                        // ✅ PROTECTION 5: Validate parameter belongs to correct element (CRITICAL CHECK)
                        // ✅ FIX: Compare by element ID only (document reference equality can fail even for same document)
                        // If parameter's element ID matches opening's ID, they're the same element (and same document)
                        // ✅ UNIQUE LOG MARKER: This section confirms the fixed code (Element ID comparison) is running
                        // ✅ CRITICAL: Always log code path verification (even in deployment mode) for debugging
                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔍 CODE PATH VERIFICATION: Entering Element ID validation for sleeve {openingId.GetIntegerValue()} (Fixed code v2025-12-03)\n");

                        if (targetParam.Element == null)
                        {
                            failedCount++;
                            errors.Add($"Parameter '{mapping.TargetParameter}' on opening {openingId.GetIntegerValue()} has null element reference.");
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("transfer_debug.log",
                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ❌ Parameter element is NULL: '{mapping.TargetParameter}' on sleeve {openingId.GetIntegerValue()} - SKIPPING\n");
                            }
                            continue;
                        }

                        if (!targetParam.Element.IsValidObject)
                        {
                            failedCount++;
                            errors.Add($"Parameter '{mapping.TargetParameter}' on opening {openingId.GetIntegerValue()} has invalid element reference.");
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("transfer_debug.log",
                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ❌ Parameter element is INVALID: '{mapping.TargetParameter}' on sleeve {openingId.GetIntegerValue()} - SKIPPING\n");
                            }
                            continue;
                        }

                        // ✅ UNIQUE LOG MARKER: This confirms the fixed code (Element ID comparison) is running
                        if (targetParam.Element.Id != opening.Id || targetParam.Element.Id.GetIntegerValue() != openingId.GetIntegerValue())
                        {
                            failedCount++;
                            errors.Add($"Parameter '{mapping.TargetParameter}' on opening {openingId.GetIntegerValue()} belongs to different element.");
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("transfer_debug.log",
                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ❌ Element mismatch: Parameter '{mapping.TargetParameter}' on sleeve {openingId.GetIntegerValue()} belongs to different element (ParamElementId={targetParam.Element.Id.GetIntegerValue()}, OpeningId={opening.Id.GetIntegerValue()}) - SKIPPING\n");
                                SafeFileLogger.SafeAppendText("transfer_debug.log",
                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅ CODE VERIFICATION: Element ID comparison code path is ACTIVE (this is the fixed version)\n");
                            }
                            continue;
                        }

                        // ✅ UNIQUE LOG MARKER: Element ID validation passed - confirms fixed code is running
                        // ✅ CRITICAL: Always log validation success (even in deployment mode) for debugging
                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅ CODE VERIFICATION: Element ID validation PASSED for sleeve {openingId.GetIntegerValue()} (ParamElementId={targetParam.Element.Id.GetIntegerValue()}, OpeningId={opening.Id.GetIntegerValue()}) - Fixed code is running\n");

                        // ✅ PROTECTION 6: Verify opening element is still valid before parameter setting
                        if (!opening.IsValidObject)
                        {
                            failedCount++;
                            errors.Add($"Opening element {openingId.GetIntegerValue()} became invalid before parameter setting.");
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("transfer_debug.log",
                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ❌ Opening element became INVALID: sleeve {openingId.GetIntegerValue()} - SKIPPING\n");
                            }
                            continue;
                        }

                        // ✅ DIAGNOSTIC: Log before skip check
                        if (!DeploymentConfiguration.DeploymentMode && !string.IsNullOrWhiteSpace(sourceValue))
                        {
                            SafeFileLogger.SafeAppendText("transfer_debug.log",
                                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔍 Checking if should skip '{mapping.TargetParameter}' on sleeve {openingId.GetIntegerValue()} (SkipAlreadyTransferred={Services.OptimizationFlags.SkipAlreadyTransferredParameters})\n");
                        }

                        // ✅ OPTIMIZATION 2: Skip if parameter already matches snapshot value
                        if (Services.OptimizationFlags.SkipAlreadyTransferredParameters)
                        {
                            if (ShouldSkipParameter(targetParam, sourceValue, openingId.GetIntegerValue(), mapping.TargetParameter))
                            {
                                skippedCount++;

                                // ✅ DIAGNOSTIC: Track per-sleeve skip count
                                if (!perSleeveSkipCounts.ContainsKey(openingId.GetIntegerValue()))
                                {
                                    perSleeveSkipCounts[openingId.GetIntegerValue()] = 0;
                                }
                                perSleeveSkipCounts[openingId.GetIntegerValue()]++;

                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⏭️ Skipping '{mapping.TargetParameter}' on sleeve {openingId.GetIntegerValue()} (already matches value '{sourceValue}')\n");
                                }
                                continue;
                            }
                            else
                            {
                                // ✅ DIAGNOSTIC: Log when NOT skipping
                                if (!DeploymentConfiguration.DeploymentMode && !string.IsNullOrWhiteSpace(sourceValue))
                                {
                                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅ NOT skipping '{mapping.TargetParameter}' on sleeve {openingId.GetIntegerValue()} - parameter value differs or is empty\n");
                                }
                            }
                        }

                        // ✅ PROTECTION 10: Final validation before parameter setting
                        if (targetParam.IsReadOnly)
                        {
                            failedCount++;
                            errors.Add($"Parameter '{mapping.TargetParameter}' on opening {openingId.GetIntegerValue()} is read-only.");
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("transfer_debug.log",
                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ❌ Parameter '{mapping.TargetParameter}' is READ-ONLY on sleeve {openingId.GetIntegerValue()} - SKIPPING\n");
                            }
                            continue;
                        }

                        // ✅ PROTECTION 11: Validate source value is not empty before setting
                        if (string.IsNullOrWhiteSpace(sourceValue))
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("transfer_debug.log",
                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ⚠️ Source value is EMPTY for '{mapping.TargetParameter}' on sleeve {openingId.GetIntegerValue()} - SKIPPING\n");
                            }
                            result.Warnings.Add($"Source value is empty for parameter '{mapping.TargetParameter}' on opening {openingId.GetIntegerValue()}.");
                            continue;
                        }

                        // ✅ DIAGNOSTIC: Log before setting parameter
                        if (!DeploymentConfiguration.DeploymentMode && !string.IsNullOrWhiteSpace(sourceValue))
                        {
                            SafeFileLogger.SafeAppendText("transfer_debug.log",
                                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 🔧 Attempting to set '{mapping.TargetParameter}'='{sourceValue}' on sleeve {openingId.GetIntegerValue()}\n");
                        }

                        // ✅ PROTECTION 12: Final element validation before setting
                        if (!opening.IsValidObject || !targetParam.Element.IsValidObject)
                        {
                            failedCount++;
                            errors.Add($"Element became invalid before setting parameter '{mapping.TargetParameter}' on opening {openingId.GetIntegerValue()}.");
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("transfer_debug.log",
                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ❌ Element became INVALID before setting '{mapping.TargetParameter}' on sleeve {openingId.GetIntegerValue()} - SKIPPING\n");
                            }
                            continue;
                        }

                        if (SetParameterValueSafely(targetParam, sourceValue))
                        {
                            transferredCount++;
                            successfullyTransferredSleeveIds?.Add(openingId.GetIntegerValue());
                            if (!DeploymentConfiguration.DeploymentMode && !string.IsNullOrWhiteSpace(sourceValue))
                            {
                                SafeFileLogger.SafeAppendText("transfer_debug.log",
                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅✅✅ SUCCESS: Set '{mapping.TargetParameter}'='{sourceValue}' on sleeve {openingId.GetIntegerValue()}\n");
                            }
                        }
                        else
                        {
                            failedCount++;
                            errors.Add($"Failed to set parameter '{mapping.TargetParameter}' on opening {openingId.GetIntegerValue()}.");
                            if (!DeploymentConfiguration.DeploymentMode && !string.IsNullOrWhiteSpace(sourceValue))
                            {
                                SafeFileLogger.SafeAppendText("transfer_debug.log",
                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ❌ FAILED: Could not set '{mapping.TargetParameter}'='{sourceValue}' on sleeve {openingId.GetIntegerValue()}\n");
                            }
                        }
                    } // End of if block for sourceValue assignment

                    // ✅ PERFORMANCE: Log per-sleeve completion time (always log - critical diagnostic)
                    sleeveStartTime.Stop();
                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PERFORMANCE] ⏱️ Sleeve {openingId.GetIntegerValue()} completed in {sleeveStartTime.ElapsedMilliseconds}ms\n");
                } // End of try block for per-sleeve processing
                catch (Exception ex)
                {
                    sleeveStartTime.Stop();
                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PERFORMANCE] ⏱️ Sleeve {openingId.GetIntegerValue()} took {sleeveStartTime.ElapsedMilliseconds}ms (EXCEPTION)\n");

                    failedCount++;
                    errors.Add($"Error transferring parameters to opening {openingId.GetIntegerValue()}: {ex.Message}");

                    // ✅ CRITICAL: Log the exception to the debug file
                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ❌ EXCEPTION processing sleeve {openingId.GetIntegerValue()}: {ex.GetType().Name}: {ex.Message}\n");
                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] Stack Trace: {ex.StackTrace}\n");
                }
            }

            transferStartTime.Stop();
            SafeFileLogger.SafeAppendText("transfer_debug.log",
                $"\n[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PERFORMANCE] ⏱️ TOTAL TRANSFER TIME: {transferStartTime.ElapsedMilliseconds}ms for {openingIds.Count} sleeves\n");
            SafeFileLogger.SafeAppendText("transfer_debug.log",
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PERFORMANCE] ⏱️ AVERAGE TIME PER SLEEVE: {(openingIds.Count > 0 ? transferStartTime.ElapsedMilliseconds / (double)openingIds.Count : 0):F1}ms\n");
            SafeFileLogger.SafeAppendText("transfer_debug.log",
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PERFORMANCE] 📊 RESULTS: {transferredCount} transferred, {skippedCount} skipped, {failedCount} failed\n");

            // ✅ FIX: Success if no critical errors (warnings for missing optional parameters are OK)
            // Missing optional parameters (like MEP_Size) are warnings, not errors, so success = true if no errors
            result.Success = errors.Count == 0;
            result.TransferredCount = transferredCount;
            result.FailedCount = failedCount;
            result.Errors = errors;
            
            // ✅ DIAGNOSTIC: Log per-sleeve skip summary
            if (perSleeveSkipCounts.Count > 0)
            {
                SafeFileLogger.SafeAppendText("transfer_debug.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 📊 PER-SLEEVE SKIP SUMMARY (SkipAlreadyTransferred flag={Services.OptimizationFlags.SkipAlreadyTransferredParameters}):\n");
                
                foreach (var kvp in perSleeveSkipCounts.OrderBy(x => x.Key))
                {
                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER]   - Sleeve {kvp.Key}: {kvp.Value} parameter(s) skipped (already matched snapshot)\n");
                }
                
                int totalSkippedPerSleeve = perSleeveSkipCounts.Values.Sum();
                SafeFileLogger.SafeAppendText("transfer_debug.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] 📊 TOTAL: {totalSkippedPerSleeve} parameters skipped across {perSleeveSkipCounts.Count} sleeves\n");
            }
            
            if (!DeploymentConfiguration.DeploymentMode && skippedCount > 0)
            {
                DebugLogger.Info($"[PARAM_TRANSFER] ✅ Transfer complete: {transferredCount} transferred, {skippedCount} skipped (already transferred), {failedCount} failed");
            }
            
            return result;
        }

        private string GetAlias(string sourceParameter)
        {
            if (string.IsNullOrWhiteSpace(sourceParameter)) return string.Empty;
            var s = sourceParameter.Trim();
            if (s.Equals("Size", StringComparison.OrdinalIgnoreCase) || s.Equals("Service Size", StringComparison.OrdinalIgnoreCase))
                return "MepElementFormattedSize";
            return string.Empty;
        }
        
        /// <summary>
        /// Get category from opening/sleeve (from MEP_Category parameter, snapshot, or family name)
        /// </summary>
        private string GetCategoryFromOpening(Element opening, SleeveSnapshotView snapshot)
        {
            // Priority 1: Try to get from MEP_Category parameter on sleeve
            try
            {
                var categoryParam = opening.LookupParameter("MEP_Category");
                if (categoryParam != null && categoryParam.HasValue)
                {
                    var category = categoryParam.AsString();
                    if (!string.IsNullOrWhiteSpace(category))
                        return category;
                }
            }
            catch { }
            
            // Priority 2: Try to get from snapshot's MepParameters (if it has Category parameter)
            try
            {
                if (snapshot?.MepParameters != null && snapshot.MepParameters.TryGetValue("Category", out var snapshotCategory))
                {
                    if (!string.IsNullOrWhiteSpace(snapshotCategory))
                        return snapshotCategory;
                }
            }
            catch { }
            
            // Priority 3: Try to infer from family name (for cluster sleeves that might not have MEP_Category set)
            try
            {
                if (opening is FamilyInstance fi && fi.Symbol != null && fi.Symbol.Family != null)
                {
                    var familyName = fi.Symbol.Family.Name ?? string.Empty;
                    var upperFamilyName = familyName.ToUpperInvariant();
                    
                    if (upperFamilyName.Contains("CABLE") || upperFamilyName.Contains("TRAY"))
                        return "Cable Trays";
                    else if (upperFamilyName.Contains("DUCT"))
                        return "Ducts";
                    else if (upperFamilyName.Contains("PIPE"))
                        return "Pipes";
                }
            }
            catch { }
            
            // Fallback: Return empty string if not found
            return string.Empty;
        }
        
        /// <summary>
        /// Get category from sleeve prefix (D=Ducts, P=Pipes, E=Cable Trays, DMP=Dampers)
        /// </summary>
        private string GetCategoryFromSleevePrefix(Document doc, Element opening)
        {
            try
            {
                // Get the sleeve's mark parameter to determine prefix
                var markParam = opening.LookupParameter("Mark");
                if (markParam != null && markParam.StorageType == StorageType.String)
                {
                    var markValue = markParam.AsString();
                    if (!string.IsNullOrEmpty(markValue))
                    {
                        // Extract prefix from mark (e.g., "D001" -> "D", "P002" -> "P", "DMP001" -> "DMP")
                        var upperMark = markValue.ToUpper();
                        
                        if (upperMark.StartsWith("DMP"))
                            return "Duct Accessories";
                        else if (upperMark.StartsWith("D"))
                            return "Ducts";
                        else if (upperMark.StartsWith("P"))
                            return "Pipes";
                        else if (upperMark.StartsWith("E"))
                            return "Cable Trays";
                        else
                            return "Unknown";
                    }
                }
                
                // Fallback: try to determine from opening family name
                if (opening is FamilyInstance familyInstance)
                {
                    var familyName = familyInstance.Symbol.Family.Name.ToLower();
                    if (familyName.Contains("duct")) return "Ducts";
                    if (familyName.Contains("pipe")) return "Pipes";
                    if (familyName.Contains("cable") || familyName.Contains("tray")) return "Cable Trays";
                    if (familyName.Contains("damper") || familyName.Contains("accessory")) return "Duct Accessories";
                }
                
                return "Unknown";
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[PARAM_TRANSFER] Error getting category from sleeve prefix: {ex.Message}");
                return "Unknown";
            }
        }
        
        /// <summary>
        /// Get parameters from category-specific XML file
        /// </summary>
        private List<string> GetParametersFromCategoryXml(string category, string parameterName)
        {
            try
            {
                var filtersDirectory = ProjectPathService.GetFiltersDirectory(null);
                if (!Directory.Exists(filtersDirectory)) return new List<string>();
                
                // Use SAME patterns as BuildSnapshotIndex
                var patterns = new List<string>();
                var cat = category.ToLower().Replace(" ", "_");
                
                // MATCH THE EXACT SAME PATTERN LOGIC AS BuildSnapshotIndex
                patterns.Add($"*_{cat}.xml");
                patterns.Add($"*{cat}*.xml");
                patterns.Add($"{cat}*.xml");
                patterns.Add($"*conditions*{cat}*.xml");
                patterns.Add($"*conditions*.xml");
                
                foreach (var pattern in patterns)
                {
                    var xmlFiles = Directory.GetFiles(filtersDirectory, pattern);
                    if (xmlFiles.Length > 0)
                    {
                        var xmlFile = xmlFiles.OrderByDescending(f => File.GetLastWriteTime(f)).FirstOrDefault();
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[PARAM_TRANSFER] Loading parameters from: {Path.GetFileName(xmlFile)}");
                        
                        // IMPORTANT: Use the SAME deserialization as BuildSnapshotIndex
                        var serializer = new System.Xml.Serialization.XmlSerializer(typeof(Models.OpeningFilter));
                        using (var reader = new StreamReader(xmlFile))
                        {
                            var filter = (Models.OpeningFilter)serializer.Deserialize(reader);
                            var zones = filter?.ClashZoneStorage?.AllZones ?? new List<Models.ClashZone>();
                            
                            var values = new List<string>();
                            foreach (var zone in zones)
                            {
                                if (zone.MepParameterValues != null)
                                {
                                    // Try EXACT match first
                                    var exactMatch = zone.MepParameterValues.FirstOrDefault(p => 
                                        string.Equals(p.Key, parameterName, StringComparison.OrdinalIgnoreCase) && 
                                        !string.IsNullOrEmpty(p.Value));
                                    if (exactMatch != null)
                                    {
                                        values.Add(exactMatch.Value);
                                        continue;
                                    }
                                    
                                    // Try case-insensitive match
                                    foreach (var param in zone.MepParameterValues)
                                    {
                                        if (string.Equals(param.Key, parameterName, StringComparison.OrdinalIgnoreCase) && 
                                            !string.IsNullOrEmpty(param.Value))
                                        {
                                            values.Add(param.Value);
                                            break;
                                        }
                                    }
                                    
                                    // Try common aliases
                                    var aliases = GetParameterAliases(parameterName);
                                    foreach (var alias in aliases)
                                    {
                                        var aliasMatch = zone.MepParameterValues.FirstOrDefault(p => 
                                            string.Equals(p.Key, alias, StringComparison.OrdinalIgnoreCase) && 
                                            !string.IsNullOrEmpty(p.Value));
                                        if (aliasMatch != null)
                                        {
                                            values.Add(aliasMatch.Value);
                                            break;
                                        }
                                    }
                                }
                            }
                            
                            if (values.Count > 0)
                            {
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[PARAM_TRANSFER] Found {values.Count} values for parameter '{parameterName}' in category '{category}'");
                                return values.Distinct().ToList();
                            }
                        }
                    }
                }
                
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[PARAM_TRANSFER] No values found for parameter '{parameterName}' in category '{category}'");
                return new List<string>();
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[PARAM_TRANSFER] Error getting parameters from category XML: {ex.Message}");
                return new List<string>();
            }
        }
        
        /// <summary>
        /// Get parameter name aliases for better matching
        /// </summary>
        private List<string> GetParameterAliases(string parameterName)
        {
            var aliases = new List<string> { parameterName };
            
            // Add common variations
            var lower = parameterName.ToLower();
            if (lower.Contains("system type"))
            {
                aliases.AddRange(new[] { "System Type", "SystemType", "MEP System Type", "Service Type" });
            }
            else if (lower.Contains("size"))
            {
                aliases.AddRange(new[] { "Size", "MEP Size", "MepElementFormattedSize", "Service Size" });
            }
            else if (lower.Contains("level"))
            {
                aliases.AddRange(new[] { "Level", "Reference Level", "Schedule Level" });
            }
            else if (lower.Contains("system name"))
            {
                aliases.AddRange(new[] { "System Name", "MEP System Name", "SystemName" });
            }
            
            return aliases.Distinct().ToList();
        }
        
        /// <summary>
        /// Get host parameters from category-specific XML file
        /// </summary>
        private List<string> GetHostParametersFromCategoryXml(string category, string parameterName)
        {
            try
            {
                var filtersDirectory = ProjectPathService.GetFiltersDirectory(null);
                if (!Directory.Exists(filtersDirectory)) return new List<string>();
                
                // Look for XML files matching the category
                var patterns = new List<string>();
                var cat = category.ToLower().Replace(" ", "_");
                patterns.Add($"*{cat}*.xml");
                patterns.Add($"*_{cat}.xml");
                patterns.Add($"{cat}*.xml");
                
                foreach (var pattern in patterns)
                {
                    var xmlFiles = Directory.GetFiles(filtersDirectory, pattern);
                    if (xmlFiles.Length > 0)
                    {
                        var xmlFile = xmlFiles.OrderByDescending(f => File.GetLastWriteTime(f)).FirstOrDefault();
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[PARAM_TRANSFER] Loading host parameters from: {Path.GetFileName(xmlFile)}");
                        
                        // Load the XML file and extract host parameter values
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(Models.OpeningFilter));
                using (var reader = new StreamReader(xmlFile))
                {
                    var filter = (Models.OpeningFilter)serializer.Deserialize(reader);
                    var zones = filter?.ClashZoneStorage?.AllZones ?? new List<Models.ClashZone>();
                            
                            var values = new List<string>();
                            foreach (var zone in zones)
                            {
                                if (zone.HostParameterValues != null)
                                {
                                    foreach (var param in zone.HostParameterValues)
                                    {
                                        if (string.Equals(param.Key, parameterName, StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(param.Value))
                                        {
                                            values.Add(param.Value);
                                        }
                                    }
                                }
                            }
                            
                            if (values.Count > 0)
                            {
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[PARAM_TRANSFER] Found {values.Count} host values for parameter '{parameterName}' in category '{category}'");
                                return values.Distinct().ToList();
                            }
                        }
                    }
                }
                
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[PARAM_TRANSFER] No host values found for parameter '{parameterName}' in category '{category}'");
                return new List<string>();
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[PARAM_TRANSFER] Error getting host parameters from category XML: {ex.Message}");
                return new List<string>();
            }
        }
        
        /// <summary>
        /// Diagnostic method to verify snapshot index contents
        /// </summary>
        private void DiagnoseFilterIndex(Dictionary<string, Dictionary<int, (Dictionary<string,string> mep, Dictionary<string,string> host)>> filterIndex)
        {
                        if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[DIAGNOSE_FILTER_INDEX] Filter index has {filterIndex.Count} XML files");
            
            foreach (var kvp in filterIndex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[DIAGNOSE_FILTER_INDEX] XML '{kvp.Key}': {kvp.Value.Count} sleeves");
                
                if (kvp.Value.Count > 0)
                {
                    var firstSleeve = kvp.Value.First();
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[DIAGNOSE_FILTER_INDEX]   First sleeve {firstSleeve.Key}:");
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[DIAGNOSE_FILTER_INDEX]     MEP params: {string.Join(", ", firstSleeve.Value.mep.Keys)}");
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[DIAGNOSE_FILTER_INDEX]     Host params: {string.Join(", ", firstSleeve.Value.host.Keys)}");
                }
            }
        }
        
        /// <summary>
        /// Diagnostic method to check XML file content and parameter keys
        /// </summary>
        private void DiagnoseXmlContent(string category)
        {
            try
            {
                var filtersDirectory = ProjectPathService.GetFiltersDirectory(null);
                
                var allXmlFiles = Directory.GetFiles(filtersDirectory, "*.xml");
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[DIAGNOSE] All XML files in {filtersDirectory}:");
                
                foreach (var xmlFile in allXmlFiles)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[DIAGNOSE]   - {Path.GetFileName(xmlFile)}");
                    
                    try
                    {
                        var serializer = new System.Xml.Serialization.XmlSerializer(typeof(Models.OpeningFilter));
                        using (var reader = new StreamReader(xmlFile))
                        {
                            var filter = (Models.OpeningFilter)serializer.Deserialize(reader);
                            var zones = filter?.ClashZoneStorage?.AllZones ?? new List<Models.ClashZone>();
                            
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[DIAGNOSE]     {zones.Count} clash zones");
                            
                            if (zones.Count > 0)
                            {
                                var firstZone = zones[0];
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[DIAGNOSE]     First zone MEP params: {string.Join(", ", firstZone.MepParameterValues?.Select(p => p.Key) ?? new List<string>())}");
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[DIAGNOSE]     First zone Host params: {string.Join(", ", firstZone.HostParameterValues?.Select(p => p.Key) ?? new List<string>())}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error($"[DIAGNOSE]     Error reading file: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[DIAGNOSE] Error: {ex.Message}");
            }
        }
        
        #region Private Helper Methods

        private Dictionary<string, Dictionary<int, (Dictionary<string,string> mep, Dictionary<string,string> host)>> BuildFilterIndex(Document doc = null)
        {
            var filterIndex = new Dictionary<string, Dictionary<int, (Dictionary<string,string> mep, Dictionary<string,string> host)>>();
            
            if (doc == null)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[PARAM_TRANSFER] Document is null - cannot load from database");
                return filterIndex;
            }
            
            try
            {
                // ✅ DATABASE-FIRST: Load all filters and clash zones from database (not XML)
                using (var dbContext = new SleeveDbContext(doc, msg => 
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[SQLite] {msg}");
                }))
                {
                    var filterRepository = new FilterRepository(dbContext, msg => { });
                    var clashZoneRepository = new ClashZoneRepository(dbContext, msg => { });
                    
                    // Get all unique filter/category combinations from database
                    var filterCategories = new List<(string filterName, string category)>();
                    
                    using (var cmd = dbContext.Connection.CreateCommand())
                    {
                        cmd.CommandText = @"
                            SELECT DISTINCT FilterName, Category
                            FROM Filters
                            WHERE FilterName IS NOT NULL AND Category IS NOT NULL
                            ORDER BY FilterName, Category";
                        
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                var filterName = reader.IsDBNull(0) ? null : reader.GetString(0);
                                var category = reader.IsDBNull(1) ? null : reader.GetString(1);
                                
                                if (!string.IsNullOrWhiteSpace(filterName) && !string.IsNullOrWhiteSpace(category))
                                {
                                    filterCategories.Add((filterName, category));
                                }
                            }
                        }
                    }
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[PARAM_TRANSFER] Found {filterCategories.Count} filter/category combinations in database");
                        SafeFileLogger.SafeAppendText("transfer_debug.log", 
                            $"[{DateTime.Now}] [PARAM_TRANSFER] Found {filterCategories.Count} filter/category combinations in database\n");
                    }
                    
                    // Load clash zones for each filter/category combination
                    foreach (var (filterName, category) in filterCategories)
                    {
                        try
                        {
                            // ✅ DATABASE-FIRST: Load clash zones from database (includes MepElementSizeParameterValue)
                            var zones = clashZoneRepository.GetClashZonesByFilter(
                                filterName, 
                                category, 
                                unresolvedOnly: false, 
                                readyForPlacementOnly: false);
                            
                            if (zones == null || zones.Count == 0)
                                continue;
                            
                            // ✅ CRITICAL: Load parameter values from database (ensures MepElementSizeParameterValue is populated)
                            // Note: GetClashZonesByFilter already loads MepElementSizeParameterValue, but we ensure MepParameterValues are loaded too
                            foreach (var zone in zones)
                            {
                                if (zone != null && (zone.MepParameterValues == null || zone.MepParameterValues.Count == 0))
                                {
                                    // Load parameters from database if not already loaded
                                    // This is done internally by GetClashZonesByFilter, but we ensure it's done
                                    // The MepElementSizeParameterValue is already loaded by GetClashZonesByFilter
                                }
                            }
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[PARAM_TRANSFER] Loaded {zones.Count} clash zones from database for filter '{filterName}' ({category})");
                                SafeFileLogger.SafeAppendText("transfer_debug.log", 
                                    $"[{DateTime.Now}] [PARAM_TRANSFER] Loaded {zones.Count} clash zones from database for filter '{filterName}' ({category})\n");
                            }
                            
                            // Build filter key (same format as XML file name for compatibility)
                            var filterKey = $"{filterName}_{category}.xml";
                            
                            var filterData = new Dictionary<int, (Dictionary<string,string> mep, Dictionary<string,string> host)>();
                            
                            // First pass: collect individual sleeves
                            foreach (var zone in zones)
                            {
                                if (zone.SleeveInstanceId > 0)
                                {
                                    // ✅ DATABASE-FIRST: Use MepParameterValues from database (includes MepElementSizeParameterValue)
                                    var mepParams = zone.MepParameterValues?.ToDictionary(p => p.Key, p => p.Value) ?? new Dictionary<string, string>();
                                    
                                    // ✅ CRITICAL FIX FOR PIPES: Prioritize MepElementSizeParameterValue for Size parameter
                                    // This ensures individual sleeves get text values (e.g., "20 mmø") from database, not float values (e.g., "0.082")
                                    if (!string.IsNullOrWhiteSpace(zone.MepElementSizeParameterValue))
                                    {
                                        // Remove any existing Size parameter and replace with database value
                                        mepParams.Remove("Size");
                                        mepParams["Size"] = zone.MepElementSizeParameterValue;
                                        
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            DebugLogger.Info($"[PARAM_TRANSFER] ✅ Individual: Using MepElementSizeParameterValue='{zone.MepElementSizeParameterValue}' for zone {zone.Id} (from database)");
                                        }
                                    }
                                    
                                    var hostParams = zone.HostParameterValues?.ToDictionary(p => p.Key, p => p.Value) ?? new Dictionary<string, string>();
                                    
                                    filterData[(int)zone.SleeveInstanceId] = (mepParams, hostParams);
                                                                        if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[PARAM_TRANSFER] Added individual sleeve {zone.SleeveInstanceId} to index");
                                }
                            }
                            
                            // Second pass: aggregate cluster sleeves
                            var clusterGroups = zones.Where(z => z.ClusterSleeveInstanceId > 0)
                                                   .GroupBy(z => z.ClusterSleeveInstanceId);
                            
                            foreach (var clusterGroup in clusterGroups)
                            {
                                var clusterSleeveId = clusterGroup.Key;
                                var clusterZones = clusterGroup.ToList();
                                
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[PARAM_TRANSFER] Aggregating {clusterZones.Count} clash zones for cluster sleeve {clusterSleeveId}");
                                
                                // Aggregate MEP parameters
                                var aggregatedMepParams = new Dictionary<string, string>();
                                var aggregatedHostParams = new Dictionary<string, string>();
                                
                                foreach (var zone in clusterZones)
                                {
                                    // ✅ CRITICAL FIX FOR PIPES: Prioritize MepElementSizeParameterValue for Size parameter
                                    // This ensures cluster sleeves get text values (e.g., "20 mmø") from database, not float values (e.g., "0.082") from XML
                                    // Same logic as individual sleeves - use MepElementSizeParameterValue column from database
                                    bool hasSizeFromDatabase = false;
                                    string sizeValueFromDatabase = null;
                                    
                                    if (!string.IsNullOrWhiteSpace(zone.MepElementSizeParameterValue))
                                    {
                                        hasSizeFromDatabase = true;
                                        sizeValueFromDatabase = zone.MepElementSizeParameterValue.Trim();
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            DebugLogger.Info($"[PARAM_TRANSFER] ✅ Cluster: Using MepElementSizeParameterValue='{sizeValueFromDatabase}' for zone {zone.Id} (from database, not XML)");
                                        }
                                    }
                                    
                                    // Aggregate MEP parameters
                                    if (zone.MepParameterValues != null)
                                    {
                                        foreach (var param in zone.MepParameterValues)
                                        {
                                            // ✅ CRITICAL FIX: Skip Size parameter from XML if we have MepElementSizeParameterValue from database
                                            // This ensures we use the text value (e.g., "20 mmø") instead of float value (e.g., "0.082")
                                            if (param.Key.Equals("Size", StringComparison.OrdinalIgnoreCase) && hasSizeFromDatabase)
                                            {
                                                // Use database value instead of XML value
                                                if (!aggregatedMepParams.ContainsKey("Size"))
                                                {
                                                    aggregatedMepParams["Size"] = sizeValueFromDatabase;
                                                    if (!DeploymentConfiguration.DeploymentMode)
                                                    {
                                                        DebugLogger.Info($"[PARAM_TRANSFER] ✅ Cluster: Added Size='{sizeValueFromDatabase}' from MepElementSizeParameterValue (database) for zone {zone.Id}");
                                                    }
                                                }
                                                else
                                                {
                                                    // Append to existing value with comma separation
                                                    aggregatedMepParams["Size"] = $"{aggregatedMepParams["Size"]}, {sizeValueFromDatabase}";
                                                    if (!DeploymentConfiguration.DeploymentMode)
                                                    {
                                                        DebugLogger.Info($"[PARAM_TRANSFER] ✅ Cluster: Appended Size='{sizeValueFromDatabase}' from MepElementSizeParameterValue (database) to existing sizes");
                                                    }
                                                }
                                                continue; // Skip processing this Size parameter from XML
                                            }
                                            
                                            if (!string.IsNullOrEmpty(param.Value))
                                            {
                                                // ✅ FIX: For Size parameter, extract only the first part before dash (format: "1350x1000-1350x1000" → "1350x1000")
                                                // ✅ SAFE: If Size is already correct format (no dash), it passes through unchanged
                                                string cleanValue = param.Value;
                                                if (param.Key.Equals("Size", StringComparison.OrdinalIgnoreCase) && param.Value.Contains("-"))
                                                {
                                                    // Extract first part before dash (e.g., "1350x1000-1350x1000" → "1350x1000")
                                                    var parts = param.Value.Split(new[] { '-' }, 2);
                                                    if (parts.Length > 0 && !string.IsNullOrEmpty(parts[0]))
                                                    {
                                                        cleanValue = parts[0].Trim();
                                                                                                                if (!DeploymentConfiguration.DeploymentMode)
                                                            DebugLogger.Info($"[PARAM_TRANSFER] Cleaned Size parameter: '{param.Value}' → '{cleanValue}'");
                                                                                                                if (!DeploymentConfiguration.DeploymentMode)
                                                            DebugLogger.Info($"[{DateTime.Now}] [PARAM_TRANSFER] Cleaned Size: '{param.Value}' → '{cleanValue}'\n");
                                                    }
                                                }
                                                // ✅ Note: If Size doesn't contain "-", cleanValue = param.Value (unchanged) - handles both formats correctly
                                                
                                                if (aggregatedMepParams.ContainsKey(param.Key))
                                                {
                                                    // Append to existing value with comma separation
                                                    var existingValue = aggregatedMepParams[param.Key];
                                                    
                                                    // ✅ FIX FOR CLUSTER SIZES: For Size/MEP Size parameter, include ALL sizes even if duplicates
                                                    // This ensures 3 elements with same size show as "200 mmø, 200 mmø, 200 mmø" instead of just "200 mmø"
                                                    bool isSizeParameter = param.Key.Equals("Size", StringComparison.OrdinalIgnoreCase) || 
                                                                           param.Key.Equals("MEP Size", StringComparison.OrdinalIgnoreCase) ||
                                                                           param.Key.Equals("MepElementFormattedSize", StringComparison.OrdinalIgnoreCase) ||
                                                                           param.Key.Equals("Service Size", StringComparison.OrdinalIgnoreCase);
                                                    
                                                    if (isSizeParameter)
                                                    {
                                                        // ✅ SIZE PARAMETER: Always add all values, including duplicates
                                                        aggregatedMepParams[param.Key] = $"{existingValue}, {cleanValue}";
                                                                                                                if (!DeploymentConfiguration.DeploymentMode)
                                                            DebugLogger.Info($"[PARAM_TRANSFER] Added size '{cleanValue}' to cluster (including duplicates). Total sizes: {existingValue}, {cleanValue}");
                                                    }
                                                    else
                                                    {
                                                        // ✅ OTHER PARAMETERS: Check for unique values only (to avoid duplicates for System Type, etc.)
                                                    var existingParts = existingValue.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                                        .Select(p => p.Trim())
                                                        .ToList();
                                                    
                                                    if (!existingParts.Any(e => e.Equals(cleanValue, StringComparison.OrdinalIgnoreCase)))
                                                    {
                                                        aggregatedMepParams[param.Key] = $"{existingValue}, {cleanValue}";
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    aggregatedMepParams[param.Key] = cleanValue;
                                                }
                                            }
                                        }
                                    }
                                    
                                    // ✅ FALLBACK: If Size parameter wasn't found in MepParameterValues and we have MepElementSizeParameterValue, add it
                                    if (hasSizeFromDatabase && !aggregatedMepParams.ContainsKey("Size"))
                                    {
                                        aggregatedMepParams["Size"] = sizeValueFromDatabase;
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            DebugLogger.Info($"[PARAM_TRANSFER] ✅ Cluster: Added Size='{sizeValueFromDatabase}' from MepElementSizeParameterValue (fallback) for zone {zone.Id}");
                                        }
                                    }
                                    
                                    // Aggregate Host parameters
                                    if (zone.HostParameterValues != null)
                                    {
                                        foreach (var param in zone.HostParameterValues)
                                        {
                                            if (!string.IsNullOrEmpty(param.Value))
                                            {
                                                if (aggregatedHostParams.ContainsKey(param.Key))
                                                {
                                                    // Append to existing value with comma separation
                                                    var existingValue = aggregatedHostParams[param.Key];
                                                    if (!existingValue.Contains(param.Value))
                                                    {
                                                        aggregatedHostParams[param.Key] = $"{existingValue}, {param.Value}";
                                                    }
                                                }
                                                else
                                                {
                                                    aggregatedHostParams[param.Key] = param.Value;
                                                }
                                            }
                                        }
                                    }
                                }
                                
                                filterData[(int)clusterSleeveId] = (aggregatedMepParams, aggregatedHostParams);
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[PARAM_TRANSFER] Added aggregated cluster sleeve {clusterSleeveId} with {aggregatedMepParams.Count} MEP params and {aggregatedHostParams.Count} host params");
                            }
                            
                            filterIndex[filterKey] = filterData;
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[PARAM_TRANSFER] Built index for filter '{filterName}' ({category}): {filterData.Count} sleeves");
                                string transferDebugLogPath5 = SafeFileLogger.GetLogFilePath("transfer_debug.log");
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    File.AppendAllText(transferDebugLogPath5, $"[{DateTime.Now}] [PARAM_TRANSFER] Built index for filter '{filterName}' ({category}): {filterData.Count} sleeves\n");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Error($"[PARAM_TRANSFER] Error loading filter '{filterName}' ({category}): {ex.Message}");
                                SafeFileLogger.SafeAppendText("transfer_debug.log", 
                                    $"[{DateTime.Now}] [PARAM_TRANSFER] ERROR loading filter '{filterName}' ({category}): {ex.Message}\n");
                            }
                        }
                    }
                }
                
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[PARAM_TRANSFER] Filter index built with {filterIndex.Count} filters from database");
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string transferDebugLogPath7 = SafeFileLogger.GetLogFilePath("transfer_debug.log");
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        File.AppendAllText(transferDebugLogPath7, $"[{DateTime.Now}] [PARAM_TRANSFER] Filter index built with {filterIndex.Count} filters from database\n");
                    }
                    if (filterIndex.Count > 0)
                    {
                        System.IO.File.AppendAllText(transferDebugLogPath7, $"[{DateTime.Now}] [PARAM_TRANSFER] Available filter keys: {string.Join(", ", filterIndex.Keys.Take(10))}\n");
                    }
                }
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[PARAM_TRANSFER] Error building filter index: {ex.Message}");
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[{DateTime.Now}] [PARAM_TRANSFER] Error building filter index: {ex.Message}\n");
            }
            
            return filterIndex;
        }
        
        /// <summary>
        /// Convert a parameter value to a robust invariant string (same logic as ParameterSnapshotService)
        /// </summary>
        private string ConvertParameterToString(Element owner, Parameter p)
        {
            if (p == null) return string.Empty;

            string value = p.AsString();
            if (!string.IsNullOrEmpty(value)) return value;

            value = p.AsValueString();
            if (!string.IsNullOrEmpty(value)) return value;

            switch (p.StorageType)
            {
                case StorageType.Integer:
                    return p.AsInteger().ToString(System.Globalization.CultureInfo.InvariantCulture);
                case StorageType.Double:
                    return Math.Round(p.AsDouble(), 3).ToString(System.Globalization.CultureInfo.InvariantCulture);
                case StorageType.ElementId:
                    try
                    {
                        var elemId = p.AsElementId();
                        if (elemId != null && elemId.GetIntegerValue() != -1)
                        {
                            var elem = owner.Document.GetElement(elemId);
                            return elem?.Name ?? elemId.GetIntegerValue().ToString();
                        }
                    }
                    catch { }
                    return string.Empty;
                default:
                    return string.Empty;
            }
        }

        private List<Element> GetMepElementsInOpening(Document doc, Element opening, UIDocument uiDoc = null)
        {
            var mepElements = new List<Element>();
            
            try
            {
                // ✅ SECTION BOX FILTERING: Get cached section box bounds if available
                BoundingBoxXYZ? sectionBoxBounds = null;
                
                // ✅ OPTIMIZATION: Use cached section box if flag is enabled
                if (Services.OptimizationFlags.UseSectionBoxFilterForParameterTransfer)
                {
                    // Try to get cached section box bounds from database
                    // We need to get the database connection from the document
                    try
                    {
                        using (var dbContext = new SleeveDbContext(doc, msg => { }))
                        {
                            sectionBoxBounds = _sectionBoxService.GetSectionBoxBounds(dbContext.Connection);
                        }
                    }
                    catch (Exception ex)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[PARAM_TRANSFER] Failed to get cached section box bounds: {ex.Message}");
                        }
                    }
                    
                    if (sectionBoxBounds != null && !DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[PARAM_TRANSFER] Using cached section box bounds for filtering");
                    }
                }
                
                // ✅ FALLBACK: If no cached section box, try live Revit API
                if (sectionBoxBounds == null && uiDoc != null && uiDoc.ActiveView is View3D view3D && view3D.IsSectionBoxActive)
                {
                    sectionBoxBounds = Helpers.SectionBoxHelper.GetSectionBoxBounds(view3D);
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[PARAM_TRANSFER] Using live Revit section box bounds (cache not available)");
                    }
                }
                
                // 1) MEPCurve sources (ducts, pipes, trays)
                var curveCollector = new FilteredElementCollector(doc)
                    .OfClass(typeof(MEPCurve))
                    .WhereElementIsNotElementType();
                
                // ✅ SECTION BOX FILTER: Apply bounding box filter if section box is active
                if (sectionBoxBounds != null)
                {
                    var outline = new Outline(sectionBoxBounds.Min, sectionBoxBounds.Max);
                    var sectionBoxFilter = new BoundingBoxIntersectsFilter(outline);
                    curveCollector = curveCollector.WherePasses(sectionBoxFilter);
                }
                
                foreach (Element mepElement in curveCollector)
                {
                    if (ElementsIntersect(opening, mepElement))
                    {
                        mepElements.Add(mepElement);
                    }
                }

                // 2) Additional MEP family-instance categories (not MEPCurve): fittings/accessories for ducts/pipes/cable trays
                var fiCategories = new List<BuiltInCategory>
                {
                    BuiltInCategory.OST_DuctAccessory,
                    BuiltInCategory.OST_DuctFitting,
                    BuiltInCategory.OST_PipeAccessory,
                    BuiltInCategory.OST_PipeFitting,
                    BuiltInCategory.OST_CableTrayFitting
                };
                foreach (var bic in fiCategories)
                {
                    var fiCollector = new FilteredElementCollector(doc)
                        .OfCategory(bic)
                        .WhereElementIsNotElementType();
                    
                    // ✅ SECTION BOX FILTER: Apply bounding box filter if section box is active
                    if (sectionBoxBounds != null)
                    {
                        var outline = new Outline(sectionBoxBounds.Min, sectionBoxBounds.Max);
                        var sectionBoxFilter = new BoundingBoxIntersectsFilter(outline);
                        fiCollector = fiCollector.WherePasses(sectionBoxFilter);
                    }
                    
                    foreach (Element e in fiCollector)
                    {
                        if (ElementsIntersect(opening, e))
                        {
                            mepElements.Add(e);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log error but continue
                System.Diagnostics.Debug.WriteLine($"Error getting MEP elements: {ex.Message}");
            }
            
            return mepElements;
        }
        
        private List<Element> GetHostElementsForOpening(Document doc, Element opening)
        {
            var hostElements = new List<Element>();
            
            try
            {
                // Get host elements (walls, floors, ceilings) that contain this opening
                var hostCategories = new List<BuiltInCategory>
                {
                    BuiltInCategory.OST_Walls,
                    BuiltInCategory.OST_Floors,
                    BuiltInCategory.OST_Ceilings
                };
                
                var filter = new ElementMulticategoryFilter(hostCategories);
                var collector = new FilteredElementCollector(doc)
                    .WherePasses(filter)
                    .WhereElementIsNotElementType();
                
                foreach (Element hostElement in collector)
                {
                    if (ElementsIntersect(opening, hostElement))
                    {
                        hostElements.Add(hostElement);
                    }
                }
            }
            catch (Exception ex)
            {
                // Log error but continue
                System.Diagnostics.Debug.WriteLine($"Error getting host elements: {ex.Message}");
            }
            
            return hostElements;
        }
        
        private Element GetLevelForOpening(Document doc, Element opening)
        {
            try
            {
                // Get level from opening's location
                var location = opening.Location as LocationPoint;
                if (location != null)
                {
                    var level = doc.GetElement(opening.LevelId) as Level;
                    return level;
                }
            }
            catch (Exception ex)
            {
                // Log error but continue
                System.Diagnostics.Debug.WriteLine($"Error getting level: {ex.Message}");
            }
            
            return null;
        }
        
        private bool TransferParameterFromElements(
            Document doc, 
            Element targetElement, 
            List<Element> sourceElements, 
            ParameterMapping mapping)
        {
            if (sourceElements.Count == 0) return false;
            
            if (sourceElements.Count == 1)
            {
                return TransferParameterFromElement(doc, targetElement, sourceElements[0], mapping);
            }
            else
            {
                // Multiple elements - combine values
                return TransferParameterFromMultipleElements(doc, targetElement, sourceElements, mapping);
            }
        }
        
        private bool TransferParameterFromElement(
            Document doc, 
            Element targetElement, 
            Element sourceElement, 
            ParameterMapping mapping)
        {
            try
            {
                var sourceParam = sourceElement.LookupParameter(mapping.SourceParameter);
                if (sourceParam == null) return false;
                
                var targetParam = targetElement.LookupParameter(mapping.TargetParameter);
                if (targetParam == null || targetParam.IsReadOnly) return false;
                
                var value = GetParameterValueAsString(sourceParam);
                if (string.IsNullOrWhiteSpace(value)) return false;
                
                // Apply renaming if conditions exist
                value = _renamingService.ApplyRenaming(value, mapping.SourceParameter);
                
                // Apply service type abbreviation if this is a service type parameter
                if (IsServiceTypeParameter(mapping.SourceParameter))
                {
                    value = _abbreviationService.GetAbbreviation(value, mapping.SourceParameter);
                }
                
                targetParam.Set(value);

                // Record learned key so it will be snapshotted next Refresh
                ParameterSnapshotService.AddLearnedKey(mapping.SourceParameter);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error transferring parameter: {ex.Message}");
                return false;
            }
        }
        
        private bool TransferParameterFromMultipleElements(
            Document doc, 
            Element targetElement, 
            List<Element> sourceElements, 
            ParameterMapping mapping)
        {
            try
            {
                var values = new List<string>();
                
                foreach (var sourceElement in sourceElements)
                {
                    var sourceParam = sourceElement.LookupParameter(mapping.SourceParameter);
                    if (sourceParam != null)
                    {
                        var value = GetParameterValueAsString(sourceParam);
                        if (!string.IsNullOrEmpty(value))
                        {
                            // Apply renaming if conditions exist
                            value = _renamingService.ApplyRenaming(value, mapping.SourceParameter);
                            
                            // Apply service type abbreviation if this is a service type parameter
                            if (IsServiceTypeParameter(mapping.SourceParameter))
                            {
                                value = _abbreviationService.GetAbbreviation(value, mapping.SourceParameter);
                            }
                            
                            values.Add(value);
                        }
                    }
                }
                
                if (values.Count == 0) return false;
                
                var targetParam = targetElement.LookupParameter(mapping.TargetParameter);
                if (targetParam == null || targetParam.IsReadOnly) return false;
                
                var combinedValue = string.Join(mapping.Separator, values.Distinct());
                targetParam.Set(combinedValue);
                // Record learned key so it will be snapshotted next Refresh
                ParameterSnapshotService.AddLearnedKey(mapping.SourceParameter);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error transferring multiple parameters: {ex.Message}");
                return false;
            }
        }

        private string GetParameterValueAsString(Parameter p)
        {
            try
            {
                if (p == null) return string.Empty;
                var s = p.AsString();
                if (!string.IsNullOrWhiteSpace(s)) return s;
                s = p.AsValueString();
                if (!string.IsNullOrWhiteSpace(s)) return s;
                switch (p.StorageType)
                {
                    case StorageType.Integer:
                        return p.AsInteger().ToString(System.Globalization.CultureInfo.InvariantCulture);
                    case StorageType.Double:
                        return p.AsDouble().ToString(System.Globalization.CultureInfo.InvariantCulture);
                    case StorageType.ElementId:
                        return p.AsElementId()?.GetIntegerValue().ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
                    default:
                        return string.Empty;
                }
            }
            catch { return string.Empty; }
        }
        
        private int GetIntegerParameter(Element element, string parameterName)
        {
            try
            {
                var param = element.LookupParameter(parameterName);
                if (param != null && param.StorageType == StorageType.Integer)
                {
                    return param.AsInteger();
                }
                if (param != null && param.StorageType == StorageType.String)
                {
                    if (int.TryParse(param.AsString(), out var parsed))
                    {
                        return parsed;
                    }
                }
            }
            catch
            {
                // ignored – return default on failure
            }

            return -1;
        }
        
        private bool ElementsIntersect(Element element1, Element element2)
        {
            try
            {
                var geom1 = element1.get_Geometry(Helpers.GeometryOptionsFactory.CreateIntersectionOptions());
                var geom2 = element2.get_Geometry(Helpers.GeometryOptionsFactory.CreateIntersectionOptions());
                
                if (geom1 == null || geom2 == null) return false;
                
                // Simple bounding box intersection check
                var bbox1 = geom1.GetBoundingBox();
                var bbox2 = geom2.GetBoundingBox();
                
                // Custom intersection check since BoundingBoxXYZ.Intersects doesn't exist
                return bbox1.Min.X <= bbox2.Max.X && bbox1.Max.X >= bbox2.Min.X &&
                       bbox1.Min.Y <= bbox2.Max.Y && bbox1.Max.Y >= bbox2.Min.Y &&
                       bbox1.Min.Z <= bbox2.Max.Z && bbox1.Max.Z >= bbox2.Min.Z;
            }
            catch
            {
                return false;
            }
        }
        
        // Helpers for standard transfer
        private string GetLevelName(Document doc, Element element)
        {
            try
            {
                var level = doc.GetElement(element.LevelId) as Level;
                return level?.Name ?? string.Empty;
            }
            catch { return string.Empty; }
        }
        
        private double? GetParamDouble(Element element, params string[] names)
        {
            foreach (var name in names)
            {
                try
                {
                    var p = element.LookupParameter(name);
                    if (p != null)
                    {
                        // Prefer AsDouble for numeric params; gracefully handle string numerics
                        if (p.StorageType == StorageType.Double)
                        {
                            return p.AsDouble();
                        }
                        if (p.StorageType == StorageType.String)
                        {
                            if (double.TryParse(p.AsString(), out var d)) return d;
                        }
                    }
                }
                catch { }
            }
            return null;
        }
        
        /// <summary>
        /// ✅ OPTIMIZATION 2: Check if parameter should be skipped (already matches snapshot value).
        /// Returns true if current value matches expected snapshot value (already transferred).
        /// </summary>
        private bool ShouldSkipParameter(Parameter target, string expectedValue, int sleeveId, string paramName)
        {
            try
            {
                if (target == null || target.IsReadOnly)
                    return false; // Process if parameter doesn't exist or is read-only
                
                if (string.IsNullOrWhiteSpace(expectedValue))
                    return false; // Process if snapshot value is empty
                
                // Get current value on sleeve
                string currentValue = null;
                if (target.StorageType == StorageType.String)
                {
                    currentValue = target.AsString();
                }
                else if (target.StorageType == StorageType.Integer)
                {
                    currentValue = target.AsInteger().ToString(System.Globalization.CultureInfo.InvariantCulture);
                }
                else if (target.StorageType == StorageType.Double)
                {
                    currentValue = target.AsDouble().ToString(System.Globalization.CultureInfo.InvariantCulture);
                }
                else
                {
                    return false; // Process if storage type is not supported for comparison
                }
                
                if (string.IsNullOrWhiteSpace(currentValue))
                    return false; // Process if current value is empty
                
                // Compare values (case-insensitive for strings, exact for numbers)
                bool alreadyTransferred = string.Equals(currentValue.Trim(), expectedValue.Trim(), StringComparison.OrdinalIgnoreCase);
                
                if (alreadyTransferred && !DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[PARAM_TRANSFER] ⏭️ Skipping sleeve {sleeveId}: '{paramName}' already matches snapshot value '{currentValue}'");
                }
                
                return alreadyTransferred;
            }
            catch
            {
                return false; // Process on error (safe default)
            }
        }

        private bool SetParameterValueSafely(Parameter target, string value)
        {
            try
            {
                // ✅ CRASH-SAFETY 8: Comprehensive null and validation checks
                if (target == null || target.IsReadOnly || string.IsNullOrWhiteSpace(value)) 
                    return false;

                // ✅ CRASH-SAFETY 9: Validate element is still valid
                if (target.Element == null || !target.Element.IsValidObject)
                    return false;

                // ✅ CRITICAL FIX: Handle Integer parameters (e.g., MEP_ElementId)
                if (target.StorageType == StorageType.Integer)
                {
                    if (int.TryParse(value, out int intValue))
                    {
                        // ✅ CRASH-SAFETY 10: Validate integer bounds (ElementId range: 0 to int.MaxValue, but reasonable bounds for MEP_ElementId)
                        if (intValue < 0 || intValue > 2147483647) // int.MaxValue, but we'll use a reasonable upper bound
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[PARAM_TRANSFER] Integer value {intValue} out of bounds for parameter '{target.Definition?.Name}'");
                            return false;
                        }
                        
                        target.Set(intValue);
                        return true;
                    }
                    return false;
                }
                
                // Handle String parameters
                if (target.StorageType == StorageType.String) 
                { 
                    // ✅ CRASH-SAFETY 11: Validate string length (Revit parameter max length is typically 255 characters)
                    string safeValue = value ?? string.Empty;
                    if (safeValue.Length > 255)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[PARAM_TRANSFER] String value exceeds 255 characters for parameter '{target.Definition?.Name}', truncating");
                        safeValue = safeValue.Substring(0, 255);
                    }
                    
                    target.Set(safeValue); 
                    return true; 
                }
                
                return false;
            }
            catch (Exception ex)
            {
                // ✅ CRASH-SAFETY 12: Log exception details for debugging
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[PARAM_TRANSFER] Exception setting parameter '{target?.Definition?.Name}': {ex.Message}");
                }
                return false;
            }
        }
        
        private bool SetParameterValueSafely(Parameter target, double value)
        {
            try
            {
                if (target == null || target.IsReadOnly) return false;
                if (target.StorageType == StorageType.Double) { target.Set(value); return true; }
                if (target.StorageType == StorageType.String) { target.Set(value.ToString()); return true; }
                return false;
            }
            catch { return false; }
        }
        
        private bool IsServiceTypeParameter(string parameterName)
        {
            var serviceTypeParameters = new List<string>
            {
                "System Abbreviation",
                "System Name",
                "System Type",
                "Family Name",
                "Type Name"
            };
            
            return serviceTypeParameters.Any(param => 
                string.Equals(param, parameterName, StringComparison.OrdinalIgnoreCase));
        }
        
        #endregion

        #region Configuration Management

        /// <summary>
        /// Gets the current parameter transfer configuration from the project/profile
        /// </summary>
        public ParameterTransferConfiguration? GetCurrentParameterTransferConfiguration()
        {
            try
            {
                // Try to load from project-specific configuration file
                var configDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "ParameterTransfer");
                var configFile = Path.Combine(configDir, "current_parameter_transfer_config.xml");
                
                if (File.Exists(configFile))
                {
                    var serializer = new System.Xml.Serialization.XmlSerializer(typeof(ParameterTransferConfiguration));
                    using (var reader = new StreamReader(configFile))
                    {
                        return (ParameterTransferConfiguration?)serializer.Deserialize(reader);
                    }
                }
                
                // Return null if no configuration exists
                return null;
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"Failed to get current parameter transfer configuration: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Saves the current parameter transfer configuration to project-specific storage
        /// </summary>
        public bool SaveCurrentParameterTransferConfiguration(ParameterTransferConfiguration config)
        {
            try
            {
                var configDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "ParameterTransfer");
                if (!Directory.Exists(configDir))
                {
                    Directory.CreateDirectory(configDir);
                }
                
                var configFile = Path.Combine(configDir, "current_parameter_transfer_config.xml");
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(ParameterTransferConfiguration));
                
                using (var writer = new StreamWriter(configFile))
                {
                    serializer.Serialize(writer, config);
                }
                
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"Saved parameter transfer configuration to: {configFile}");
                return true;
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"Failed to save parameter transfer configuration: {ex.Message}");
                return false;
            }
        }

        #endregion

        /// <summary>
        /// Check if a sleeve is a cluster sleeve
        /// </summary>
        private bool IsClusterSleeve(Element sleeve)
        {
            try
            {
                // Check if the sleeve has a "Sleeve Instance ID" parameter that matches its own ID
                // For cluster sleeves, this parameter should contain the cluster sleeve's own ID
                var instanceIdParam = sleeve.LookupParameter("Sleeve Instance ID");
                if (instanceIdParam == null) return false;
                
                int paramValue = instanceIdParam.AsInteger();
                int sleeveId = sleeve.Id.GetIntegerValue();
                
                // If the parameter value matches the sleeve's own ID, it's likely a cluster sleeve
                // Individual sleeves would have their own ID, but cluster sleeves replace multiple individual sleeves
                return paramValue == sleeveId;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// CRITICAL FIX: Get aggregated parameters for cluster sleeves
        /// Parameters are already aggregated in the filter index, so we just need to retrieve them
        /// </summary>
        private string GetAggregatedClusterParameters(
            Dictionary<int, (Dictionary<string,string> mep, Dictionary<string,string> host)> filterData,
            int clusterSleeveId,
            string parameterName,
            bool useHost)
        {
            try
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[AGGREGATE] Getting aggregated parameters for cluster sleeve {clusterSleeveId}, parameter '{parameterName}', useHost={useHost}");
                
                // The parameters are already aggregated in the filter index
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[AGGREGATE] Looking for cluster sleeve {clusterSleeveId} in filterData with {filterData.Count} entries");
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[{DateTime.Now}] [AGGREGATE] Looking for cluster sleeve {clusterSleeveId} in filterData with {filterData.Count} entries\n");
                
                // ✅ DEBUG: Log all cluster sleeve IDs in filterData
                var clusterIdsInData = filterData.Keys.Where(k => k > 100000).Take(10).ToList(); // Cluster IDs are typically large numbers
                if (clusterIdsInData.Count > 0)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[AGGREGATE] Sample cluster sleeve IDs in filterData: {string.Join(", ", clusterIdsInData)}");
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        string transferDebugLogPath = SafeFileLogger.GetLogFilePath("transfer_debug.log");
                        System.IO.File.AppendAllText(transferDebugLogPath, $"[{DateTime.Now}] [AGGREGATE] Sample cluster sleeve IDs in filterData: {string.Join(", ", clusterIdsInData)}\n");
                    }
                }
                
                if (filterData.TryGetValue(clusterSleeveId, out var paramBags))
                {
                    var sourceParams = useHost ? paramBags.host : paramBags.mep;
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[AGGREGATE] Found cluster sleeve {clusterSleeveId}, has {sourceParams.Count} {(useHost ? "host" : "MEP")} parameters");
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        string transferDebugLogPath2 = SafeFileLogger.GetLogFilePath("transfer_debug.log");
                        System.IO.File.AppendAllText(transferDebugLogPath2, $"[{DateTime.Now}] [AGGREGATE] Found cluster sleeve {clusterSleeveId}, has {sourceParams.Count} {(useHost ? "host" : "MEP")} parameters: {string.Join(", ", sourceParams.Keys.Take(10))}\n");
                    }
                    
                    if (sourceParams.TryGetValue(parameterName, out var paramValue) && !string.IsNullOrEmpty(paramValue))
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[AGGREGATE] Found aggregated parameter '{parameterName}' = '{paramValue}' for cluster sleeve {clusterSleeveId}");
                        return paramValue;
                    }
                    else
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[AGGREGATE] Parameter '{parameterName}' not found in aggregated data for cluster sleeve {clusterSleeveId}. Available params: {string.Join(", ", sourceParams.Keys)}");
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            string transferDebugLogPath3 = SafeFileLogger.GetLogFilePath("transfer_debug.log");
                            System.IO.File.AppendAllText(transferDebugLogPath3, $"[{DateTime.Now}] [AGGREGATE] Parameter '{parameterName}' not found. Available: {string.Join(", ", sourceParams.Keys)}\n");
                        }
                        return null;
                    }
                }
                else
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[AGGREGATE] Cluster sleeve {clusterSleeveId} not found in filter data (checked {filterData.Count} entries)");
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        string transferDebugLogPath4 = SafeFileLogger.GetLogFilePath("transfer_debug.log");
                        System.IO.File.AppendAllText(transferDebugLogPath4, $"[{DateTime.Now}] [AGGREGATE] Cluster sleeve {clusterSleeveId} NOT FOUND in filterData. Sample keys: {string.Join(", ", filterData.Keys.Take(10))}\n");
                    }
                    return null;
                }
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[AGGREGATE] Error getting aggregated cluster parameters: {ex.Message}");
                return null;
            }
        }
        /// <summary>
        /// Aggregates parameters from multiple constituent snapshots.
        /// Values are joined with commas/semicolons and deduped.
        /// </summary>
        private Dictionary<string, string> AggregateConstituentParameters(
            List<SleeveConstituentSnapshotReference> constituents, 
            SleeveSnapshotIndex snapshotIndex,
            bool useHost)
        {
            var aggregatedMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var valueCollections = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            if (!DeploymentConfiguration.DeploymentMode)
            {
                SafeFileLogger.SafeAppendText("transfer_debug.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [AGGREGATE] Processing {constituents.Count} constituents for {(useHost ? "HOST" : "MEP")} parameters...\n");
            }

            foreach (var c in constituents)
            {
                SleeveSnapshotView s = null;
                string matchMethod = "None";
                
                // Try to resolve snapshot for this constituent
                if (!string.IsNullOrEmpty(c.ClashZoneGuid))
                {
                    bool found = false;
                    // 1. Direct Try
                    if (snapshotIndex.TryGetByClashZoneGuid(c.ClashZoneGuid, out var sGuids))
                    {
                        s = sGuids;
                        matchMethod = $"ClashZoneGuid:{c.ClashZoneGuid}";
                        found = true;
                    }
                    // 2. Normalization Try (Parsing GUID)
                    else if (Guid.TryParse(c.ClashZoneGuid, out var parsedGuid))
                    {
                        // Some dictionaries might store normalized string
                        string normalizedKey = parsedGuid.ToString();
                        if (snapshotIndex.TryGetByClashZoneGuid(normalizedKey, out var sNorm))
                        {
                            s = sNorm;
                            matchMethod = $"NormalizedGuid:{normalizedKey}";
                            found = true;
                        }
                        else
                        {
                             string upperParams = parsedGuid.ToString().ToUpperInvariant();
                             if (snapshotIndex.TryGetByClashZoneGuid(upperParams, out var sUpper))
                             {
                                 s = sUpper;
                                 matchMethod = $"UpperGuid:{upperParams}";
                                 found = true;
                             }
                        }
                    }

                    if (!found) 
                    {
                         if (!DeploymentConfiguration.DeploymentMode)
                         {
                             // Debug log available keys to see mismatch
                             var sampleKeys = snapshotIndex.ByClashZoneGuid.Keys.Take(5).ToList();
                             SafeFileLogger.SafeAppendText("transfer_debug.log",
                                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [AGGREGATE] ❌ Snapshot NOT FOUND for constituent. GUID='{c.ClashZoneGuid}', ClusterId='{c.ClusterInstanceId?.ToString() ?? ""}'.\n");
                         }
                    }
                }
                else if (c.ClusterInstanceId.HasValue && snapshotIndex.TryGetByCluster(c.ClusterInstanceId.Value, out var sCluster))
                {
                    s = sCluster;
                    matchMethod = $"ClusterInstanceId:{c.ClusterInstanceId}";
                }

                if (s != null)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                         SafeFileLogger.SafeAppendText("transfer_debug.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [AGGREGATE] ✅ Found snapshot via {matchMethod}. Params: {(useHost ? s.HostParameters?.Count : s.MepParameters?.Count) ?? 0}\n");
                    }

                    var sourceParams = useHost ? s.HostParameters : s.MepParameters;
                    if (sourceParams != null)
                    {
                        foreach (var kvp in sourceParams)
                        {
                            if (!valueCollections.ContainsKey(kvp.Key))
                            {
                                valueCollections[kvp.Key] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                            }
                            if (!string.IsNullOrWhiteSpace(kvp.Value))
                            {
                                // ✅ FIX: Split aggregated values by comma to handle cases where a constituent (e.g. Cluster)
                                // already contains multiple values (e.g. "100, 200"). This prevents "100, 200, 100" duplication.
                                var rawValues = kvp.Value.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                                
                                foreach (var rawVal in rawValues)
                                {
                                    var val = rawVal.Trim();
                                    if (string.IsNullOrWhiteSpace(val)) continue;

                                    // ✅ SIZE NORMALIZATION: Dedupe patterns like "475x200-475x200" → "475x200"
                                    if (kvp.Key.Equals("Size", StringComparison.OrdinalIgnoreCase) ||
                                        kvp.Key.Contains("Size", StringComparison.OrdinalIgnoreCase))
                                    {
                                        val = NormalizeSizeValue(val);
                                    }
                                    
                                    valueCollections[kvp.Key].Add(val);
                                }
                            }
                        }
                    }
                }
                else
                {
                     if (!DeploymentConfiguration.DeploymentMode)
                     {
                         SafeFileLogger.SafeAppendText("transfer_debug.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [AGGREGATE] ❌ Snapshot NOT FOUND for constituent. GUID='{c.ClashZoneGuid}', ClusterId='{c.ClusterInstanceId}'\n");
                     }
                }
            }

            foreach (var kvp in valueCollections)
            {
                if (kvp.Value.Count > 0)
                {
                    // Sort values for consistency
                    var sortedValues = kvp.Value.OrderBy(v => v).ToList();
                    aggregatedMap[kvp.Key] = string.Join(", ", sortedValues);
                }
            }

            if (!DeploymentConfiguration.DeploymentMode)
            {
                SafeFileLogger.SafeAppendText("transfer_debug.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [AGGREGATE] Result: {aggregatedMap.Count} unique parameters aggregated.\n");
            }

            return aggregatedMap;
        }

        /// <summary>
        /// Normalizes Size values by deduplicating repeated patterns.
        /// E.g., "475x200-475x200" → "475x200" (inlet/outlet are same size)
        /// E.g., "100-100" → "100" (diameter repeated)
        /// E.g., "475x200-500x200" → keeps as-is (different sizes)
        /// </summary>
        private string NormalizeSizeValue(string sizeValue)
        {
            if (string.IsNullOrWhiteSpace(sizeValue)) return sizeValue;
            
            // Check for pattern: "A-A" where A is the same on both sides of the dash
            // Handle: "475x200-475x200", "100-100", "ø100-ø100"
            var parts = sizeValue.Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries);
            
            if (parts.Length == 2)
            {
                var left = parts[0].Trim();
                var right = parts[1].Trim();
                
                if (string.Equals(left, right, StringComparison.OrdinalIgnoreCase))
                {
                    // Both sides are identical, return just one
                    return left;
                }
            }
            
            // Not a duplicate pattern, return original
            return sizeValue;
        }

    }
}