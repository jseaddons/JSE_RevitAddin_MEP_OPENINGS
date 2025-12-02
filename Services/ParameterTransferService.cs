using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for transferring parameters from various sources to openings
    /// </summary>
    public class ParameterTransferService
    {
        private readonly ParameterRenamingService _renamingService;
        private readonly ParameterMappingService _mappingService;
        private readonly ServiceTypeAbbreviationService _abbreviationService;
        private readonly MepElementAnalysisService _mepAnalysisService;
        
        public ParameterTransferService()
        {
            _renamingService = new ParameterRenamingService();
            _mappingService = new ParameterMappingService();
            _abbreviationService = new ServiceTypeAbbreviationService();
            _mepAnalysisService = new MepElementAnalysisService();
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
                        var serviceSizeCalculation = _mepAnalysisService.CalculateServiceSize(mepElements, clearance);

                        // Set parameter value
                        var param = opening.LookupParameter(targetParameter);
                        if (param != null && !param.IsReadOnly)
                        {
                            param.Set(serviceSizeCalculation.CalculationString);
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
            int transferred = 0;
            int failed = 0;
            
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
                        var categoryId = source.Category?.Id.IntegerValue ?? -1;

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
            
            // Backward-compatible wrapper which creates a transaction and calls the in-transaction implementation
            using (var t = new Transaction(doc, "Execute Parameter Transfer Configuration (wrapper)"))
            {
                t.Start();
                var r = ExecuteTransferConfigurationInTransaction(doc, openingIds, config, null);
                t.Commit();
                
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[PARAM_TRANSFER] ExecuteTransferConfiguration completed: Success={r.Success}, TransferredCount={r.TransferredCount}, FailedCount={r.FailedCount}");
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[{DateTime.Now}] [PARAM_TRANSFER] ExecuteTransferConfiguration completed: Success={r.Success}, TransferredCount={r.TransferredCount}, FailedCount={r.FailedCount}\n");
                
                return r;
            }
        }

        /// <summary>
        /// Execute the whole transfer configuration assuming the caller owns the transaction.
        /// This method performs all transfers by calling the InTransaction variants.
        /// </summary>
        public ParameterTransferResult ExecuteTransferConfigurationInTransaction(
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
                
                // Execute each mapping
                foreach (var mapping in config.Mappings)
                {
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
                    _mepAnalysisService.SetClearanceParameters(config.DefaultClearance, config.ClearanceSuffix);

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
            var result = new ParameterTransferResult();
            var transferredCount = 0;
            var failedCount = 0;
            var errors = new List<string>();

            foreach (var openingId in openingIds)
            {
                try
                {
                    var opening = doc.GetElement(openingId);
                    if (opening == null)
                    {
                        failedCount++;
                        errors.Add($"Opening element {openingId.IntegerValue} not found.");
                        continue;
                    }

                    var sleeveInstanceId = GetIntegerParameter(opening, "Sleeve Instance ID");
                    var clusterInstanceId = GetIntegerParameter(opening, "Cluster Sleeve Instance ID");

                    // ✅ DIAGNOSTIC: Log matching attempt
                    var transferDebugLogPath = SafeFileLogger.GetLogFilePath("transfer_debug.log");
                    SafeFileLogger.SafeAppendText("transfer_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] Matching sleeve {openingId.IntegerValue}: SleeveInstanceId={sleeveInstanceId}, ClusterInstanceId={clusterInstanceId}, SnapshotIndex.BySleeve.Count={snapshotIndex.BySleeve.Count}, SnapshotIndex.ByCluster.Count={snapshotIndex.ByCluster.Count}\n");

                    SleeveSnapshotView snapshot = null;
                    if (clusterInstanceId > 0 && snapshotIndex.TryGetByCluster(clusterInstanceId, out var clusterView))
                    {
                        snapshot = clusterView;
                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅ Matched by ClusterInstanceId={clusterInstanceId}\n");
                    }
                    else if (sleeveInstanceId > 0 && snapshotIndex.TryGetBySleeve(sleeveInstanceId, out var sleeveView))
                    {
                        snapshot = sleeveView;
                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ✅ Matched by SleeveInstanceId={sleeveInstanceId}\n");
                    }

                    if (snapshot == null)
                    {
                        // ✅ DIAGNOSTIC: Log available snapshot IDs for debugging
                        var availableSleeveIds = string.Join(", ", snapshotIndex.BySleeve.Keys.Take(10));
                        var availableClusterIds = string.Join(", ", snapshotIndex.ByCluster.Keys.Take(10));
                        SafeFileLogger.SafeAppendText("transfer_debug.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_TRANSFER] ❌ No snapshot found for sleeve {openingId.IntegerValue} (SleeveId={sleeveInstanceId}, ClusterId={clusterInstanceId}). Available SleeveIds: [{availableSleeveIds}], Available ClusterIds: [{availableClusterIds}]\n");
                        result.Warnings.Add($"No persisted snapshot found for sleeve {openingId.IntegerValue} (SleeveId={sleeveInstanceId}, ClusterId={clusterInstanceId}).");
                        continue;
                    }

                    var sourceParams = useHost ? snapshot.HostParameters : snapshot.MepParameters;
                    if (sourceParams == null || sourceParams.Count == 0)
                    {
                        result.Warnings.Add($"No {(useHost ? "host" : "MEP")} parameters captured for sleeve {openingId.IntegerValue}.");
                        continue;
                    }

                    if (!sourceParams.TryGetValue(mapping.SourceParameter, out var sourceValue) || string.IsNullOrWhiteSpace(sourceValue))
                    {
                        result.Warnings.Add($"Parameter '{mapping.SourceParameter}' not found in snapshot for sleeve {openingId.IntegerValue}.");
                        continue;
                    }

                    var targetParam = opening.LookupParameter(mapping.TargetParameter);
                    if (targetParam == null)
                    {
                        failedCount++;
                        errors.Add($"Target parameter '{mapping.TargetParameter}' not found on opening {openingId.IntegerValue}.");
                        continue;
                    }

                    if (SetParameterValueSafely(targetParam, sourceValue))
                    {
                        transferredCount++;
                        successfullyTransferredSleeveIds?.Add(openingId.IntegerValue);
                    }
                    else
                    {
                        failedCount++;
                        errors.Add($"Failed to set parameter '{mapping.TargetParameter}' on opening {openingId.IntegerValue}.");
                    }
                }
                catch (Exception ex)
                {
                    failedCount++;
                    errors.Add($"Error transferring parameters to opening {openingId.IntegerValue}: {ex.Message}");
                }
            }

            result.Success = errors.Count == 0;
            result.TransferredCount = transferredCount;
            result.FailedCount = failedCount;
            result.Errors = errors;
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
                                    
                                    filterData[zone.SleeveInstanceId] = (mepParams, hostParams);
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
                                
                                filterData[clusterSleeveId] = (aggregatedMepParams, aggregatedHostParams);
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
                        if (elemId != null && elemId.IntegerValue != -1)
                        {
                            var elem = owner.Document.GetElement(elemId);
                            return elem?.Name ?? elemId.IntegerValue.ToString();
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
                // ✅ SECTION BOX FILTERING: Get section box bounds if available
                BoundingBoxXYZ sectionBoxBounds = null;
                if (uiDoc != null && uiDoc.ActiveView is View3D view3D && view3D.IsSectionBoxActive)
                {
                    sectionBoxBounds = Helpers.SectionBoxHelper.GetSectionBoxBounds(view3D);
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
                        return p.AsElementId()?.IntegerValue.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
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
        
        private bool SetParameterValueSafely(Parameter target, string value)
        {
            try
            {
                if (target == null || target.IsReadOnly || string.IsNullOrWhiteSpace(value)) return false;
                
                // ✅ CRITICAL FIX: Handle Integer parameters (e.g., MEP_ElementId)
                if (target.StorageType == StorageType.Integer)
                {
                    if (int.TryParse(value, out int intValue))
                    {
                        target.Set(intValue);
                        return true;
                    }
                    return false;
                }
                
                // Handle String parameters
                if (target.StorageType == StorageType.String) 
                { 
                    target.Set(value ?? string.Empty); 
                    return true; 
                }
                
                return false;
            }
            catch { return false; }
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
                int sleeveId = sleeve.Id.IntegerValue;
                
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
    }
}

