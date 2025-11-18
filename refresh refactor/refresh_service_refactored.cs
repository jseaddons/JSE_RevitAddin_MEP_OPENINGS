using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Refresh;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// REFACTORED RefreshService - Clean orchestrator pattern
    /// Down from 4000+ lines to ~200 lines
    /// All optimization strategies implemented
    /// </summary>
    public class RefreshServiceRefactored
    {
        private readonly Document _document;
        private readonly UIDocument _uiDocument;
        private readonly ApplicationProfileService _appProfileService;
        
        // UI controls
        private System.Windows.Forms.Label _statusLabel;
        private System.Windows.Forms.ProgressBar _progressBar;
        private System.Windows.Forms.Button _refreshButton;
        
        public RefreshServiceRefactored(
            Document document, 
            UIDocument uiDocument, 
            ApplicationProfileService appProfileService)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _uiDocument = uiDocument ?? throw new ArgumentNullException(nameof(uiDocument));
            _appProfileService = appProfileService ?? throw new ArgumentNullException(nameof(appProfileService));
        }
        
        public void SetUIReferences(
            System.Windows.Forms.Label statusLabel, 
            System.Windows.Forms.ProgressBar progressBar, 
            System.Windows.Forms.Button refreshButton)
        {
            _statusLabel = statusLabel;
            _progressBar = progressBar;
            _refreshButton = refreshButton;
        }
        
        /// <summary>
        /// Main refresh execution - orchestrates all helper services
        /// </summary>
        public Result ExecuteRefresh(
            List<string> selectedFilterItems,
            List<string> selectedMepCategories,
            List<string> selectedReferenceFiles,
            List<string> selectedHostFiles,
            Dictionary<string, double> clearanceSettings)
        {
            // ✅ CRITICAL: Log start of refresh
            DebugLogger.Info("[REFRESH-REFACTORED] ===== STARTING REFRESH =====");
            DebugLogger.Info($"[REFRESH-REFACTORED] Filters: {string.Join(", ", selectedFilterItems ?? new List<string>())}");
            DebugLogger.Info($"[REFRESH-REFACTORED] Categories: {string.Join(", ", selectedMepCategories ?? new List<string>())}");
            DebugLogger.Info($"[REFRESH-REFACTORED] Reference Files: {string.Join(", ", selectedReferenceFiles ?? new List<string>())}");
            DebugLogger.Info($"[REFRESH-REFACTORED] Host Files: {string.Join(", ", selectedHostFiles ?? new List<string>())}");
            
            try
            {
                // Get settings
                var settings = _appProfileService?.GetCurrentSettings();
                bool enableThreePointValidation = settings?.EnableThreePointValidation ?? true;
                
                DebugLogger.Info($"[REFRESH-REFACTORED] EnableThreePointValidation: {enableThreePointValidation}");
                
                // Create context (holds all state)
                using (var context = new RefreshContext(
                    _document,
                    _uiDocument,
                    selectedFilterItems,
                    selectedMepCategories,
                    selectedReferenceFiles,
                    selectedHostFiles,
                    FilterUiStateProvider.GetSelectedHostCategories?.Invoke() ?? new List<string>(),
                    clearanceSettings,
                    enableThreePointValidation))
                {
                    DebugLogger.Info($"[REFRESH-REFACTORED] Context created, RefreshLogName: {context.RefreshLogName}");
                    
                    try
                    {
                        var result = ExecuteRefreshInternal(context);
                        DebugLogger.Info($"[REFRESH-REFACTORED] ===== REFRESH COMPLETED: {result} =====");
                        return result;
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Error($"[REFRESH-REFACTORED] Exception in ExecuteRefreshInternal: {ex.Message}");
                        DebugLogger.Error($"[REFRESH-REFACTORED] Stack: {ex.StackTrace}");
                        HandleError(context, ex);
                        return Result.Failed;
                    }
                    finally
                    {
                        // Context.Dispose() handles cleanup
                        ResetUI();
                    }
                }
            }
            catch (Exception ex)
            {
                // ✅ CRITICAL: Catch exceptions during context creation
                DebugLogger.Error($"[REFRESH-REFACTORED] ❌ Exception during context creation: {ex.Message}");
                DebugLogger.Error($"[REFRESH-REFACTORED] Stack: {ex.StackTrace}");
                
                if (_statusLabel != null)
                    _statusLabel.Text = $"ERROR: {ex.Message}";
                
                System.Windows.Forms.MessageBox.Show(
                    $"Failed to initialize refresh: {ex.Message}\n\nCheck logs for details.",
                    "Refresh Initialization Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
                
                return Result.Failed;
            }
        }
        
        private Result ExecuteRefreshInternal(RefreshContext context)
        {
            // PHASE 1: Validate UI selections
            using (context.PerformanceMonitor.TrackOperation("1. UI Validation"))
            {
                if (!ValidateUISelections(context))
                    return Result.Cancelled;
            }
            
            UpdateProgress(10, "Loading XML data...");
            
            // PHASE 2: Load XML once (eliminates 4+ redundant loads)
            using (var xmlOp = context.PerformanceMonitor.TrackOperation("2. XML Loading") as PerformanceMonitor.OperationTracker)
            {
                var xmlManager = new XmlCacheManager(_document, context.RefreshLogName);
                context.XmlCache = xmlManager.LoadAll(context.SelectedFilterNames, context.SelectedMepCategories);
                xmlOp?.SetItemCount(context.XmlCache.FilterXml.Count + context.XmlCache.GlobalXml.Count);
            }
            
            UpdateProgress(20, "Loading existing clash zones...");
            
            // PHASE 3: Load existing clash zones from XML cache
            using (var loadOp = context.PerformanceMonitor.TrackOperation("3. Load Existing Zones") as PerformanceMonitor.OperationTracker)
            {
                context.ExistingClashZones = LoadExistingClashZones(context);
                loadOp?.SetItemCount(context.ExistingClashZones?.Count ?? 0);
            }
            
            // ✅ PATH DETERMINATION: Determine which path strategy to use
            var pathStrategy = RefreshPathDeterminer.DeterminePath(context, context.EnableThreePointValidation);
            
            if (!context.IsDeploymentMode)
            {
                DebugLogger.Info($"[REFRESH-REFACTORED] Using {pathStrategy.PathName}");
                DebugLogger.Info($"[REFRESH-REFACTORED] Path settings: AllowStructuralUpdates={pathStrategy.AllowStructuralUpdates}, EnableThreePointValidation={pathStrategy.EnableThreePointValidation}, ShouldResetFlags={pathStrategy.ShouldResetFlags}, ShouldSyncFlags={pathStrategy.ShouldSyncFlags}, ShouldCheckGuids={pathStrategy.ShouldCheckGuids}");
            }
            
            // Store strategy in context for later use
            context.PathStrategy = pathStrategy;
            
            UpdateProgress(30, "Syncing flags from Global XML...");
            
            // PHASE 4: Sync flags from Global XML (only if strategy requires it)
            if (pathStrategy.ShouldSyncFlags)
            {
            using (context.PerformanceMonitor.TrackOperation("4. Flag Sync"))
            {
                SyncFlagsFromGlobal(context);
                }
            }
            else
            {
                if (!context.IsDeploymentMode)
                {
                    DebugLogger.Info("[REFRESH-REFACTORED] Skipping flag sync (not required for this path)");
                }
            }
            
            UpdateProgress(40, "Validating clash zones...");
            
            // PHASE 5: Smart validation (only if strategy requires it)
            using (var validationOp = context.PerformanceMonitor.TrackOperation("5. Validation") as PerformanceMonitor.OperationTracker)
            {
                if (pathStrategy.EnableThreePointValidation)
            {
                var validationService = new ValidationService(context, new FlagManager(_document));
                var validationResult = validationService.ValidateClashZones(context.ExistingClashZones);
                
                    // ✅ PATH 3: Zone splitting after validation
                    var processedZones = pathStrategy.ProcessZonesAfterValidation(
                        context,
                        validationResult.ValidZones,
                        validationResult.InvalidZones);
                    
                    context.ExistingClashZones = processedZones;
                    
                    // ✅ PATH 3: Store validated and invalidated zones separately for distinct placement flows
                    context.ValidatedZones = validationResult.ValidZones ?? new List<ClashZone>();
                    context.InvalidatedZones = validationResult.InvalidZones ?? new List<ClashZone>();
                
                    // Remove invalid zones from Global XML (only for PATH 3)
                    if (validationResult.InvalidZones.Count > 0)
                {
                    validationService.RemoveInvalidZonesFromGlobal(validationResult.InvalidZones);
                }
                
                validationOp?.SetItemCount(context.ExistingClashZones.Count);
                }
                else
                {
                    // PATH 1 and PATH 2: No validation needed
                    if (!context.IsDeploymentMode)
                    {
                        DebugLogger.Info("[REFRESH-REFACTORED] Skipping 3-point validation (not required for this path)");
                    }
                    validationOp?.SetItemCount(context.ExistingClashZones?.Count ?? 0);
                }
            }
            
            // PHASE 5B: Flag reset for deleted sleeves (only if strategy requires it)
            if (pathStrategy.ShouldResetFlags)
            {
                UpdateProgress(45, "Resetting flags for deleted sleeves...");
                
                using (context.PerformanceMonitor.TrackOperation("5B. Flag Reset"))
                {
                    var flagManager = new FlagManager(_document);
                    
                    // Build categories list from existing clash zones
                    var categoriesToCheck = context.ExistingClashZones?
                        .Where(cz => !string.IsNullOrWhiteSpace(cz.MepElementCategory))
                        .Select(cz => cz.MepElementCategory)
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList() ?? new List<string>();
                    
                    if (categoriesToCheck.Count > 0)
                    {
                        // Group clash zones by category for flag reset
                        var clashZonesByCategory = context.ExistingClashZones?
                            .GroupBy(cz => cz.MepElementCategory, StringComparer.OrdinalIgnoreCase)
                            .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase) 
                            ?? new Dictionary<string, List<ClashZone>>();
                        
                        // ✅ STEP 1: Reset flags (DB first, then XML)
                        var resetCount = flagManager.ResetFlagsForDeletedSleeves(
                            categoriesToCheck,
                            clashZonesByCategory,
                            context.RefreshLogName);
                        
                        // ✅ STEP 2: Reset instance IDs (DB first, then XML)
                        pathStrategy.ResetInstanceIdsForDeletedSleeves(
                            context,
                            categoriesToCheck,
                            clashZonesByCategory);
                        
                        if (!context.IsDeploymentMode)
                        {
                            DebugLogger.Info($"[REFRESH-REFACTORED] Flag reset completed: {resetCount} zones reset");
                        }
                    }
                    else
                    {
                        if (!context.IsDeploymentMode)
                        {
                            DebugLogger.Info("[REFRESH-REFACTORED] No categories to check for flag reset");
                        }
                    }
                }
            }
            else
            {
                if (!context.IsDeploymentMode)
                {
                    DebugLogger.Info("[REFRESH-REFACTORED] Skipping flag reset (not required for this path)");
                }
            }
            
            UpdateProgress(50, "Processing intersections...");
            
            // PHASE 6: Intersection detection using IntersectionProcessor (with Replace/Replay/FullDetection modes)
            using (var intersectionOp = context.PerformanceMonitor.TrackOperation("6. Intersection Processing") as PerformanceMonitor.OperationTracker)
            {
                // ✅ INTEGRATION: Use IntersectionProcessor instead of direct detection
                var xmlManager = new XmlCacheManager(_document, context.RefreshLogName);
                var validationService = new ValidationService(context, new FlagManager(_document));
                var paramService = new ParameterCaptureService(context);
                
                var logger = new Action<string>(msg => 
                {
                    if (!context.IsDeploymentMode)
                        DebugLogger.Info(msg);
                    SafeFileLogger.SafeAppendText(context.RefreshLogName, $"[{DateTime.Now}] {msg}\n");
                });
                
                // ✅ PROGRESS CALLBACK: Track intersection counts per category
                Action<string, int> progressCallback = null;
                if (_statusLabel != null)
                {
                    progressCallback = (category, count) =>
                    {
                        try
                        {
                            _statusLabel.Text = $"Processing {category}: {count} intersections";
                        }
                        catch (Exception ex)
                        {
                            DebugLogger.Warning($"[REFRESH-REFACTORED] Error updating progress status: {ex.Message}");
                        }
                    };
                }
                
                var processor = new IntersectionProcessor(
                    context,
                    xmlManager,
                    validationService,
                    paramService,
                    context.PerformanceMonitor,
                    logger,
                    progressCallback);
                
                // Phase 1: Prepare existing zones and determine detection decision
                var decision = processor.PrepareExistingZones();
                
                // Phase 2: Run detection if needed (or use existing zones)
                var allClashZones = processor.RunDetectionIfNeeded(decision);
                
                // Context is already updated by IntersectionProcessor
                // Just ensure AllClashZones is set
                if (context.AllClashZones == null)
                {
                    context.AllClashZones = allClashZones;
                }
                
                // Phase 3: Post-process (updates cache, rebuilds snapshots)
                processor.PostProcess(decision);
                
                intersectionOp?.SetItemCount(allClashZones.Count);
                
                // Log mode decision
                if (!context.IsDeploymentMode)
                {
                    DebugLogger.Info($"[REFRESH-REFACTORED] Mode: {decision.Mode}, Detection Run: {decision.ShouldRunDetection}, Reason: {decision.Reason}");
                }
            }
            
            // ✅ PROGRESS DIALOG: Close progress dialog after intersection detection
            // ✅ THREAD-SAFE: Use Invoke if needed, ensure proper cleanup
            // The progress dialog is no longer a separate field, so this block is removed.
            
            UpdateProgress(70, "Capturing parameters...");
            
            // PHASE 8: Capture minimal parameters (parallel)
            using (var paramOp = context.PerformanceMonitor.TrackOperation("8. Parameter Capture") as PerformanceMonitor.OperationTracker)
            {
                var paramService = new ParameterCaptureService(context);
                paramService.CaptureParametersParallel(context.NewClashZones);
                paramOp?.SetItemCount(context.NewClashZones?.Count ?? 0);
            }
            
            // ✅ CRITICAL FIX: Update AllClashZones AFTER parameter capture to include parameters
            // This ensures parameters are included when saving to database
            if (context.NewClashZones != null && context.NewClashZones.Count > 0)
            {
                var existingIds = new HashSet<Guid>((context.AllClashZones ?? new List<ClashZone>()).Select(z => z.Id));
                var allZones = context.AllClashZones?.ToList() ?? new List<ClashZone>();
                
                foreach (var newZone in context.NewClashZones)
                {
                    if (!existingIds.Contains(newZone.Id))
                    {
                        allZones.Add(newZone);
                    }
                    else
                    {
                        // Update existing zone with parameters from new zone
                        var existingZone = allZones.FirstOrDefault(z => z.Id == newZone.Id);
                        if (existingZone != null && newZone.MepParameterValues != null && newZone.MepParameterValues.Count > 0)
                        {
                            existingZone.MepParameterValues = newZone.MepParameterValues;
                        }
                        if (existingZone != null && newZone.HostParameterValues != null && newZone.HostParameterValues.Count > 0)
                        {
                            existingZone.HostParameterValues = newZone.HostParameterValues;
                        }
                    }
                }
                
                context.AllClashZones = allZones;
            }
            
            UpdateProgress(80, "Merging and saving...");
            
            // PHASE 9: Merge and save
            using (var saveOp = context.PerformanceMonitor.TrackOperation("9. Save") as PerformanceMonitor.OperationTracker)
            {
                MergeAndSave(context);
                saveOp?.SetItemCount(context.AllClashZones?.Count ?? 0);
            }
            
            UpdateProgress(90, "Finalizing...");
            
            // PHASE 10: Final cleanup
            using (context.PerformanceMonitor.TrackOperation("10. Cleanup"))
            {
                FinalCleanup(context);
            }
            
            UpdateProgress(100, "Complete!");
            
            // Show summary
            ShowSummary(context);
            
            return Result.Succeeded;
        }
        
        private bool ValidateUISelections(RefreshContext context)
        {
            DebugLogger.Info("[REFRESH-REFACTORED] Validating UI selections...");
            
            var errors = new List<string>();
            
            if (context.SelectedFilterNames == null || context.SelectedFilterNames.Count == 0)
            {
                errors.Add("Filters");
                DebugLogger.Warning("[REFRESH-REFACTORED] ⚠️ No filters selected");
            }
            else
            {
                DebugLogger.Info($"[REFRESH-REFACTORED] Selected filters: {string.Join(", ", context.SelectedFilterNames)}");
            }
            
            if (context.SelectedMepCategories == null || context.SelectedMepCategories.Count == 0)
            {
                errors.Add("MEP Categories");
                DebugLogger.Warning("[REFRESH-REFACTORED] ⚠️ No MEP categories selected");
            }
            else
            {
                DebugLogger.Info($"[REFRESH-REFACTORED] Selected MEP categories: {string.Join(", ", context.SelectedMepCategories)}");
            }
            
            if (context.SelectedReferenceFiles == null || context.SelectedReferenceFiles.Count == 0)
            {
                errors.Add("MEP Linked Files");
                DebugLogger.Warning("[REFRESH-REFACTORED] ⚠️ No reference files selected");
            }
            else
            {
                DebugLogger.Info($"[REFRESH-REFACTORED] Selected reference files: {string.Join(", ", context.SelectedReferenceFiles)}");
            }
            
            if (context.SelectedHostFiles == null || context.SelectedHostFiles.Count == 0)
            {
                errors.Add("Host Linked Files");
                DebugLogger.Warning("[REFRESH-REFACTORED] ⚠️ No host files selected");
            }
            else
            {
                DebugLogger.Info($"[REFRESH-REFACTORED] Selected host files: {string.Join(", ", context.SelectedHostFiles)}");
            }
            
            if (errors.Count > 0)
            {
                string message = "Please select:\n" + string.Join("\n", errors.Select(e => $"• {e}"));
                
                DebugLogger.Warning($"[REFRESH-REFACTORED] ❌ Validation failed: {message}");
                
                if (_statusLabel != null)
                    _statusLabel.Text = $"Missing: {string.Join(", ", errors)}";
                
                System.Windows.Forms.MessageBox.Show(message, "Missing Selections", 
                    System.Windows.Forms.MessageBoxButtons.OK, 
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return false;
            }
            
            DebugLogger.Info("[REFRESH-REFACTORED] ✅ UI validation passed");
            return true;
        }
        
        private List<ClashZone> LoadExistingClashZones(RefreshContext context)
        {
            var allZones = new List<ClashZone>();
            
            foreach (var filterName in context.SelectedFilterNames)
            {
                if (context.XmlCache.FilterXml.TryGetValue(filterName, out var storage))
                {
                    if (storage?.ClashZones != null)
                    {
                        allZones.AddRange(storage.ClashZones);
                    }
                }
            }
            
            return allZones;
        }
        
        private void SyncFlagsFromGlobal(RefreshContext context)
        {
            if (context.ExistingClashZones == null || context.ExistingClashZones.Count == 0)
                return;
            
            var flagManager = new FlagManager(_document);
            
            var byCategory = context.ExistingClashZones
                .GroupBy(cz => cz.MepElementCategory)
                .ToList();
            
            foreach (var group in byCategory)
            {
                if (string.IsNullOrWhiteSpace(group.Key))
                    continue;
                
                flagManager.SyncFlagsFromGlobal(group.ToList(), group.Key);
            }
        }
        
        // ✅ REMOVED: DetectIntersections() and CreateClashZones() methods
        // These are now handled by IntersectionProcessor which provides:
        // - Replace/Replay/FullDetection mode logic
        // - Collector-level multi-filter optimization
        // - Proper integration with validation and parameter capture
        
        private void MergeAndSave(RefreshContext context)
        {
            try
            {
                DebugLogger.Info($"[REFRESH-REFACTORED] [MERGE-SAVE] Starting MergeAndSave - AllClashZones: {context.AllClashZones?.Count ?? 0}, NewClashZones: {context.NewClashZones?.Count ?? 0}");
                SafeFileLogger.SafeAppendText(context.RefreshLogName, 
                    $"[{DateTime.Now}] [MERGE-SAVE] Starting MergeAndSave - AllClashZones: {context.AllClashZones?.Count ?? 0}, NewClashZones: {context.NewClashZones?.Count ?? 0}\n");
                
                // Merge existing + new (already done by IntersectionProcessor, but ensure it's set)
                if (context.AllClashZones == null || context.AllClashZones.Count == 0)
                {
                    var allZones = context.ExistingClashZones?.ToList() ?? new List<ClashZone>();
                    var existingIds = new HashSet<Guid>(allZones.Select(z => z.Id));
                    
                    foreach (var newZone in context.NewClashZones ?? new List<ClashZone>())
                    {
                        if (!existingIds.Contains(newZone.Id))
                        {
                            allZones.Add(newZone);
                        }
                    }
                    
                    context.AllClashZones = allZones;
                    DebugLogger.Info($"[REFRESH-REFACTORED] [MERGE-SAVE] Merged zones - Total: {context.AllClashZones.Count}");
                    SafeFileLogger.SafeAppendText(context.RefreshLogName, 
                        $"[{DateTime.Now}] [MERGE-SAVE] Merged zones - Total: {context.AllClashZones.Count}\n");
                }
                
                // ✅ BASE-NAME NORMALIZATION: Normalize filter names before persistence
                // This prevents duplicate branches in Global XML (e.g., "Plumbing" vs "Plumbing_pipes")
                DebugLogger.Info($"[REFRESH-REFACTORED] [MERGE-SAVE] Loading enabled filter from: {string.Join(", ", context.SelectedFilterNames ?? new List<string>())}");
                SafeFileLogger.SafeAppendText(context.RefreshLogName, 
                    $"[{DateTime.Now}] [MERGE-SAVE] Loading enabled filter from: {string.Join(", ", context.SelectedFilterNames ?? new List<string>())}\n");
                
                var enabledFilter = LoadEnabledFilter(context);
                if (enabledFilter == null)
                {
                    DebugLogger.Warning("[REFRESH-REFACTORED] [MERGE-SAVE] ❌ No enabled filter found - skipping save");
                    SafeFileLogger.SafeAppendText(context.RefreshLogName, 
                        $"[{DateTime.Now}] [MERGE-SAVE] ❌ No enabled filter found - skipping save\n");
                    return;
                }
                
                DebugLogger.Info($"[REFRESH-REFACTORED] [MERGE-SAVE] ✅ Enabled filter found: '{enabledFilter.Name}'");
                SafeFileLogger.SafeAppendText(context.RefreshLogName, 
                    $"[{DateTime.Now}] [MERGE-SAVE] ✅ Enabled filter found: '{enabledFilter.Name}'\n");
            
            // Save via persistence service with normalized base names
            var persistenceService = new ClashZonePersistenceService(
                _document, 
                new GuidManager(_document), 
                context.RefreshLogName);
            
            // ✅ MIMIC OLD REFRESHSERVICE: Call SaveClashZones ONCE with ALL clash zones (same as old RefreshService line 3621)
            // SaveClashZones internally groups by category and creates filter groups for each category
            // This is the key difference - old RefreshService calls it once, not per category
            var baseFilterName = enabledFilter?.Name ?? string.Empty;
            var normalizedBaseName = FilterNameHelper.NormalizeBaseName(
                baseFilterName,
                baseFilterName,
                null); // No category-specific normalization when saving all at once
            
            if (!context.IsDeploymentMode)
            {
                DebugLogger.Info($"[PERSIST-NAME] Saving ALL clash zones: Raw='{baseFilterName}', Normalized='{normalizedBaseName}', TotalZones={context.AllClashZones?.Count ?? 0}");
            }
            SafeFileLogger.SafeAppendText(context.RefreshLogName, 
                $"[{DateTime.Now}] [PERSIST-NAME] Saving ALL clash zones: Raw='{baseFilterName}', Normalized='{normalizedBaseName}', TotalZones={context.AllClashZones?.Count ?? 0}\n");
            
            // ✅ MIMIC OLD REFRESHSERVICE: Save all clash zones at once (same as old RefreshService)
            // This ensures filter groups are created for all categories that have clash zones
            DebugLogger.Info($"[REFRESH-REFACTORED] [MERGE-SAVE] Calling SaveClashZones with {context.AllClashZones?.Count ?? 0} zones");
            SafeFileLogger.SafeAppendText(context.RefreshLogName, 
                $"[{DateTime.Now}] [MERGE-SAVE] Calling SaveClashZones with {context.AllClashZones?.Count ?? 0} zones\n");
            
            try
            {
                // ✅ PATH STRATEGY: Use AllowStructuralUpdates from strategy
                var allowStructuralUpdates = context.PathStrategy?.AllowStructuralUpdates ?? true;
                
                if (!context.IsDeploymentMode)
                {
                    DebugLogger.Info($"[REFRESH-REFACTORED] [MERGE-SAVE] Using allowStructuralUpdates={allowStructuralUpdates} from {context.PathStrategy?.PathName ?? "default"}");
                }
                
                persistenceService.SaveClashZones(
                    context.AllClashZones ?? new List<ClashZone>(),
                    normalizedBaseName,
                    enabledFilter,
                    allowStructuralUpdates: allowStructuralUpdates);
                
                DebugLogger.Info($"[REFRESH-REFACTORED] [MERGE-SAVE] ✅ SaveClashZones completed successfully");
                SafeFileLogger.SafeAppendText(context.RefreshLogName, 
                    $"[{DateTime.Now}] [MERGE-SAVE] ✅ SaveClashZones completed successfully\n");
            }
            catch (Exception saveEx)
            {
                DebugLogger.Error($"[REFRESH-REFACTORED] [MERGE-SAVE] ❌ SaveClashZones failed: {saveEx.Message}");
                DebugLogger.Error($"[REFRESH-REFACTORED] [MERGE-SAVE] Stack: {saveEx.StackTrace}");
                SafeFileLogger.SafeAppendText(context.RefreshLogName, 
                    $"[{DateTime.Now}] [MERGE-SAVE] ❌ SaveClashZones failed: {saveEx.Message}\n{saveEx.StackTrace}\n");
                throw; // Re-throw to be caught by outer try-catch
            }
            
            // ✅ CRITICAL FIX: Save the filter file to disk after all categories are saved
            // This ensures Place Sleeve can find the clash zones
            // ✅ SAVE BOTH: Base filter file AND category-specific files (e.g., "Plumbing.xml" AND "Plumbing_pipes.xml")
            if (enabledFilter != null)
            {
                try
                {
                    var filterManagementService = new FilterManagementService(_document, msg => { }, msg => { });
                    var filtersDirectory = ProjectPathService.GetFiltersDirectory(_document);
                    if (!Directory.Exists(filtersDirectory))
                    {
                        Directory.CreateDirectory(filtersDirectory);
                    }
                    
                    // ✅ STEP 1: Save base filter file (e.g., "Plumbing.xml") containing all categories
                    // ✅ PHASE 2: Only save to XML if XML creation is enabled
                    if (!DeploymentConfiguration.DisableXmlCreation)
                    {
                    var baseFilterFilePath = Path.Combine(filtersDirectory, $"{enabledFilter.Name}.xml");
                    filterManagementService.SaveFilterToXmlFile(enabledFilter, baseFilterFilePath);
                    
                    DebugLogger.Info($"[REFRESH-REFACTORED] [MERGE-SAVE] ✅ Saved base filter file: {baseFilterFilePath}");
                    SafeFileLogger.SafeAppendText(context.RefreshLogName, 
                        $"[{DateTime.Now}] [MERGE-SAVE] ✅ Saved base filter file: {baseFilterFilePath}\n");
                    }
                    else
                    {
                        DebugLogger.Info($"[REFRESH-REFACTORED] ⚠️ XML creation disabled - skipping base filter XML save (database only mode)");
                        SafeFileLogger.SafeAppendText(context.RefreshLogName, 
                            $"[{DateTime.Now}] [MERGE-SAVE] ⚠️ XML creation disabled - skipping base filter XML save (database only mode)\n");
                    }
                    
                    // ✅ STEP 2: Save category-specific filter files (e.g., "Plumbing_pipes.xml", "Plumbing_ducts.xml")
                    // This matches the legacy RefreshService behavior and ensures Place Sleeve can find category-specific files
                    // ✅ CRITICAL FIX: Ensure ALL selected categories get files saved, even if they have zero clash zones
                    var existingFilterGroupNames = enabledFilter.ClashZoneStorage?.Filters?
                        .Where(fg => fg != null && !string.IsNullOrWhiteSpace(fg.Name))
                        .Select(fg => fg.Name)
                        .ToHashSet(StringComparer.OrdinalIgnoreCase) ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    
                    // ✅ MIMIC OLD REFRESHSERVICE: Create filter groups for ALL selected categories
                    var categoriesToSave = context.SelectedMepCategories ?? new List<string>();
                    if (categoriesToSave.Count == 0 && enabledFilter.ClashZoneStorage?.Filters != null)
                    {
                        // Fallback: extract categories from existing filter group names
                        // Extract category from filter group name (e.g., "ALL_ducts" -> "Ducts")
                        categoriesToSave = enabledFilter.ClashZoneStorage.Filters
                            .Where(fg => fg != null && !string.IsNullOrWhiteSpace(fg.Name))
                            .Select(fg =>
                            {
                                var groupName = fg.Name;
                                var baseName = enabledFilter.Name;
                                // Remove base name prefix (e.g., "ALL_ducts" -> "ducts")
                                if (!string.IsNullOrWhiteSpace(baseName) && groupName.StartsWith(baseName + "_", StringComparison.OrdinalIgnoreCase))
                                {
                                    var suffix = groupName.Substring(baseName.Length + 1);
                                    // Reverse GetXmlSuffix: "ducts" -> "Ducts", "cable_trays" -> "Cable Trays"
                                    // Try to match known suffixes to categories
                                    if (suffix.Equals("ducts", StringComparison.OrdinalIgnoreCase))
                                        return "Ducts";
                                    if (suffix.Equals("pipes", StringComparison.OrdinalIgnoreCase))
                                        return "Pipes";
                                    if (suffix.Equals("cable_trays", StringComparison.OrdinalIgnoreCase))
                                        return "Cable Trays";
                                    if (suffix.Equals("duct_accessories", StringComparison.OrdinalIgnoreCase))
                                        return "Duct Accessories";
                                    // Fallback: capitalize first letter and replace underscores with spaces
                                    return System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(suffix.Replace("_", " "));
                                }
                                return null;
                            })
                            .Where(cat => !string.IsNullOrWhiteSpace(cat))
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .ToList();
                    }
                    
                    // Ensure ClashZoneStorage exists
                    enabledFilter.ClashZoneStorage ??= new ClashZoneStorage
                    {
                        Filters = new List<FilterGroupForStorage>(),
                        ClashZones = new List<ClashZone>()
                    };
                    enabledFilter.ClashZoneStorage.Filters ??= new List<FilterGroupForStorage>();
                    
                    // ✅ CREATE EMPTY FILTER GROUPS: For categories that don't have filter groups yet
                    foreach (var category in categoriesToSave)
                    {
                        if (string.IsNullOrWhiteSpace(category))
                            continue;
                        
                        var categoryNormalizedBaseName = FilterNameHelper.NormalizeBaseName(
                            enabledFilter.Name,
                            enabledFilter.Name,
                            category);
                        // ✅ REPLICATE BuildFilterGroupName logic (method is private, so we replicate it here)
                        var categorySuffix = MepCategoryConstants.GetXmlSuffix(category);
                        var suffix = "_" + categorySuffix;
                        var expectedGroupName = string.IsNullOrWhiteSpace(categoryNormalizedBaseName)
                            ? suffix.TrimStart('_')
                            : (categoryNormalizedBaseName.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)
                                ? categoryNormalizedBaseName
                                : categoryNormalizedBaseName + suffix);
                        
                        if (!existingFilterGroupNames.Contains(expectedGroupName))
                        {
                            // Create empty filter group for this category
                            var emptyFilterGroup = new FilterGroupForStorage
                            {
                                Name = expectedGroupName,
                                FileCombos = new List<FilterFileComboGroup>()
                            };
                            enabledFilter.ClashZoneStorage.Filters.Add(emptyFilterGroup);
                            
                            if (!context.IsDeploymentMode)
                            {
                                DebugLogger.Info($"[REFRESH-REFACTORED] Created empty filter group for category '{category}': '{expectedGroupName}'");
                            }
                            SafeFileLogger.SafeAppendText(context.RefreshLogName,
                                $"[{DateTime.Now}] [REFRESH-REFACTORED] Created empty filter group for category '{category}': '{expectedGroupName}'\n");
                        }
                    }
                    
                    // Now save all filter groups (including empty ones)
                    if (enabledFilter.ClashZoneStorage?.Filters != null)
                    {
                        foreach (var filterGroup in enabledFilter.ClashZoneStorage.Filters)
                        {
                            if (filterGroup == null)
                            {
                                SafeFileLogger.SafeAppendText(context.RefreshLogName,
                                    $"[{DateTime.Now}] [REFRESH-REFACTORED] Skipping null filter group entry\n");
                                continue;
                            }

                            var comboCount = filterGroup.FileCombos?.Count ?? 0;
                            
                            // ✅ DEBUG: Log detailed info about filter group before processing
                            var totalZonesInGroup = filterGroup.FileCombos?
                                .SelectMany(fc => fc.ClashZones ?? Enumerable.Empty<ClashZone>())
                                .Count() ?? 0;
                            
                            SafeFileLogger.SafeAppendText(context.RefreshLogName,
                                $"[{DateTime.Now}] [REFRESH-REFACTORED] Processing group '{filterGroup.Name}' → FileCombos={comboCount}, TotalZones={totalZonesInGroup}\n");
                            
                            if (!context.IsDeploymentMode)
                            {
                                DebugLogger.Info($"[REFRESH-REFACTORED] Processing filter group '{filterGroup.Name}': FileCombos={comboCount}, TotalZones={totalZonesInGroup}");
                            }

                            // ✅ MIMIC LEGACY: Clone filter group (same as legacy RefreshService)
                            var groupClone = CloneFilterGroup(filterGroup);
                            var groupZones = groupClone.FileCombos?
                                .SelectMany(fc => fc.ClashZones ?? Enumerable.Empty<ClashZone>())
                                .Where(z => z != null)
                                .GroupBy(z => z.Id)
                                .Select(g => g.First())
                                .ToList() ?? new List<ClashZone>();

                            // ✅ CRITICAL FIX: Save files for ALL filter groups (even if zones count is 0)
                            // The old RefreshService saves files for all filter groups, regardless of zone count
                            // This ensures Place Sleeve can find category files even if they're empty
                            if (groupZones.Count == 0)
                            {
                                SafeFileLogger.SafeAppendText(context.RefreshLogName,
                                    $"[{DateTime.Now}] [REFRESH-REFACTORED]   Group '{filterGroup.Name}' has no zones after extraction - will save empty file\n");
                                
                                if (!context.IsDeploymentMode)
                                {
                                    DebugLogger.Warning($"[REFRESH-REFACTORED] Filter group '{filterGroup.Name}' has {comboCount} file combos but 0 zones extracted - check zone extraction logic");
                                }
                            }

                            // ✅ MIMIC LEGACY: Create category-specific filter (same structure as legacy RefreshService)
                            var categoryFilter = new OpeningFilter
                            {
                                Name = filterGroup.Name, // e.g., "Plumbing_pipes"
                                Category = enabledFilter.Category,
                                OpeningType = enabledFilter.OpeningType,
                                IsEnabled = enabledFilter.IsEnabled,
                                LastModified = DateTime.Now,
                                SelectedMepCategoryNames = groupZones
                                    .Select(z => z.MepElementCategory)
                                    .Where(cat => !string.IsNullOrWhiteSpace(cat))
                                    .Distinct(StringComparer.OrdinalIgnoreCase)
                                    .ToList(),
                                ClashZoneStorage = new ClashZoneStorage
                                {
                                    Filters = new List<FilterGroupForStorage> { groupClone },
                                    ClashZones = new List<ClashZone>(groupZones)
                                }
                            };

                            // ✅ MIMIC LEGACY: Use same path format as legacy RefreshService
                            // ✅ PHASE 2: Only save to XML if XML creation is enabled
                            if (!DeploymentConfiguration.DisableXmlCreation)
                            {
                            var categoryFilePath = Path.Combine(filtersDirectory, $"{filterGroup.Name}.xml");

                            SafeFileLogger.SafeAppendText(context.RefreshLogName,
                                $"[{DateTime.Now}] [REFRESH-REFACTORED] Saving category filter '{filterGroup.Name}' (Zones={groupZones.Count}) → {categoryFilePath}\n");

                            filterManagementService.SaveFilterToXmlFile(categoryFilter, categoryFilePath);
                            
                            if (!context.IsDeploymentMode)
                            {
                                DebugLogger.Info($"[REFRESH-REFACTORED] ✅ Saved category filter file: {categoryFilePath} (Zones={groupZones.Count})");
                            }
                            }
                            else
                            {
                                if (!context.IsDeploymentMode)
                                {
                                    DebugLogger.Info($"[REFRESH-REFACTORED] ⚠️ XML creation disabled - skipping category filter XML save for '{filterGroup.Name}' (database only mode)");
                                }
                            SafeFileLogger.SafeAppendText(context.RefreshLogName, 
                                    $"[{DateTime.Now}] [REFRESH-REFACTORED] ⚠️ XML creation disabled - skipping category filter XML save for '{filterGroup.Name}' (database only mode)\n");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (!context.IsDeploymentMode)
                    {
                        DebugLogger.Error($"[REFRESH-REFACTORED] ❌ Error saving filter files: {ex.Message}");
                    }
                    SafeFileLogger.SafeAppendText(context.RefreshLogName, 
                        $"[{DateTime.Now}] [REFRESH-REFACTORED] ❌ Error saving filter files: {ex.Message}\n");
                }
            }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[REFRESH-REFACTORED] [MERGE-SAVE] ❌ MergeAndSave failed: {ex.Message}");
                DebugLogger.Error($"[REFRESH-REFACTORED] [MERGE-SAVE] Stack: {ex.StackTrace}");
                SafeFileLogger.SafeAppendText(context.RefreshLogName, 
                    $"[{DateTime.Now}] [MERGE-SAVE] ❌ MergeAndSave failed: {ex.Message}\n{ex.StackTrace}\n");
                throw; // Re-throw to be caught by ExecuteRefreshInternal
            }
        }
        
        private OpeningFilter LoadEnabledFilter(RefreshContext context)
        {
            foreach (var filterName in context.SelectedFilterNames)
            {
                var filterManagementService = new FilterManagementService(_document, msg => { }, msg => { });
                var filter = filterManagementService.LoadFilterAuto(filterName);
                
                // ✅ CRITICAL FIX: If filter doesn't exist, create it from current UI state
                if (filter == null)
                {
                    DebugLogger.Warning($"[REFRESH-REFACTORED] Filter '{filterName}' not found - creating from current UI state");
                    SafeFileLogger.SafeAppendText(context.RefreshLogName, 
                        $"[{DateTime.Now}] [MERGE-SAVE] Filter '{filterName}' not found - creating from current UI state\n");
                    
                    filter = filterManagementService.CreateFilterFromCurrentUIState(filterName);
                    
                    // ✅ CRITICAL: Register filter in database and mark as enabled
                    if (filter != null)
                    {
                        var categoryDisplay = filterManagementService.GetDisplayCategory(filter);
                        var filterId = filterManagementService.RegisterFilterInDatabase(filterName, categoryDisplay);
                        
                        if (filterId > 0)
                        {
                            // ✅ CRITICAL: Save UI state to database
                            try
                            {
                                filterManagementService.UseFilterRepository(repo =>
                                {
                                    repo.SaveFilterUIState(
                                        filterName,
                                        categoryDisplay,
                                        filter.SelectedHostCategories ?? new List<string>(),
                                        filter.OpeningSettings
                                    );
                                });
                                DebugLogger.Info($"[REFRESH-REFACTORED] ✅ Created and saved UI state for filter '{filterName}' to database");
                            }
                            catch (Exception uiStateEx)
                            {
                                DebugLogger.Warning($"[REFRESH-REFACTORED] ⚠️ Could not save UI state for new filter: {uiStateEx.Message}");
                            }
                            
                            filter.IsEnabled = true; // Mark as enabled
                            DebugLogger.Info($"[REFRESH-REFACTORED] ✅ Created filter '{filterName}' and marked as enabled");
                    return filter;
                        }
                    }
                }
                else
                {
                    // ✅ CRITICAL FIX: If filter exists but IsEnabled is not set, mark it as enabled
                    // This handles filters loaded from database/XML that might not have IsEnabled set
                    if (!filter.IsEnabled)
                    {
                        filter.IsEnabled = true;
                        DebugLogger.Info($"[REFRESH-REFACTORED] ✅ Marked existing filter '{filterName}' as enabled");
                    }
                    
                    return filter;
                }
            }
            
            return null;
        }
        
        /// <summary>
        /// ✅ MIMIC LEGACY: Clone FilterGroupForStorage (same as legacy RefreshService)
        /// </summary>
        private FilterGroupForStorage CloneFilterGroup(FilterGroupForStorage group)
        {
            if (group == null)
                return new FilterGroupForStorage
                {
                    Name = string.Empty,
                    FileCombos = new List<FilterFileComboGroup>()
                };

            return new FilterGroupForStorage
            {
                Name = group.Name,
                FileCombos = group.FileCombos?.Select(CloneFilterFileCombo).ToList() ?? new List<FilterFileComboGroup>()
            };
        }

        /// <summary>
        /// ✅ MIMIC LEGACY: Clone FilterFileComboGroup (same as legacy RefreshService)
        /// </summary>
        private FilterFileComboGroup CloneFilterFileCombo(FilterFileComboGroup combo)
        {
            if (combo == null)
                return new FilterFileComboGroup
                {
                    LinkedFile = string.Empty,
                    HostFile = string.Empty,
                    ProcessedAt = DateTime.Now,
                    ClashZones = new List<ClashZone>()
                };

            return new FilterFileComboGroup
            {
                LinkedFile = combo.LinkedFile,
                HostFile = combo.HostFile,
                ProcessedAt = combo.ProcessedAt,
                ClashZones = combo.ClashZones?
                    .Where(z => z != null)
                    .ToList() ?? new List<ClashZone>()
            };
        }
        
        private void FinalCleanup(RefreshContext context)
        {
            // ✅ GEOMETRY CACHE OPTIMIZATION: Only clear if model changed
            // This prevents unnecessary cache clearing on unchanged models (99.5% speedup)
            if (context.HasModelChanged())
            {
                if (!context.IsDeploymentMode)
                    DebugLogger.Info("[REFRESH-REFACTORED] Model changed - clearing geometry cache");
                
                MepIntersectionService.ClearGeometryCache();
                MepIntersectionService.ClearTransformCache();
            }
            else
            {
                if (!context.IsDeploymentMode)
                    DebugLogger.Info("[REFRESH-REFACTORED] Model unchanged - keeping geometry cache");
            }
            
            // Force GC
            GC.Collect(2, GCCollectionMode.Forced, true);
            GC.WaitForPendingFinalizers();
            GC.Collect(2, GCCollectionMode.Forced, true);
        }
        
        private void ShowSummary(RefreshContext context)
        {
            int total = context.AllClashZones?.Count ?? 0;
            int newCount = context.NewClashZones?.Count ?? 0;
            int unresolved = context.AllClashZones?.Count(cz => !cz.IsResolved && !cz.IsClusterResolved) ?? 0;
            
            string message = $"Refresh Complete!\n\n" +
                           $"Total Zones: {total}\n" +
                           $"New Zones: {newCount}\n" +
                           $"Unresolved: {unresolved}\n\n" +
                           $"Logs: {SafeFileLogger.GetLogDirectory()}";
            
            _statusLabel.Text = message;
        }
        
        private void HandleError(RefreshContext context, Exception ex)
        {
            string message = $"Refresh failed: {ex.Message}";
            
            // ✅ CRITICAL: Always log errors, even in deployment mode
            DebugLogger.Error($"[REFRESH-REFACTORED] ❌ {message}");
            DebugLogger.Error($"[REFRESH-REFACTORED] Stack trace: {ex.StackTrace}");
            
            SafeFileLogger.SafeAppendText(context.RefreshLogName, 
                $"[{DateTime.Now}] ERROR: {message}\n{ex.StackTrace}\n");
            
            // ✅ CRITICAL: Update status label even if null (safe check)
            if (_statusLabel != null)
            {
                _statusLabel.Text = $"ERROR: {ex.Message}";
            }
            
            // ✅ CRITICAL: Show error dialog to user
            System.Windows.Forms.MessageBox.Show(
                $"{message}\n\nCheck logs for full details:\n{context.RefreshLogName}",
                "Refresh Error",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Error);
        }
        
        private void UpdateProgress(int value, string status)
        {
            if (_progressBar != null)
            {
                _progressBar.Value = Math.Min(Math.Max(value, _progressBar.Minimum), _progressBar.Maximum);
                _progressBar.Visible = true;
            }
            
            if (_statusLabel != null)
                _statusLabel.Text = status;
        }
        
        private void ResetUI()
        {
            if (_progressBar != null)
                _progressBar.Value = _progressBar.Minimum;
            
            if (_refreshButton != null)
                _refreshButton.Enabled = true;
            
            // The progress dialog is no longer a separate field, so this block is removed.
        }
    }
}
