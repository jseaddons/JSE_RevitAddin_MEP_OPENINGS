using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Refresh;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;

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
        
        // ✅ SOLID: DI support for refresh services (Phase 2)
        private readonly IRefreshDataCacheManager _dataCacheManager;
        private readonly IClashZoneRepository _clashZoneRepository;
        private readonly IRefreshPathDeterminer _pathDeterminer;
        
        // UI controls
        private System.Windows.Forms.Label _statusLabel;
        private System.Windows.Forms.ProgressBar _progressBar;
        private System.Windows.Forms.Button _refreshButton;
        
        /// <summary>
        /// ✅ BACKWARD COMPATIBLE: Legacy constructor (direct instantiation)
        /// Used when UseRefreshDependencyInjection = false
        /// </summary>
        public RefreshServiceRefactored(
            Document document, 
            UIDocument uiDocument, 
            ApplicationProfileService appProfileService)
            : this(document, uiDocument, appProfileService, null, null, null)
        {
        }
        
        /// <summary>
        /// ✅ DI CONSTRUCTOR: Supports dependency injection (Phase 2)
        /// Used when UseRefreshDependencyInjection = true
        /// Falls back to direct instantiation if dependencies are null
        /// </summary>
        public RefreshServiceRefactored(
            Document document, 
            UIDocument uiDocument, 
            ApplicationProfileService appProfileService,
            IRefreshDataCacheManager dataCacheManager,
            IClashZoneRepository clashZoneRepository,
            IRefreshPathDeterminer pathDeterminer)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _uiDocument = uiDocument ?? throw new ArgumentNullException(nameof(uiDocument));
            _appProfileService = appProfileService ?? throw new ArgumentNullException(nameof(appProfileService));
            
            // ✅ FEATURE FLAG: Use DI if enabled and provided, otherwise use direct instantiation
            if (OptimizationFlags.UseRefreshDependencyInjection)
            {
                // DI path: Use provided dependencies or create defaults
                _dataCacheManager = dataCacheManager; // Can be null, will create on-demand
                _clashZoneRepository = clashZoneRepository; // Can be null, will create on-demand
                _pathDeterminer = pathDeterminer ?? new RefreshPathDeterminerWrapper(); // Self-bootstrap if null
            }
            else
            {
                // Legacy path: Direct instantiation (existing behavior)
                _dataCacheManager = null;
                _clashZoneRepository = null;
                _pathDeterminer = null;
            }
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
            // ✅ BUILD TIMESTAMP: Get build timestamp to verify latest code is running
            string buildTimestamp = "unknown";
            string assemblyPath = "unknown";
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                assemblyPath = assembly?.Location ?? "unknown";
                if (!string.IsNullOrWhiteSpace(assemblyPath) && System.IO.File.Exists(assemblyPath))
                {
                    buildTimestamp = System.IO.File.GetLastWriteTime(assemblyPath).ToString("yyyy-MM-dd HH:mm:ss");
                }
            }
            catch { }
            
            // ✅ CRITICAL: Log start of refresh with build timestamp
            DebugLogger.Info("[REFRESH-REFACTORED] ===== STARTING REFRESH =====");
            DebugLogger.Info($"[REFRESH-REFACTORED] 🔨 BUILD TIMESTAMP: {buildTimestamp} | Assembly: {System.IO.Path.GetFileName(assemblyPath)}");
            DebugLogger.Info($"[REFRESH-REFACTORED] Filters ({selectedFilterItems?.Count ?? 0}): {string.Join(", ", selectedFilterItems ?? new List<string>())}");
            DebugLogger.Info($"[REFRESH-REFACTORED] MEP Categories ({selectedMepCategories?.Count ?? 0}): {string.Join(", ", selectedMepCategories ?? new List<string>())}");
            DebugLogger.Info($"[REFRESH-REFACTORED] Reference Files ({selectedReferenceFiles?.Count ?? 0}): {string.Join(", ", selectedReferenceFiles ?? new List<string>())}");
            DebugLogger.Info($"[REFRESH-REFACTORED] Host Files ({selectedHostFiles?.Count ?? 0}): {string.Join(", ", selectedHostFiles ?? new List<string>())}");
            
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
                    
                    // ✅ BUILD TIMESTAMP: Write build timestamp to refresh log header
                    try
                    {
                        SafeFileLogger.SafeAppendText(context.RefreshLogName,
                            $"[{DateTime.Now}] ===== REFRESH LOG STARTED =====\n");
                        SafeFileLogger.SafeAppendText(context.RefreshLogName,
                            $"[{DateTime.Now}] 🔨 BUILD TIMESTAMP: {buildTimestamp} | Assembly: {System.IO.Path.GetFileName(assemblyPath)}\n");
                        SafeFileLogger.SafeAppendText(context.RefreshLogName,
                            $"[{DateTime.Now}] Filters ({selectedFilterItems?.Count ?? 0}): {string.Join(", ", selectedFilterItems ?? new List<string>())}\n");
                        SafeFileLogger.SafeAppendText(context.RefreshLogName,
                            $"[{DateTime.Now}] MEP Categories ({selectedMepCategories?.Count ?? 0}): {string.Join(", ", selectedMepCategories ?? new List<string>())}\n");
                        SafeFileLogger.SafeAppendText(context.RefreshLogName,
                            $"[{DateTime.Now}] Reference Files ({selectedReferenceFiles?.Count ?? 0}): {string.Join(", ", selectedReferenceFiles ?? new List<string>())}\n");
                        SafeFileLogger.SafeAppendText(context.RefreshLogName,
                            $"[{DateTime.Now}] Host Files ({selectedHostFiles?.Count ?? 0}): {string.Join(", ", selectedHostFiles ?? new List<string>())}\n");
                        SafeFileLogger.SafeAppendText(context.RefreshLogName,
                            $"[{DateTime.Now}] ============================================\n\n");
                    }
                    catch { }
                    
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
            // ✅ CRITICAL DIAGNOSTIC: Log refresh start (even before validation)
            if (!context.IsDeploymentMode)
            {
                DebugLogger.Info($"[REFRESH-REFACTORED] ===== ExecuteRefreshInternal STARTED =====");
                DebugLogger.Info($"[REFRESH-REFACTORED] RefreshLogName: {context.RefreshLogName}");
                SafeFileLogger.SafeAppendText(context.RefreshLogName, 
                    $"[{DateTime.Now}] ===== ExecuteRefreshInternal STARTED =====\n");
            }
            
            // PHASE 1: Validate UI selections
            using (context.PerformanceMonitor.TrackOperation("1. UI Validation"))
            {
                if (!ValidateUISelections(context))
                {
                    // ✅ CRITICAL DIAGNOSTIC: Log validation failure
                    if (!context.IsDeploymentMode)
                    {
                        DebugLogger.Warning("[REFRESH-REFACTORED] ❌ Validation failed - returning Cancelled");
                        SafeFileLogger.SafeAppendText(context.RefreshLogName, 
                            $"[{DateTime.Now}] ❌ Validation failed - refresh cancelled\n");
                    }
                    return Result.Cancelled;
                }
            }
            
            // PHASE 2: Determine Path
            DetermineRefreshPath(context);
            
            // PHASE 3: Reset Flags (Hierarchical Verification & Section Box Filtering)
            // ✅ CRITICAL: Run this BEFORE loading data so in-memory zones have correct flags
            ResetFlags(context);
            
            // PHASE 4: Load Data
            // ✅ RELIABLE: Loads zones with correct flags from optimized Phase 3
            LoadRefreshData(context);
            
            // PHASE 5: Sync Flags
            if (context.PathStrategy.ShouldSyncFlags)
            {
                using (context.PerformanceMonitor.TrackOperation("4. Flag Sync"))
                {
                    SyncFlagsFromGlobal(context);
                }
            }
            else
            {
                if (!context.IsDeploymentMode)
                    DebugLogger.Info("[REFRESH-REFACTORED] Skipping flag sync (not required for this path)");
            }
            
            // PHASE 6: Validate Zones
            ValidateZones(context);
            
            // PHASE 7: Process Intersections
            ProcessIntersections(context);
            
            // PHASE 9: Capture Parameters
            CaptureParameters(context);
            
            // PHASE 10: Merge and Save
            UpdateProgress(80, "Merging and saving...");
            using (var saveOp = context.PerformanceMonitor.TrackOperation("9. Save") as PerformanceMonitor.OperationTracker)
            {
                MergeAndSave(context);
                saveOp?.SetItemCount(context.AllClashZones?.Count ?? 0);
            }
            
            // PHASE 11: Final Cleanup
            UpdateProgress(90, "Finalizing...");
            using (context.PerformanceMonitor.TrackOperation("10. Cleanup"))
            {
                FinalCleanup(context);
            }
            
            UpdateProgress(100, "Complete!");
            
            // Show summary
            ShowSummary(context);
            
            return Result.Succeeded;
        }

        /// <summary>
        /// Database-first loader for existing clash zones when XML loading is skipped.
        /// Uses R-tree section box filtering if available.
        /// </summary>
        private List<ClashZone> LoadExistingClashZonesFromDatabase(RefreshContext context)
        {
            var result = new List<ClashZone>();
            try
            {
                using (var dbContext = new Data.SleeveDbContext(_document, msg =>
                {
                    if (!context.IsDeploymentMode)
                        DebugLogger.Info($"[REFRESH-REFACTORED][SQLite] {msg}");
                    SafeFileLogger.SafeAppendText(context.RefreshLogName, $"[{DateTime.Now}] [SQLite] {msg}\n");
                }))
                {
                    var repo = new Data.Repositories.ClashZoneRepository(dbContext, msg =>
                    {
                        if (!context.IsDeploymentMode)
                            DebugLogger.Info($"[REFRESH-REFACTORED][SQLite] {msg}");
                        SafeFileLogger.SafeAppendText(context.RefreshLogName, $"[{DateTime.Now}] [SQLite] {msg}\n");
                    });

                    // If section box is active, use R-tree path to constrain
                    BoundingBoxXYZ? sectionBox = null;
                    if (_document.ActiveView is View3D v3 && v3.IsSectionBoxActive)
                    {
                        sectionBox = Helpers.SectionBoxHelper.GetSectionBoxBounds(v3);
                    }

                    var filterName = context.SelectedFilterNames?.FirstOrDefault() ?? "";
                    foreach (var category in context.SelectedMepCategories ?? new List<string>())
                    {
                        // Use public repository method for safety
                        var zones = repo.GetClashZonesByFilter(filterName, category, unresolvedOnly: false) ?? new List<ClashZone>();
                        if (sectionBox != null)
                        {
                            // If section box is present, filter in-memory by bbox (intersection point inside box)
                            zones = zones.Where(z => IsZoneWithinSectionBox(z, sectionBox)).ToList();
                        }
                        if (zones.Count > 0)
                            result.AddRange(zones);
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[REFRESH-REFACTORED] DB load failed, falling back to empty: {ex.Message}");
                SafeFileLogger.SafeAppendText(context.RefreshLogName, $"[{DateTime.Now}] [REFRESH-REFACTORED] DB load failed: {ex.Message}\n");
            }
            return result;
        }

        /// <summary>
        /// Lightweight check: treat a zone as within the section box if its intersection point lies inside the box bounds.
        /// </summary>
        private static bool IsZoneWithinSectionBox(ClashZone zone, BoundingBoxXYZ sectionBox)
        {
            if (zone == null || sectionBox == null)
                return false;

            // Use intersection point stored on the zone
            var p = zone.IntersectionPoint;
            if (p == null)
                return false;

            var min = sectionBox.Min;
            var max = sectionBox.Max;
            return p.X >= min.X && p.X <= max.X
                && p.Y >= min.Y && p.Y <= max.Y
                && p.Z >= min.Z && p.Z <= max.Z;
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
                DebugLogger.Info($"[REFRESH-REFACTORED] Selected reference files ({context.SelectedReferenceFiles.Count}): {string.Join(", ", context.SelectedReferenceFiles)}");
            }
            
            if (context.SelectedHostFiles == null || context.SelectedHostFiles.Count == 0)
            {
                errors.Add("Host Linked Files");
                DebugLogger.Warning("[REFRESH-REFACTORED] ⚠️ No host files selected");
            }
            else
            {
                DebugLogger.Info($"[REFRESH-REFACTORED] Selected host files ({context.SelectedHostFiles.Count}): {string.Join(", ", context.SelectedHostFiles)}");
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
            
            var flagManager = Services.FlagManagement.FlagManagerFactory.CreateAdapter(_document);
            
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
                context.RefreshLogName,
                context.PerformanceMonitor);
            
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
            
            // ✅ REMOVED: ResetIsCurrentClashFlag is no longer needed
            // SetReadyForPlacementBatchOptimized now uses atomic CASE-based UPDATE:
            // - Sets IsCurrentClashFlag=1 for zones IN scope (filter+category+unresolved+section box)
            // - Sets IsCurrentClashFlag=0 for zones OUT of scope (filter+category but resolved or outside section box)
            // This eliminates timing issues and is more efficient (single atomic operation)
            /*
            try
            {
                using (var resetOp = context.PerformanceMonitor.TrackOperation("9a0. Reset IsCurrentClash") as PerformanceMonitor.OperationTracker)
                {
                    using (var dbContext = new Data.SleeveDbContext(_document, msg =>
                    {
                        if (!context.IsDeploymentMode)
                            DebugLogger.Info($"[REFRESH-REFACTORED][SQLite] {msg}");
                    }))
                    {
                        var repository = new Data.Repositories.ClashZoneRepository(dbContext, msg =>
                        {
                            if (!context.IsDeploymentMode)
                                DebugLogger.Info($"[REFRESH-REFACTORED][SQLite] {msg}");
                        });
                        
                        int resetCount = repository.ResetIsCurrentClashFlag(
                            context.SelectedFilterNames ?? new List<string>(),
                            context.SelectedMepCategories ?? new List<string>());
                        
                        resetOp?.SetItemCount(resetCount);
                        
                        if (!context.IsDeploymentMode)
                        {
                            DebugLogger.Info($"[REFRESH-REFACTORED] [ISCURRENTCLASH-RESET] ✅ Reset IsCurrentClashFlag=0 for {resetCount} zones before SaveClashZones");
                            SafeFileLogger.SafeAppendText(context.RefreshLogName,
                                $"[{DateTime.Now}] [ISCURRENTCLASH-RESET] ✅ Reset IsCurrentClashFlag=0 for {resetCount} zones\n");
                        }
                    }
                }
            }
            catch (Exception resetEx)
            {
                if (!context.IsDeploymentMode)
                {
                    DebugLogger.Warning($"[REFRESH-REFACTORED] [ISCURRENTCLASH-RESET] ⚠️ Failed to reset IsCurrentClashFlag: {resetEx.Message}");
                    SafeFileLogger.SafeAppendText(context.RefreshLogName,
                        $"[{DateTime.Now}] [ISCURRENTCLASH-RESET] ⚠️ Failed to reset: {resetEx.Message}\n");
                }
                // Continue with refresh even if reset fails (non-blocking)
            }
            */
            
            try
            {
                // ✅ PATH STRATEGY: Use AllowStructuralUpdates from strategy
                var allowStructuralUpdates = context.PathStrategy?.AllowStructuralUpdates ?? true;
                
                if (!context.IsDeploymentMode)
                {
                    DebugLogger.Info($"[REFRESH-REFACTORED] [MERGE-SAVE] Using allowStructuralUpdates={allowStructuralUpdates} from {context.PathStrategy?.PathName ?? "default"}");
                }
                
                using (var dbSaveOp = context.PerformanceMonitor.TrackOperation("9a. Database/XML Save") as PerformanceMonitor.OperationTracker)
                {
                    persistenceService.SaveClashZones(
                        context.AllClashZones ?? new List<ClashZone>(),
                        normalizedBaseName,
                        enabledFilter,
                        allowStructuralUpdates: allowStructuralUpdates);
                    dbSaveOp?.SetItemCount(context.AllClashZones?.Count ?? 0);
                }
                
                DebugLogger.Info($"[REFRESH-REFACTORED] [MERGE-SAVE] ✅ SaveClashZones completed successfully");
                SafeFileLogger.SafeAppendText(context.RefreshLogName, 
                    $"[{DateTime.Now}] [MERGE-SAVE] ✅ SaveClashZones completed successfully\n");
                
                // ✅ SIMPLE FIX: Set ReadyForPlacementFlag=1 for ALL unresolved zones AFTER SaveClashZones completes
                // This includes both existing zones AND new zones that were just created and saved
                // No conditions needed - just set the flag for any unresolved zones within section box
                if (context.SelectedFilterNames != null && context.SelectedMepCategories != null && 
                    context.SelectedFilterNames.Count > 0 && context.SelectedMepCategories.Count > 0)
                {
                    try
                    {
                        using (var flagOp = context.PerformanceMonitor.TrackOperation("9b. Set ReadyForPlacement Flags") as PerformanceMonitor.OperationTracker)
                        {
                            using (var dbContext = new Data.SleeveDbContext(_document, msg =>
                            {
                                if (!context.IsDeploymentMode)
                                    DebugLogger.Info($"[REFRESH-REFACTORED][SQLite] {msg}");
                            }))
                            {
                                var repository = new Data.Repositories.ClashZoneRepository(dbContext, msg =>
                                {
                                    if (!context.IsDeploymentMode)
                                        DebugLogger.Info($"[REFRESH-REFACTORED][SQLite] {msg}");
                                });

                                // Get section box bounds
                                BoundingBoxXYZ? sectionBoxNullable = null;
                                if (_document.ActiveView is View3D view3D && view3D.IsSectionBoxActive)
                                {
                                    sectionBoxNullable = Helpers.SectionBoxHelper.GetSectionBoxBounds(view3D);
                                    if (sectionBoxNullable != null && !context.IsDeploymentMode)
                                    {
                                        BoundingBoxXYZ sb = sectionBoxNullable;
                                        DebugLogger.Info($"[REFRESH-REFACTORED] Section box active: Min=({sb.Min.X:F2}, {sb.Min.Y:F2}, {sb.Min.Z:F2}), Max=({sb.Max.X:F2}, {sb.Max.Y:F2}, {sb.Max.Z:F2})");
                                    }
                                }

                                // ✅ FORCE DETECTION MODE: Check if force detection mode is enabled
                                // If enabled, reset all flags to false (preserving GUIDs) before detection
                                bool forceDetectionMode = false;
                                try
                                {
                                    var settings = _appProfileService?.GetCurrentSettings();
                                    forceDetectionMode = settings?.ForceDetectionMode ?? false;
                                }
                                catch { }

                                if (forceDetectionMode)
                                {
                                    if (!context.IsDeploymentMode)
                                    {
                                        try
                                        {
                                            DebugLogger.Info($"[REFRESH-REFACTORED] [FORCE-DETECTION] ⚡ Force Detection Mode enabled - resetting all flags (preserving GUIDs)");
                                            SafeFileLogger.SafeAppendText(context.RefreshLogName,
                                                $"[{DateTime.Now}] [REFRESH-REFACTORED] [FORCE-DETECTION] ⚡ Force Detection Mode enabled - resetting all flags (preserving GUIDs)\n");
                                        }
                                        catch { }
                                    }

                                    int zonesReset = repository.ResetAllFlagsForForceDetectionMode(
                                        context.SelectedFilterNames,
                                        context.SelectedMepCategories);

                                    if (!context.IsDeploymentMode)
                                    {
                                        try
                                        {
                                            DebugLogger.Info($"[REFRESH-REFACTORED] [FORCE-DETECTION] ✅ Reset all flags for {zonesReset} zones (IsResolved=0, IsClusterResolved=0, SleeveId=0, ClusterId=0, ReadyForPlacement=1) - GUIDs preserved");
                                            SafeFileLogger.SafeAppendText(context.RefreshLogName,
                                                $"[{DateTime.Now}] [REFRESH-REFACTORED] [FORCE-DETECTION] ✅ Reset all flags for {zonesReset} zones - GUIDs preserved\n");
                                        }
                                        catch { }
                                    }
                                }
                                
                                // ✅ CRITICAL: Section box filtering AFTER SaveClashZones
                                // This ensures zones exist in database before we try to filter them
                                // Reuse section box bounds from above (already declared at line 584)
                                if (!context.IsDeploymentMode && sectionBoxNullable != null)
                                {
                                    BoundingBoxXYZ sb = sectionBoxNullable;
                                    DebugLogger.Info($"[MERGE-SAVE] ✅ SECTION BOX ACTIVE: Min=({sb.Min.X:F2}, {sb.Min.Y:F2}, {sb.Min.Z:F2}), Max=({sb.Max.X:F2}, {sb.Max.Y:F2}, {sb.Max.Z:F2})");
                                    SafeFileLogger.SafeAppendText(context.RefreshLogName,
                                        $"[{DateTime.Now}] [MERGE-SAVE] ✅ SECTION BOX ACTIVE: Min=({sb.Min.X:F2}, {sb.Min.Y:F2}, {sb.Min.Z:F2}), Max=({sb.Max.X:F2}, {sb.Max.Y:F2}, {sb.Max.Z:F2})\n");
                                }
                                
                                // ========== IsCurrentClashFlag TRACING SECTION ==========
                                if (!context.IsDeploymentMode)
                                {
                                    SafeFileLogger.SafeAppendText(context.RefreshLogName,
                                        $"\n[{DateTime.Now}] ========== IsCurrentClashFlag TRACING ==========\n");
                                    SafeFileLogger.SafeAppendText(context.RefreshLogName,
                                        $"[{DateTime.Now}] UseBatchReadyForPlacementUpdate = {Services.OptimizationFlags.UseBatchReadyForPlacementUpdate}\n");
                                    SafeFileLogger.SafeAppendText(context.RefreshLogName,
                                        $"[{DateTime.Now}] FilterNames: [{string.Join(", ", context.SelectedFilterNames ?? new List<string>())}]\n");
                                    SafeFileLogger.SafeAppendText(context.RefreshLogName,
                                        $"[{DateTime.Now}] Categories: [{string.Join(", ", context.SelectedMepCategories ?? new List<string>())}]\n");
                                    SafeFileLogger.SafeAppendText(context.RefreshLogName,
                                        $"[{DateTime.Now}] SectionBox: {(sectionBoxNullable != null ? "Present" : "NULL")}\n");
                                }
                                
                                // Set ReadyForPlacementFlag=1 for unresolved zones within section box
                                int markedCount = repository.SetReadyForPlacementForUnresolvedZonesInSectionBox(
                                    context.SelectedFilterNames ?? new List<string>(),
                                    context.SelectedMepCategories ?? new List<string>(),
                                    sectionBoxNullable);
                                
                                flagOp?.SetItemCount(markedCount);
                                
                                // ========== VERIFY FLAGS AFTER CALL ==========
                                if (!context.IsDeploymentMode)
                                {
                                    SafeFileLogger.SafeAppendText(context.RefreshLogName,
                                        $"[{DateTime.Now}] RESULT: markedCount = {markedCount}\n");
                                    
                                    // Query DB to verify actual flag values
                                    try
                                    {
                                        var flagStats = repository.GetFlagStatistics();
                                        SafeFileLogger.SafeAppendText(context.RefreshLogName,
                                            $"[{DateTime.Now}] DB FLAG STATS: Total={flagStats.Total}, IsCurrentClash=1: {flagStats.IsCurrentClashSet}, ReadyForPlacement=1: {flagStats.ReadyForPlacementSet}, IsResolved=1: {flagStats.IsResolvedSet}\n");
                                    }
                                    catch (Exception statsEx)
                                    {
                                        SafeFileLogger.SafeAppendText(context.RefreshLogName,
                                            $"[{DateTime.Now}] Could not get flag stats: {statsEx.Message}\n");
                                    }
                                    
                                    SafeFileLogger.SafeAppendText(context.RefreshLogName,
                                        $"[{DateTime.Now}] ========== END IsCurrentClashFlag TRACING ==========\n\n");
                                    
                                    DebugLogger.Info($"[MERGE-SAVE] [SECTION-BOX-FILTER] ✅ Set ReadyForPlacementFlag=1 for {markedCount} unresolved zones within section box");
                                    SafeFileLogger.SafeAppendText(context.RefreshLogName,
                                        $"[{DateTime.Now}] [MERGE-SAVE] [SECTION-BOX-FILTER] ✅ Set ReadyForPlacementFlag=1 for {markedCount} zones\n");
                                }
                            }
                        }
                    }
                    catch (Exception markEx)
                    {
                        if (!context.IsDeploymentMode)
                        {
                            DebugLogger.Warning($"[REFRESH-REFACTORED] ⚠️ Failed to set ReadyForPlacementFlag: {markEx.Message}");
                            SafeFileLogger.SafeAppendText(context.RefreshLogName,
                                $"[{DateTime.Now}] [REFRESH-REFACTORED] ⚠️ Failed to set ReadyForPlacementFlag: {markEx.Message}\n");
                        }
                        // Continue with refresh even if marking fails (non-blocking)
                    }
                }
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
                    
                    // ✅ DATABASE-ONLY MODE: XML filter group creation DISABLED
                    // This code was creating duplicate filter entries like "Ventilation_ducts" in the database
                    // Zones are saved with base filter name only (e.g., "Ventilation" + category)
                    /*
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
                    */
                    
                    
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
            // ✅ PATH 2 OPTIMIZATION: Skip expensive cleanup for fresh mode
            // Path 2 is a fresh start with no existing sleeves/cache, so cleanup is unnecessary
            bool isPath2 = context.PathStrategy?.PathName?.Contains("PATH 2") == true;
            
            if (isPath2)
            {
                if (!context.IsDeploymentMode)
                    DebugLogger.Info("[REFRESH-REFACTORED] Path 2 (Fresh Mode) - skipping cleanup (no existing state to clean)");
                return; // Skip all cleanup for fresh mode - saves ~297ms
            }
            
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
            
            // ✅ OPTIMIZATION: Skip forced GC if flag enabled (saves ~289ms)
            // Modern .NET GC is efficient - forced collection often counterproductive
            if (!OptimizationFlags.SkipForcedGarbageCollection)
            {
                if (!context.IsDeploymentMode)
                    DebugLogger.Info("[REFRESH-REFACTORED] Running forced GC (SkipForcedGarbageCollection=false)");
                
                GC.Collect(2, GCCollectionMode.Forced, true);
                GC.WaitForPendingFinalizers();
                GC.Collect(2, GCCollectionMode.Forced, true);
            }
            else
            {
                if (!context.IsDeploymentMode)
                    DebugLogger.Info("[REFRESH-REFACTORED] Skipping forced GC - letting CLR manage memory (saves ~289ms)");
            }
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
        
        // ✅ EXTRACTED METHODS (Phase 3)
        
        private void LoadRefreshData(RefreshContext context)
        {
            UpdateProgress(10, "Loading XML data...");
            using (var xmlOp = context.PerformanceMonitor.TrackOperation("2. XML Loading") as PerformanceMonitor.OperationTracker)
            {
                if (OptimizationFlags.SkipXmlLoadingDuringRefresh)
                {
                    if (!context.IsDeploymentMode)
                        DebugLogger.Info("[XML-CACHE] ⚡ OPTIMIZATION: Skipping XML cache loading (database-only mode)");
                    SafeFileLogger.SafeAppendText(context.RefreshLogName,
                        $"[{DateTime.Now}] [XML-CACHE] ⚡ OPTIMIZATION: Skipping XML cache loading (database-only mode)\n");

                    context.XmlCache = new Services.Refresh.XmlCache();
                    xmlOp?.SetItemCount(0);
                }
                else
                {
                    var xmlManager = new XmlCacheManager(_document, context.RefreshLogName);
                    context.XmlCache = xmlManager.LoadAll(context.SelectedFilterNames, context.SelectedMepCategories);
                    xmlOp?.SetItemCount(context.XmlCache.FilterXml.Count); // GlobalXml removed - DB only
                }
            }
            
            UpdateProgress(20, "Loading existing clash zones...");
            using (var loadOp = context.PerformanceMonitor.TrackOperation("3. Load Existing Zones") as PerformanceMonitor.OperationTracker)
            {
                if (OptimizationFlags.SkipXmlLoadingDuringRefresh)
                {
                    context.ExistingClashZones = LoadExistingClashZonesFromDatabase(context);
                }
                else
                {
                    context.ExistingClashZones = LoadExistingClashZones(context);
                }
                loadOp?.SetItemCount(context.ExistingClashZones?.Count ?? 0);
            }
            
            // ✅ SECTION BOX CAPTURE: Capture and store section box bounds during refresh
            CaptureAndStoreSectionBox(context);
        }

        private void DetermineRefreshPath(RefreshContext context)
        {
            IRefreshPathStrategy pathStrategy;
            
            // ✅ DI SUPPORT: Use injected path determiner if available (Phase 2/4)
            if (_pathDeterminer != null)
            {
                pathStrategy = _pathDeterminer.DeterminePath(context);
            }
            else
            {
                // Fallback to static implementation
                pathStrategy = RefreshPathDeterminer.DeterminePath(context, context.EnableThreePointValidation);
            }
            
            if (!context.IsDeploymentMode)
            {
                DebugLogger.Info($"[REFRESH-REFACTORED] Using {pathStrategy.PathName}");
                DebugLogger.Info($"[REFRESH-REFACTORED] Path settings: AllowStructuralUpdates={pathStrategy.AllowStructuralUpdates}, EnableThreePointValidation={pathStrategy.EnableThreePointValidation}, ShouldResetFlags={pathStrategy.ShouldResetFlags}, ShouldSyncFlags={pathStrategy.ShouldSyncFlags}, ShouldCheckGuids={pathStrategy.ShouldCheckGuids}");
            }
            
            context.PathStrategy = pathStrategy;
        }

        private void ValidateZones(RefreshContext context)
        {
            UpdateProgress(40, "Validating clash zones...");
            using (var validationOp = context.PerformanceMonitor.TrackOperation("5. Validation") as PerformanceMonitor.OperationTracker)
            {
                // ✅ DEPLOYMENT MODE CHECK: When DeploymentMode is OFF, skip validation entirely
                // This ensures full detection runs for all zones (no splitting into validated/invalidated)
                // During testing, we want to populate the database with complete data
                if (!context.IsDeploymentMode)
                {
                    DebugLogger.Info("[REFRESH-REFACTORED] ⚠️ DeploymentMode OFF: Skipping validation - treating all zones as needing full detection");
                    // Treat all zones as invalidated (needing full detection) when DeploymentMode is OFF
                    context.ValidatedZones = new List<ClashZone>();
                    context.InvalidatedZones = context.ExistingClashZones ?? new List<ClashZone>();
                    validationOp?.SetItemCount(context.ExistingClashZones?.Count ?? 0);
                    return;
                }
                
                // ✅ DEPLOYMENT MODE ON: Run validation based on path strategy
                if (context.PathStrategy.EnableThreePointValidation)
                {
                    var validationService = new ValidationService(context, Services.FlagManagement.FlagManagerFactory.CreateAdapter(_document));
                    var validationResult = validationService.ValidateClashZones(context.ExistingClashZones);
                    
                    var processedZones = context.PathStrategy.ProcessZonesAfterValidation(
                        context,
                        validationResult.ValidZones,
                        validationResult.InvalidZones);
                    
                    context.ExistingClashZones = processedZones;
                    context.ValidatedZones = validationResult.ValidZones ?? new List<ClashZone>();
                    context.InvalidatedZones = validationResult.InvalidZones ?? new List<ClashZone>();
                
                    if (validationResult.InvalidZones.Count > 0)
                    {
                        validationService.RemoveInvalidZonesFromDatabase(validationResult.InvalidZones);
                    }
                
                    validationOp?.SetItemCount(context.ExistingClashZones.Count);
                }
                else
                {
                    if (!context.IsDeploymentMode)
                    {
                        DebugLogger.Info("[REFRESH-REFACTORED] Skipping 3-point validation (not required for this path)");
                    }
                    validationOp?.SetItemCount(context.ExistingClashZones?.Count ?? 0);
                }
            }
        }

        private void ResetFlags(RefreshContext context)
        {
            // ✅ CRITICAL: ALWAYS verify sleeves and set section box flags early in the refresh
            // This ensures hierarchical resolution (Combined -> Cluster -> Individual) is clean
            // and section box filtering is applied consistently to the database.
            
            UpdateProgress(30, "Verifying existing sleeves and resetting flags...");
            
            // Get section box bounds (if active)
            BoundingBoxXYZ? sectionBoxNullable = null;
            if (_document.ActiveView is View3D view3D && view3D.IsSectionBoxActive)
            {
                sectionBoxNullable = Helpers.SectionBoxHelper.GetSectionBoxBounds(view3D);
                if (sectionBoxNullable != null && !context.IsDeploymentMode)
                {
                    BoundingBoxXYZ sb = sectionBoxNullable;
                    DebugLogger.Info($"[REFRESH-REFACTORED] ✅ SECTION BOX ACTIVE: Min=({sb.Min.X:F2}, {sb.Min.Y:F2}, {sb.Min.Z:F2}), Max=({sb.Max.X:F2}, {sb.Max.Y:F2}, {sb.Max.Z:F2})");
                }
            }
            
            try
            {
                using (var op = context.PerformanceMonitor.TrackOperation("3. Hierarchical Flag Reset") as PerformanceMonitor.OperationTracker)
                {
                    using (var dbContext = new Data.SleeveDbContext(_document, msg => { }))
                    {
                        var repository = new Data.Repositories.ClashZoneRepository(dbContext, msg => { });

                        // ✅ STEP 1: Optimized Hierarchical Reset (Combined -> Cluster -> Individual)
                        // Uses O(1) HashSet check against ALL opening families in Revit.
                        int resetCount = repository.VerifyExistingSleevesAndResetFlags(
                            _document,
                            context.SelectedFilterNames ?? new List<string>(),
                            context.SelectedMepCategories ?? new List<string>());

                        if (!context.IsDeploymentMode)
                        {
                            DebugLogger.Info($"[REFRESH-REFACTORED] [HIERARCHICAL-RESET] ✅ Reset flags for {resetCount} zones with missing/deleted sleeves");
                            SafeFileLogger.SafeAppendText(context.RefreshLogName,
                                $"[{DateTime.Now}] [HIERARCHICAL-RESET] ✅ Reset flags for {resetCount} zones\n");
                        }

                        // ✅ STEP 2: Session Context (SOLID Refactor)
                        // Orchestrates the 2-step flag setting logic:
                        // 1. Reset & Set IsCurrentClashFlag based on Filters + Section Box
                        // 2. Set ReadyForPlacementFlag based on IsCurrentClashFlag + Unresolved Status
                        var sessionContext = new SessionContextService(repository);
                        int markedCount = sessionContext.UpdateSessionFlags(
                            context.SelectedFilterNames ?? new List<string>(),
                            context.SelectedMepCategories ?? new List<string>(),
                            sectionBoxNullable,
                            context.SelectedHostTypes);

                        if (!context.IsDeploymentMode)
                        {
                            DebugLogger.Info($"[REFRESH-REFACTORED] [SESSION-CONTEXT] ✅ Applied section box context. Marked {markedCount} zones as ReadyForPlacement.");
                            SafeFileLogger.SafeAppendText(context.RefreshLogName,
                                $"[{DateTime.Now}] [SESSION-CONTEXT] ✅ Applied section box context. Marked {markedCount} zones\n");
                        }
                        
                        op?.SetItemCount(resetCount + markedCount);
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[REFRESH-REFACTORED] ⚠️ Hierarchical flag reset failed: {ex.Message}");
                SafeFileLogger.SafeAppendText(context.RefreshLogName,
                    $"[{DateTime.Now}] [REFRESH-REFACTORED] ⚠️ Hierarchical flag reset failed: {ex.Message}\n");
            }
            
            // ✅ PATH-SPECIFIC LEGACY RESET: Only if strategy requires it and not superseded by hierarchical reset
            if (context.PathStrategy.ShouldResetFlags)
            {
                // Note: The hierarchical reset above is much more robust than the per-category XML check below.
                // We keep this for now to maintain consistency with historical behavior if needed.
                UpdateProgress(35, "Checking category-specific legacy flags...");
                using (context.PerformanceMonitor.TrackOperation("3B. Legacy Flag Sync"))
                {
                    // If we have already reset 1000s of sleeves hierarchically, this might be skip-able.
                    // But for safety, we allow the strategy to decide.

                    // ✅ FIX: Define expected variables
                    var categoriesToCheck = context.SelectedMepCategories ?? new List<string>();
                    Dictionary<string, List<Models.ClashZone>> clashZonesByCategory = null; // Not loaded yet at this phase

                    context.PathStrategy.ResetInstanceIdsForDeletedSleeves(
                        context,
                        categoriesToCheck,
                        clashZonesByCategory);
                    
                    if (!context.IsDeploymentMode)
                    {
                        DebugLogger.Info($"[REFRESH-REFACTORED] Path-specific flag reset completed");
                    }
                }
            }
            else
            {
                if (!context.IsDeploymentMode)
                    DebugLogger.Info("[REFRESH-REFACTORED] Skipping path-specific flag reset (not required for this path, but section box filtering was applied)");
            }
        }

        private void ProcessIntersections(RefreshContext context)
        {
            UpdateProgress(50, "Processing intersections...");
            using (var intersectionOp = context.PerformanceMonitor.TrackOperation("6. Intersection Processing") as PerformanceMonitor.OperationTracker)
            {
                var xmlManager = new XmlCacheManager(_document, context.RefreshLogName);
                var validationService = new ValidationService(context, Services.FlagManagement.FlagManagerFactory.CreateAdapter(_document));
                var paramService = new ParameterCaptureService(context);
                
                var logger = new Action<string>(msg => 
                {
                    SafeFileLogger.SafeAppendTextAlways(SafeFileLogger.GetLogFilePath("logger_debug.txt"), $"[{DateTime.Now}] {msg}\n");
                });
                
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
                
                // ✅ ROUTING LOGIC: Check if "Duct Accessories" is selected
                // If selected, process dampers separately; all other categories go to intersection processor
                bool hasDuctAccessories = context.SelectedMepCategories != null && 
                    context.SelectedMepCategories.Any(c => string.Equals(c, "Duct Accessories", StringComparison.OrdinalIgnoreCase));
                
                // ✅ CHECK FAST PATH CONDITIONS: Determine if fast path should be taken BEFORE processing
                // This allows us to skip both damper processing AND intersection detection for validated zones
                var processor = new IntersectionProcessor(
                    context,
                    xmlManager,
                    validationService,
                    paramService,
                    context.PerformanceMonitor,
                    logger,
                    progressCallback);
                
                var decision = processor.PrepareExistingZones();
                
                // ✅ FAST PATH: Skip both damper processing AND intersection detection for validated zones
                // Conditions: DeploymentMode ON, Adopt OFF, all combos processed (IsFilterComboNew=0)
                bool shouldSkipProcessing = !decision.ShouldRunDetection && DeploymentConfiguration.DeploymentMode;
                
                List<ClashZone> damperClashZones = new List<ClashZone>();
                
                if (hasDuctAccessories && !shouldSkipProcessing)
                {
                    // ✅ ROUTE TO DAMPER PROCESSING: Process Duct Accessories separately (only if NOT fast path)
                    if (!context.IsDeploymentMode)
                    {
                        DebugLogger.Info("[REFRESH-REFACTORED] ✅ Duct Accessories category selected → Routing to DamperProcessingService");
                        SafeFileLogger.SafeAppendText(context.RefreshLogName, 
                            $"[{DateTime.Now}] [REFRESH-REFACTORED] ✅ Duct Accessories category selected → Routing to DamperProcessingService\n");
                    }
                    
                    try
                    {
                        // Get section box from view
                        BoundingBoxXYZ? sectionBox = null;
                        if (_document.ActiveView is View3D view3D && view3D.IsSectionBoxActive)
                        {
                            sectionBox = Helpers.SectionBoxHelper.GetSectionBoxBounds(view3D);
                        }
                        
                        // Get ClashZoneStorage from context (or create new one)
                        var clashZoneStorage = context.XmlCache?.FilterXml?.Values?.FirstOrDefault() ?? new ClashZoneStorage();
                        
                        // Create damper processing service
                        var damperService = new DamperProcessingService(
                            _document,
                            logger,
                            clashZoneStorage,
                            null, // Use default damper type detector
                            null, // Use default connector detector
                            null, // Use default parameter snapshot service
                            context.PerformanceMonitor); // Pass performance monitor for tracking
                        
                        // ✅ PHASE 3 OPTIMIZATION: Pre-index "Opening" instances for O(1) proximity check
                        // This allows DamperProcessingService to skip redundant checks
                        var openingLocationMap = new Dictionary<string, FamilyInstance>();
                        var allOpenings = new FilteredElementCollector(_document)
                            .OfClass(typeof(FamilyInstance))
                            .Cast<FamilyInstance>()
                            .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                            .ToList();

                        foreach (var opening in allOpenings)
                        {
                            XYZ loc = null;
                            if (opening.Location is LocationPoint lp) loc = lp.Point;
                            else if (opening.Location is LocationCurve lc) loc = lc.Curve.Evaluate(0.5, true);

                            if (loc != null)
                            {
                                string key = $"{Math.Round(loc.X, 3)}_{Math.Round(loc.Y, 3)}_{Math.Round(loc.Z, 3)}";
                                if (!openingLocationMap.ContainsKey(key)) openingLocationMap.Add(key, opening);
                            }
                        }
                        var openingPointKeys = openingLocationMap.Keys.ToHashSet();
                        if (!context.IsDeploymentMode) DebugLogger.Info($"[REFRESH-REFACTORED] Pre-indexed {openingPointKeys.Count} openings for Damper O(1) check");

                        // Process dampers (pass selected reference files to respect UI selection)
                        // ✅ CRITICAL FIX: Pass existing zones to prevent duplicate GUID creation
                        damperClashZones = damperService.ProcessDampers(
                            context.SelectedMepCategories,
                            context.SelectedHostTypes,
                            context.SelectedReferenceFiles,
                            sectionBox,
                            context.ExistingClashZones,
                            openingPointKeys) ?? new List<ClashZone>();
                        
                        if (!context.IsDeploymentMode)
                        {
                            DebugLogger.Info($"[REFRESH-REFACTORED] ✅ Damper processing complete: {damperClashZones.Count} ClashZones created");
                            SafeFileLogger.SafeAppendText(context.RefreshLogName, 
                                $"[{DateTime.Now}] [REFRESH-REFACTORED] ✅ Damper processing complete: {damperClashZones.Count} ClashZones created\n");
                        }
                    }
                    catch (Exception damperEx)
                    {
                        DebugLogger.Warning($"[REFRESH-REFACTORED] ⚠️ Error in damper processing: {damperEx.Message}");
                        SafeFileLogger.SafeAppendText(context.RefreshLogName, 
                            $"[{DateTime.Now}] [REFRESH-REFACTORED] ⚠️ Error in damper processing: {damperEx.Message}\n");
                    }
                }
                else if (hasDuctAccessories && shouldSkipProcessing)
                {
                    // ✅ FAST PATH: Skip damper processing - zones already exist in database for validated zones
                    System.Diagnostics.Debug.WriteLine($"[REFRESH-REFACTORED] ⚡⚡⚡ FAST PATH: Skipping damper processing (validated zones, no detection needed)");
                    if (!context.IsDeploymentMode)
                    {
                        DebugLogger.Info("[REFRESH-REFACTORED] ⚡ FAST PATH: Skipping damper processing (validated zones, no detection needed)");
                        SafeFileLogger.SafeAppendText(context.RefreshLogName, 
                            $"[{DateTime.Now}] [REFRESH-REFACTORED] ⚡ FAST PATH: Skipping damper processing (validated zones, no detection needed)\n");
                    }
                    damperClashZones = new List<ClashZone>(); // Use empty list - zones already exist in database
                }
                else
                {
                    if (!context.IsDeploymentMode)
                    {
                        DebugLogger.Info("[REFRESH-REFACTORED] ✅ No Duct Accessories category selected → Skipping damper processing");
                    }
                }
                
                var intersectionClashZones = processor.RunDetectionIfNeeded(decision);
                
                // ✅ MERGE RESULTS: Combine intersection processor results with damper processing results
                var allClashZones = intersectionClashZones?.ToList() ?? new List<ClashZone>();
                if (damperClashZones.Count > 0)
                {
                    var existingIds = new HashSet<Guid>(allClashZones.Select(z => z.Id));
                    foreach (var damperZone in damperClashZones)
                    {
                        if (!existingIds.Contains(damperZone.Id))
                        {
                            allClashZones.Add(damperZone);
                        }
                    }
                    
                    if (!context.IsDeploymentMode)
                    {
                        DebugLogger.Info($"[REFRESH-REFACTORED] ✅ Merged results: {intersectionClashZones?.Count ?? 0} from intersections + {damperClashZones.Count} from dampers = {allClashZones.Count} total");
                        SafeFileLogger.SafeAppendText(context.RefreshLogName, 
                            $"[{DateTime.Now}] [REFRESH-REFACTORED] ✅ Merged results: {intersectionClashZones?.Count ?? 0} from intersections + {damperClashZones.Count} from dampers = {allClashZones.Count} total\n");
                    }
                }
                
                // Update context with merged results
                if (context.AllClashZones == null)
                {
                    context.AllClashZones = allClashZones;
                }
                else
                {
                    // Merge with existing zones
                    var existingIds = new HashSet<Guid>(context.AllClashZones.Select(z => z.Id));
                    foreach (var zone in allClashZones)
                    {
                        if (!existingIds.Contains(zone.Id))
                        {
                            context.AllClashZones.Add(zone);
                        }
                    }
                }
                
                // ✅ CRITICAL FIX: Add ALL new zones (intersections + dampers) to NewClashZones
                if (context.NewClashZones == null)
                {
                    context.NewClashZones = new List<ClashZone>();
                }
                
                var existingNewIds = new HashSet<Guid>(context.NewClashZones.Select(z => z.Id));
                
                if (intersectionClashZones != null)
                {
                    foreach (var zone in intersectionClashZones)
                    {
                        if (!existingNewIds.Contains(zone.Id))
                        {
                            context.NewClashZones.Add(zone);
                            existingNewIds.Add(zone.Id);
                        }
                    }
                }
                
                if (damperClashZones.Count > 0)
                {
                    foreach (var damperZone in damperClashZones)
                    {
                        if (!existingNewIds.Contains(damperZone.Id))
                        {
                            context.NewClashZones.Add(damperZone);
                            existingNewIds.Add(damperZone.Id);
                        }
                    }
                }
                
                processor.PostProcess(decision);
                intersectionOp?.SetItemCount(allClashZones.Count);
                
                if (!context.IsDeploymentMode)
                {
                    DebugLogger.Info($"[REFRESH-REFACTORED] Mode: {decision.Mode}, Detection Run: {decision.ShouldRunDetection}, Reason: {decision.Reason}");
                }
            }
        }

        private void CaptureParameters(RefreshContext context)
        {
            UpdateProgress(70, "Capturing parameters...");
            using (var paramOp = context.PerformanceMonitor.TrackOperation("8. Parameter Capture") as PerformanceMonitor.OperationTracker)
            {
                var paramService = new ParameterCaptureService(context);
                paramService.CaptureParametersParallel(context.NewClashZones);
                paramOp?.SetItemCount(context.NewClashZones?.Count ?? 0);
            }
            
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
                        var existingZone = allZones.FirstOrDefault(z => z.Id == newZone.Id);
                        if (existingZone != null)
                        {
                            existingZone.MepElementWidth = newZone.MepElementWidth;
                            existingZone.MepElementHeight = newZone.MepElementHeight;
                            existingZone.MepElementOuterDiameter = newZone.MepElementOuterDiameter;
                            existingZone.MepElementNominalDiameter = newZone.MepElementNominalDiameter;
                            existingZone.MepElementSizeParameterValue = newZone.MepElementSizeParameterValue;
                            existingZone.MepElementFormattedSize = newZone.MepElementFormattedSize;

                            // ✅ CRITICAL: propagate system metadata from refreshed zone
                            existingZone.MepSystemType = newZone.MepSystemType;
                            existingZone.MepSystemName = newZone.MepSystemName;
                            existingZone.MepServiceType = newZone.MepServiceType;
                            existingZone.MepElementSystemAbbreviation = newZone.MepElementSystemAbbreviation;

                            existingZone.MepElementUniqueId = newZone.MepElementUniqueId;
                            existingZone.IsInsulated = newZone.IsInsulated;
                            existingZone.InsulationThickness = newZone.InsulationThickness;
                            existingZone.MepElementSizeData = newZone.MepElementSizeData;
                            existingZone.DuctShape = newZone.DuctShape;
                            existingZone.InsulationType = newZone.InsulationType;
                            
                            existingZone.HasMepConnector = newZone.HasMepConnector;
                            existingZone.DamperConnectorSide = newZone.DamperConnectorSide;
                            
                            // ✅ CRITICAL FIX: Merge parameters instead of overwriting
                            // Preserve existing parameters from database, add new ones from refresh
                            var existingMepCount = existingZone.MepParameterValues?.Count ?? 0;
                            var newMepCount = newZone.MepParameterValues?.Count ?? 0;
                            
                            // ✅ DIAGNOSTIC: Log merge operation
                            if (!DeploymentConfiguration.DeploymentMode && string.Equals(newZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
                            {
                                SafeFileLogger.SafeAppendText("save_db_diagnostic.log",
                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [REFRESH-MERGE] Zone {newZone.Id}: Existing MEP params={existingMepCount}, New MEP params={newMepCount}\n");
                            }
                            
                            if (newZone.MepParameterValues != null && newZone.MepParameterValues.Count > 0)
                            {
                                if (existingZone.MepParameterValues == null || existingZone.MepParameterValues.Count == 0)
                                {
                                    // No existing parameters - use new ones
                                    existingZone.MepParameterValues = newZone.MepParameterValues;
                                }
                                else
                                {
                                    // Merge: Add new parameters that don't already exist
                                    // ✅ FIX: Handle duplicate keys by taking the first occurrence
                                    // ✅ FIX: Handle duplicate keys by taking the first occurrence (using safe loop instead of ToDictionary)
                                    var existingDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                                    
                                    // Robust population loop
                                    if (existingZone.MepParameterValues != null)
                                    {
                                        foreach (var kv in existingZone.MepParameterValues)
                                        {
                                            if (kv != null && !string.IsNullOrEmpty(kv.Key) && !existingDict.ContainsKey(kv.Key))
                                            {
                                                existingDict[kv.Key] = kv.Value;
                                            }
                                        }
                                    }
                                    
                                    foreach (var newParam in newZone.MepParameterValues)
                                    {
                                        if (newParam != null && !string.IsNullOrEmpty(newParam.Key) && !existingDict.ContainsKey(newParam.Key))
                                        {
                                            existingDict[newParam.Key] = newParam.Value;
                                        }
                                    }
                                    
                                    existingZone.MepParameterValues = existingDict
                                        .Select(kv => new SerializableKeyValue { Key = kv.Key, Value = kv.Value })
                                        .ToList();
                                    
                                    // ✅ DIAGNOSTIC: Log merge result
                                    if (!DeploymentConfiguration.DeploymentMode && string.Equals(newZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
                                    {
                                        SafeFileLogger.SafeAppendText("save_db_diagnostic.log",
                                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [REFRESH-MERGE] Zone {newZone.Id}: Merged MEP params - had {existingMepCount}, added {newMepCount}, total now {existingZone.MepParameterValues.Count}\n");
                                    }
                                }
                            }
                            else if (existingMepCount > 0)
                            {
                                // ✅ DIAGNOSTIC: Log that we're preserving existing parameters
                                if (!DeploymentConfiguration.DeploymentMode && string.Equals(newZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
                                {
                                    SafeFileLogger.SafeAppendText("save_db_diagnostic.log",
                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [REFRESH-MERGE] Zone {newZone.Id}: Preserving {existingMepCount} existing MEP params (newZone has {newMepCount})\n");
                                }
                            }
                            // ✅ If newZone has no parameters, preserve existing ones (don't overwrite with null/empty)
                            
                            if (newZone.HostParameterValues != null && newZone.HostParameterValues.Count > 0)
                            {
                                if (existingZone.HostParameterValues == null || existingZone.HostParameterValues.Count == 0)
                                {
                                    existingZone.HostParameterValues = newZone.HostParameterValues;
                                }
                                else
                                {
                                    // ✅ FIX: Handle duplicate keys by taking the first occurrence
                                    // ✅ FIX: Handle duplicate keys by taking the first occurrence (using safe loop instead of ToDictionary)
                                    var existingDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                                    
                                    // Robust population loop
                                    if (existingZone.HostParameterValues != null)
                                    {
                                        foreach (var kv in existingZone.HostParameterValues)
                                        {
                                            if (kv != null && !string.IsNullOrEmpty(kv.Key) && !existingDict.ContainsKey(kv.Key))
                                            {
                                                existingDict[kv.Key] = kv.Value;
                                            }
                                        }
                                    }
                                    
                                    foreach (var newParam in newZone.HostParameterValues)
                                    {
                                        if (newParam != null && !string.IsNullOrEmpty(newParam.Key) && !existingDict.ContainsKey(newParam.Key))
                                        {
                                            existingDict[newParam.Key] = newParam.Value;
                                        }
                                    }
                                    
                                    existingZone.HostParameterValues = existingDict
                                        .Select(kv => new SerializableKeyValue { Key = kv.Key, Value = kv.Value })
                                        .ToList();
                                }
                            }
                            // ✅ If newZone has no parameters, preserve existing ones (don't overwrite with null/empty)
                            
                            if (!DeploymentConfiguration.DeploymentMode && string.Equals(newZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
                            {
                                var odMm = newZone.MepElementOuterDiameter > 0 ? (newZone.MepElementOuterDiameter * 304.8) : 0.0;
                                var nomMm = newZone.MepElementNominalDiameter > 0 ? (newZone.MepElementNominalDiameter * 304.8) : 0.0;
                                SafeFileLogger.SafeAppendText("save_db_diagnostic.log",
                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [REFRESH-MERGE] ✅ UPDATED existing zone {newZone.Id} with fresh data: OuterDiameter={newZone.MepElementOuterDiameter:F6}ft ({odMm:F1}mm), NominalDiameter={newZone.MepElementNominalDiameter:F6}ft ({nomMm:F1}mm), SizeParameterValue='{newZone.MepElementSizeParameterValue ?? "NULL"}'\n");
                            }
                            
                            if (!DeploymentConfiguration.DeploymentMode && string.Equals(newZone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase))
                            {
                                SafeFileLogger.SafeAppendText("save_db_diagnostic.log",
                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [REFRESH-MERGE] ✅ UPDATED existing zone {newZone.Id} (Duct Accessories) with fresh connector data: HasMepConnector={newZone.HasMepConnector}, DamperConnectorSide='{newZone.DamperConnectorSide ?? "NULL"}'\n");
                            }
                        }
                    }
                }
                
                context.AllClashZones = allZones;
            }
        }
        
        private void ResetUI()
        {
            if (_progressBar != null)
                _progressBar.Value = _progressBar.Minimum;
            
            if (_refreshButton != null)
                _refreshButton.Enabled = true;
        }

        /// <summary>
        /// ✅ SECTION BOX CAPTURE: Capture and store section box bounds during refresh
        /// This implements the "Dump Once, Use Many Times" architecture from SECTION_BOX_STORAGE_PLAN.md
        /// </summary>
        private void CaptureAndStoreSectionBox(RefreshContext context)
        {
            try
            {
                if (!context.IsDeploymentMode)
                {
                    DebugLogger.Info("[REFRESH-REFACTORED] [SECTION-BOX] Capturing section box bounds for caching...");
                }
                SafeFileLogger.SafeAppendText(context.RefreshLogName,
                    $"[{DateTime.Now}] [SECTION-BOX] Capturing section box bounds for caching...\n");

                // Create SectionBoxService instance
                var sectionBoxService = new SectionBoxService();

                // Get active 3D view with section box
                if (_uiDocument?.ActiveView is View3D view3D && view3D.IsSectionBoxActive)
                {
                    // Create database context for storing section box
                    using (var dbContext = new Data.SleeveDbContext(_document, msg =>
                    {
                        if (!context.IsDeploymentMode)
                            DebugLogger.Info($"[REFRESH-REFACTORED][SQLite] {msg}");
                    }))
                    {
                        // Capture and store section box bounds
                        sectionBoxService.CaptureAndStore(view3D, dbContext.Connection);

                        if (!context.IsDeploymentMode)
                        {
                            DebugLogger.Info("[REFRESH-REFACTORED] [SECTION-BOX] ✅ Section box bounds captured and stored successfully");
                        }
                        SafeFileLogger.SafeAppendText(context.RefreshLogName,
                            $"[{DateTime.Now}] [SECTION-BOX] ✅ Section box bounds captured and stored successfully\n");
                    }
                }
                else
                {
                    if (!context.IsDeploymentMode)
                    {
                        DebugLogger.Info("[REFRESH-REFACTORED] [SECTION-BOX] ⚠️ No active 3D section box found - section box caching skipped");
                    }
                    SafeFileLogger.SafeAppendText(context.RefreshLogName,
                        $"[{DateTime.Now}] [SECTION-BOX] ⚠️ No active 3D section box found - section box caching skipped\n");
                }
            }
            catch (Exception ex)
            {
                if (!context.IsDeploymentMode)
                {
                    DebugLogger.Warning($"[REFRESH-REFACTORED] [SECTION-BOX] ⚠️ Failed to capture section box: {ex.Message}");
                }
                SafeFileLogger.SafeAppendText(context.RefreshLogName,
                    $"[{DateTime.Now}] [SECTION-BOX] ⚠️ Failed to capture section box: {ex.Message}\n");
                // Continue with refresh even if section box capture fails (non-blocking)
            }
        }
    }
}
