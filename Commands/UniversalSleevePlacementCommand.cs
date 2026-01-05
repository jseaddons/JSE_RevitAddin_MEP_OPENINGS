using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Diagnostics;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Strategies;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using System.Threading.Tasks;
using JSE_RevitAddin_MEP_OPENINGS.Services.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using System.IO;
using System;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Services.Refactored;
using JSE_RevitAddin_MEP_OPENINGS.Services.FlagManagement;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// Universal sleeve placement command - handles ALL MEP categories using strategy pattern
    /// Based on proven DuctSleevePlacementCommand pattern
    /// 
    /// Supports: Ducts, Pipes, Cable Trays, Duct Accessories (Dampers)
    /// Uses 4 universal families: OpeningOnWall-Rectangular, OpeningOnWall-Circular, 
    ///                           OpeningOnSlab-Rectangular, OpeningOnSlab-Circular
    /// </summary>
    public class UniversalSleevePlacementCommand : ICommand
    {
        private readonly Document _doc;
        private readonly List<ClashZone> _clashZones;
        private readonly string _category;
        private readonly string _filterName;
        private readonly ISleevePlacementStrategy _strategy;
        private readonly string _logPrefix;
        private OpeningConditions _conditions;
        private readonly Dictionary<string, double> _clearanceSettings;
        
        // ✅ SOLID REFACTORED: Injected services (optional - created if null when flag is enabled)
        private readonly IConditionsLoader? _conditionsLoader;
        private readonly IPathDeterminer? _pathDeterminer;
        private readonly IStrategyFactory? _strategyFactory;
        private readonly IDocumentValidator? _documentValidator;
        private readonly IUiStateProvider? _uiStateProvider;
        private readonly IFileNameNormalizer? _fileNameNormalizer;
        private readonly ISectionBoxChecker? _sectionBoxChecker;
        private readonly Dictionary<Guid, SleevePlacementPlanningDto>? _externalPlanningMap;
        
        // ✅ PERFORMANCE: Properties to expose placement counts
        public int PlacedCount { get; private set; }
        public int SkippedCount { get; private set; }
        public int ErrorCount { get; private set; }

        public UniversalSleevePlacementCommand(
            Document doc, 
            List<ClashZone> clashZones, 
            string category, 
            string filterName, 
            Dictionary<string, double>? clearanceSettings = null,
            // ✅ SOLID REFACTORED: Optional injected services (created automatically if null when flag enabled)
            IConditionsLoader? conditionsLoader = null,
            IPathDeterminer? pathDeterminer = null,
            IStrategyFactory? strategyFactory = null,
            IDocumentValidator? documentValidator = null,
            IUiStateProvider? uiStateProvider = null,
            IFileNameNormalizer? fileNameNormalizer = null,
            ISectionBoxChecker? sectionBoxChecker = null,
            Dictionary<Guid, SleevePlacementPlanningDto>? externalPlanningMap = null)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _clashZones = clashZones ?? throw new ArgumentNullException(nameof(clashZones));
            _category = category ?? throw new ArgumentNullException(nameof(category));
            _filterName = filterName ?? "Unknown";
            _logPrefix = $"[UniversalSleeveCommand-{category}]";
            
            // ✅ SOLID REFACTORED: Initialize services based on flag
            if (OptimizationFlags.UseRefactoredCommandServices)
            {
                _conditionsLoader = conditionsLoader ?? new ConditionsLoaderService(doc);
                _pathDeterminer = pathDeterminer ?? new PathDeterminerService();
                _strategyFactory = strategyFactory ?? new StrategyFactoryService();
                _documentValidator = documentValidator ?? new DocumentValidatorService();
                _uiStateProvider = uiStateProvider ?? new UiStateProviderService();
                _fileNameNormalizer = fileNameNormalizer ?? new FileNameNormalizerService();
                _sectionBoxChecker = sectionBoxChecker ?? new SectionBoxCheckerService();
                
                // Use refactored services
                _strategy = _strategyFactory.CreateStrategy(category, doc);
                _conditions = _conditionsLoader.LoadConditions(_filterName, _category);
            }
            else
            {
                // Legacy inline implementations
                _strategy = CreateStrategy(category);
                LoadConditionsFromXml();
            }
            
            // ✅ NEW: Store clearance settings for direct UI access
            _clearanceSettings = clearanceSettings ?? new Dictionary<string, double>();
            _externalPlanningMap = externalPlanningMap;
            
            DebugLogger.Info($"{_logPrefix} Constructor: Received {_clearanceSettings.Count} clearance settings from UI");
            foreach (var kvp in _clearanceSettings)
            {
                DebugLogger.Info($"{_logPrefix} Clearance: {kvp.Key} = {kvp.Value}mm");
            }
        }

        public void Execute(UIApplication app)
        {
            try
            {
                // ✅ CRITICAL FIX: Reset Logger Context and FORCE ENABLE for debugging
                DebugLogger.IsEnabled = true;
                DebugLogger.SetServiceContext($"UniversalSleeve_{_category}");
                
                // Explicitly set the log file based on category
                if (_category.Contains("Duct") && !_category.Contains("Accessory"))
                {
                    DebugLogger.SetDuctLogFile(); 
                }
                else if (_category.Contains("Cable"))
                {
                    DebugLogger.SetCableTrayLogFile();
                }
                else if (_category.Contains("Accessory") || _category.Contains("Damper"))
                {
                    DebugLogger.SetDamperLogFile();
                }
                else
                {
                    // For Pipes and others, use a standardized name
                    string normCat = MepCategoryConstants.Normalize(_category);
                    DebugLogger.InitLogFile($"{normCat}sleeveplacer");
                }

                // 🚨 DEBUG: Direct file logging to bypass DebugLogger issues
                DebugLogger.Info($"[{DateTime.Now}] 🚨 UniversalSleevePlacementCommand.Execute STARTED for category '{_category}' with {_clashZones.Count} clash zones\n");
                
                DebugLogger.Info($"{_logPrefix} Starting sleeve placement for {_clashZones.Count} clash zones");
                
                // Fallback: if no clash zones were passed, load from filter XML saved during refresh
                if (_clashZones == null || _clashZones.Count == 0)
                {
                    var dataService = new ClashZoneDataService(
                        _doc,
                        message =>
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"{_logPrefix} {message}");
                        });

                    var normalizedCategory = MepCategoryConstants.Normalize(_category);
                    var normalizedFilter = FilterNameHelper.NormalizeBaseName(_filterName, _filterName, normalizedCategory);
                    var normalizationTimer = Stopwatch.StartNew();
                    var loadedZones = dataService.LoadClashZonesForCategory(normalizedFilter, normalizedCategory);
                    normalizationTimer.Stop();
                    
                    DebugLogger.Info($"{_logPrefix} Loaded {loadedZones?.Count ?? 0} zones from database in {normalizationTimer.ElapsedMilliseconds}ms");

                    // 🔍 TRACING: Log all loaded IDs and check for problematic ones
                    if (loadedZones != null)
                    {
                        foreach (var cz in loadedZones)
                        {
                            if (cz.SleeveInstanceId == 1183693 || cz.SleeveInstanceId == 1183702)
                            {
                                DebugLogger.Info($"[TRACE-FOUND] Zone {cz.Id} matches Revit ID {cz.SleeveInstanceId}. IsCurrentClash={cz.IsCurrentClash}, ReadyForPlacement={cz.ReadyForPlacementFlag}, IsClusterResolved={cz.IsClusterResolvedFlag}");
                            }
                        }
                    }

                    if (loadedZones.Count > 0)
                    {
                        _clashZones.Clear();
                        _clashZones.AddRange(loadedZones);
                        DebugLogger.Info($"{_logPrefix} Loaded {_clashZones.Count} clash zones via ClashZoneDataService (SQLite primary, XML fallback).");
                        
                        // ✅ CRITICAL FIX: Sync flags from database after loading zones
                        // Zones loaded from DB should have correct flags, but ensure they're synced
                        try
                        {
                            var flagManager = Services.FlagManagement.FlagManagerFactory.CreateAdapter(_doc);
                            flagManager.SyncFlagsFromGlobal(_clashZones, normalizedCategory);
                            
                            // ✅ DIAGNOSTIC: Log flag status after sync
                            int resolvedCount = _clashZones.Count(z => z.IsResolved);
                            int clusterResolvedCount = _clashZones.Count(z => z.IsClusterResolved);
                            int eligibleCount = _clashZones.Count(z => !z.IsResolved && !z.IsClusterResolved);
                            DebugLogger.Info($"{_logPrefix} ✅ Flags synced from database: Total={_clashZones.Count}, IsResolved={resolvedCount}, IsClusterResolved={clusterResolvedCount}, Eligible={eligibleCount}");
                            
                            // ✅ DIAGNOSTIC: Log sample zone flags for debugging
                            if (!DeploymentConfiguration.DeploymentMode && _clashZones.Count > 0)
                            {
                                var sampleZones = _clashZones.Take(5).ToList();
                                foreach (var zone in sampleZones)
                                {
                                    DebugLogger.Info($"{_logPrefix} Sample Zone {zone.Id}: IsResolved={zone.IsResolved}, IsClusterResolved={zone.IsClusterResolved}, SleeveId={zone.SleeveInstanceId}, ClusterId={zone.ClusterSleeveInstanceId}");
                                }
                            }
                        }
                        catch (Exception flagEx)
                        {
                            DebugLogger.Warning($"{_logPrefix} ⚠️ Failed to sync flags from database: {flagEx.Message}");
                        }
                    }
                }

                if (_clashZones == null || _clashZones.Count == 0)
                {
                    DebugLogger.Warning($"{_logPrefix} ⚠️ No clash zones available for placement. Ensure Refresh completed successfully and data was saved.");
                    return;
                }
                
                // ---- 1. VALIDATION: Check document state (NO transaction) ----
                bool isValid = OptimizationFlags.UseRefactoredCommandServices && _documentValidator != null
                    ? _documentValidator.ValidateDocument(_doc, _logPrefix)
                    : ValidateDocument();
                
                if (!isValid)
                {
                    DebugLogger.Error($"{_logPrefix} Document validation failed - cannot place sleeves");
                    return;
                }

                // ---- 3. PLACEMENT LOGIC ----
                // Check if we already have an active transaction (from Orchestrator)
                if (_doc.IsModifiable)
                {
                    DebugLogger.Info($"{_logPrefix} Document is already modifiable - executing within existing transaction");
                    PerformPlacement();
                }
                else
                {
                    using (var t = new Transaction(_doc, $"Place {_category} Sleeves"))
                    {
                        if (t.Start() == TransactionStatus.Started)
                        {
                            var options = t.GetFailureHandlingOptions();
                            options.SetFailuresPreprocessor(new UniversalWarningSwallower());
                            t.SetFailureHandlingOptions(options);
                            
                            PerformPlacement();
                            
                            var status = t.Commit();
                            if (status == TransactionStatus.Committed)
                            {
                                DebugLogger.Info($"{_logPrefix} ✓ Transaction committed - Placed: {PlacedCount}, Skipped: {SkippedCount}");
                            }
                            else
                            {
                                DebugLogger.Error($"{_logPrefix} Transaction failed to commit: {status}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"{_logPrefix} Exception in Execute: {ex.Message}");
                DebugLogger.Error($"{_logPrefix} Stack trace: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// Logic for sleeve placement, extracted to support both internal and external transactions.
        /// </summary>
        private void PerformPlacement()
        {
            // ✅ RESPECT MASTER SWITCH
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"{_logPrefix} ✅ Diagnostic Mode: UseDiagnosticMode={OptimizationFlags.UseDiagnosticMode}, DeploymentMode={DeploymentConfiguration.DeploymentMode}");
            }

            var filteredClashZones = FilterClashZonesByAllCriteria(_clashZones);

            // ✅ REFACTORED: Apply strict hierarchical filtering
            if (OptimizationFlags.UseRefactoredClashZoneFlagServices)
            {
                filteredClashZones = filteredClashZones?.Where(cz => 
                {
                    if (!cz.IsCurrentClash) return false;
                    if (cz.IsCombinedResolved) return false;
                    if (cz.IsClusterResolved) return false;
                    if (cz.IsResolved) return false;
                    return true;
                }).ToList();
            }

            int total = filteredClashZones?.Count ?? 0;
            int eligible = filteredClashZones?.Count(cz => !cz.IsResolved && !cz.IsClusterResolved) ?? 0;
            
            if (eligible == 0)
            {
                DebugLogger.Info($"{_logPrefix} ⚠️ No eligible clash zones found for placement (Total={total}, Eligible=0)");
                return;
            }

            // ✅ PATH DETERMINATION
            SleevePlacementPath placementPath;
            if (OptimizationFlags.UseRefactoredCommandServices && _pathDeterminer != null)
            {
                Dictionary<string, double>? clearanceDict = null;
                if (_conditions?.ClearanceSettings != null)
                {
                    clearanceDict = new Dictionary<string, double>();
                    var cs = _conditions.ClearanceSettings;
                    if (cs.RectangularNormal > 0) clearanceDict["RectangularNormal"] = cs.RectangularNormal;
                    if (cs.RectangularInsulated > 0) clearanceDict["RectangularInsulated"] = cs.RectangularInsulated;
                    if (cs.RoundNormal > 0) clearanceDict["RoundNormal"] = cs.RoundNormal;
                    if (cs.RoundInsulated > 0) clearanceDict["RoundInsulated"] = cs.RoundInsulated;
                    if (cs.PipesNormal > 0) clearanceDict["PipesNormal"] = cs.PipesNormal;
                    if (cs.PipesInsulated > 0) clearanceDict["PipesInsulated"] = cs.PipesInsulated;
                    if (cs.CableTrayTop > 0) clearanceDict["CableTrayTop"] = cs.CableTrayTop;
                    if (cs.CableTrayOther > 0) clearanceDict["CableTrayOther"] = cs.CableTrayOther;
                }
                placementPath = _pathDeterminer.DeterminePath(_doc, _filterName, _category, _logPrefix, clearanceDict);
            }
            else
            {
                placementPath = DeterminePlacementPath();
            }
            bool isReplayPath = placementPath == SleevePlacementPath.Replay;
            
            // ✅ EXECUTE PLACEMENT
            var sleeveRepository = new Services.Repositories.SleeveRepository();
            var zoneFilterService = new ZoneFilterService();
            
            IFlagManager flagManager = FlagManagerFactory.CreateAdapter(_doc);
            
            var profileService = ApplicationProfileService.Instance;
            var settings = profileService.GetCurrentSettings();
            bool isForceDetectionMode = settings.ForceDetectionMode;
            
            var newPlacerService = new NewSleevePlacerService(
                _doc,
                _conditions,
                _strategy,
                _clearanceSettings,
                sleeveRepository,
                zoneFilterService,
                null, 
                flagManager, 
                isReplayPath,
                _filterName,
                null, 
                isForceDetectionMode,
                null,
                _externalPlanningMap);
            
            var (placed, skipped, errors) = newPlacerService.PlaceAllSleevesInTransaction(filteredClashZones);
            
            PlacedCount = placed;
            SkippedCount = skipped;
            ErrorCount = errors;
            
            DebugLogger.Info($"{_logPrefix} Placement complete - Placed: {placed}, Skipped: {skipped}, Errors: {errors}");
        }
        
        private ISleevePlacementStrategy CreateStrategy(string category)
        {
            return category switch
            {
                "Ducts" => new DuctPlacementStrategy(),
                "Pipes" => new PipePlacementStrategy(),
                "Cable Trays" => new CableTrayPlacementStrategy(),
                "Duct Accessories" => new DamperPlacementStrategy(_doc),
                _ => throw new ArgumentException($"Unknown category: {category}")
            };
        }
        
        private bool ValidateDocument()
        {
            DebugLogger.Info($"{_logPrefix} Document validation - Path: {_doc.PathName}");
            DebugLogger.Info($"{_logPrefix} Document validation - IsModifiable: {_doc.IsModifiable}");
            DebugLogger.Info($"{_logPrefix} Document validation - IsLinked: {_doc.IsLinked}");
            DebugLogger.Info($"{_logPrefix} Document validation - IsWorkshared: {_doc.IsWorkshared}");
            
            // Test if we can modify
            bool canModify = false;
            try
            {
                using (var testTransaction = new Transaction(_doc, "Test Modification"))
                {
                    if (testTransaction.Start() == TransactionStatus.Started)
                    {
                        canModify = true;
                        testTransaction.RollBack();
                        DebugLogger.Info($"{_logPrefix} Document can be modified");
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"{_logPrefix} Document modification test failed: {ex.Message}");
            }
            
            return canModify;
        }

        /// <summary>
        /// ✅ OOP: Uses RefreshPathDeterminer service to determine placement path based on IsFilterComboNew flag
        /// ✅ NEW: Also checks if OpeningSettings changed - routes to PATH 2 if changed
        /// </summary>
        private SleevePlacementPath DeterminePlacementPath()
        {
            // Convert OpeningConditions.ClearanceSettings to Dictionary<string, double> for comparison
            Dictionary<string, double> currentClearanceSettings = null;
            if (_conditions?.ClearanceSettings != null)
            {
                currentClearanceSettings = new Dictionary<string, double>();
                var cs = _conditions.ClearanceSettings;
                
                // Add all clearance settings to dictionary
                if (cs.RectangularNormal > 0) currentClearanceSettings["RectangularNormal"] = cs.RectangularNormal;
                if (cs.RectangularInsulated > 0) currentClearanceSettings["RectangularInsulated"] = cs.RectangularInsulated;
                if (cs.RoundNormal > 0) currentClearanceSettings["RoundNormal"] = cs.RoundNormal;
                if (cs.RoundInsulated > 0) currentClearanceSettings["RoundInsulated"] = cs.RoundInsulated;
                if (cs.PipesNormal > 0) currentClearanceSettings["PipesNormal"] = cs.PipesNormal;
                if (cs.PipesInsulated > 0) currentClearanceSettings["PipesInsulated"] = cs.PipesInsulated;
                if (cs.CableTrayTop > 0) currentClearanceSettings["CableTrayTop"] = cs.CableTrayTop;
                if (cs.CableTrayOther > 0) currentClearanceSettings["CableTrayOther"] = cs.CableTrayOther;
            }
            
            return Services.Refresh.RefreshPathDeterminer.DeterminePlacementPath(
                _doc, 
                _filterName, 
                _category, 
                _logPrefix,
                currentClearanceSettings);
        }
        
        private void LoadConditionsFromXml()
        {
            try
            {
                // Use the actual project filters directory so CONDITIONS.xml sits next to the filter XMLs
                var projectFiltersDir = ProjectPathService.GetFiltersDirectory(_doc);
                var conditionsService = new ConditionsService(_doc, projectFiltersDir, msg => DebugLogger.Info(msg));
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    try 
                    { 
                        string orchestratorLogPath = SafeFileLogger.GetLogFilePath("orchestrator_debug.log");
                        System.IO.File.AppendAllText(orchestratorLogPath, $"[{DateTime.Now:HH:mm:ss}] CONDITIONS_DIR={projectFiltersDir}\n"); 
                    } 
                    catch { }
                }
                
                // 🛡️ ARCHITECTURE FIX: Use BOTH filter name AND category for unique CONDITIONS XML
                // This allows different clearance/opening types per category within the same filter
                // File naming: "FilterName_Category_CONDITIONS.xml" (e.g., "Ventilation_Pipes_CONDITIONS.xml")
                
                // Get the actual selected filter names from the current UI state
                var selectedFilterNames = GetSelectedFilterNamesFromUI();
                
                if (selectedFilterNames.Count == 0)
                {
                    DebugLogger.Warning($"{_logPrefix} No filter names found in UI - using default conditions");
                    _conditions = new OpeningConditions { FilterName = "Default", Category = _category };
                    return;
                }
                
                // Use the first selected filter name + current category
                string filterName = selectedFilterNames.First();
                
                // ✅ STANDARDIZED: Use normalized category names to match clash zone file naming
                string normalizedCategory = NormalizeCategoryName(_category);
                string combinedKey = $"{filterName}_{normalizedCategory}";
                
                DebugLogger.Info($"{_logPrefix} Using combined key '{combinedKey}' (Filter: '{filterName}', Category: '{_category}')");
                
                // ⚠️ DIAGNOSTIC: Log the exact file path being loaded
                string expectedFileName = $"{combinedKey}_CONDITIONS.xml";
                DebugLogger.Info($"{_logPrefix} Expected CONDITIONS file: '{expectedFileName}'");
                
                // 🔥 UNCONDITIONAL: Log before loading conditions
                SafeFileLogger.SafeAppendText("cabletray_dimension_trace.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [LOAD-CONDITIONS] Filter={filterName}, Category={_category}, CombinedKey={combinedKey}\n");
                
                _conditions = conditionsService.LoadConditions(combinedKey);
                
                // 🔥 UNCONDITIONAL: Log after loading conditions
                SafeFileLogger.SafeAppendText("cabletray_dimension_trace.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [LOADED-CONDITIONS] Loaded={_conditions != null}, " +
                    $"ClearanceSettings={_conditions?.ClearanceSettings != null}\n");
                if (_conditions?.ClearanceSettings != null)
                {
                    SafeFileLogger.SafeAppendText("cabletray_dimension_trace.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [LOADED-CLEARANCES] CableTrayTop={_conditions.ClearanceSettings.CableTrayTop}, " +
                        $"CableTrayOther={_conditions.ClearanceSettings.CableTrayOther}\n");
                }
                
                // Ensure CONDITIONS.xml exists in the project Filters directory; create if missing
                try
                {
                    var expectedPath = conditionsService.GetConditionsFilePath(combinedKey);
                    if (!System.IO.File.Exists(expectedPath))
                    {
                        // Initialize sane defaults tied to this filter/category key
                        if (_conditions == null)
                            _conditions = new OpeningConditions();
                        _conditions.FilterName = filterName;
                        _conditions.Category = _category;
                        // Save immediately so subsequent runs find it
                        conditionsService.SaveConditions(_conditions, combinedKey);
                        DebugLogger.Info($"{_logPrefix} CONDITIONS.xml created at: {expectedPath}");
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            string orchestratorLogPath2 = SafeFileLogger.GetLogFilePath("orchestrator_debug.log");
                            System.IO.File.AppendAllText(orchestratorLogPath2, $"[{DateTime.Now:HH:mm:ss}] CONDITIONS_CREATED={expectedPath}\n");
                        }
                    }
                    else
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            string orchestratorLogPath3 = SafeFileLogger.GetLogFilePath("orchestrator_debug.log");
                            System.IO.File.AppendAllText(orchestratorLogPath3, $"[{DateTime.Now:HH:mm:ss}] CONDITIONS_EXISTS={expectedPath}\n");
                        }
                    }
                }
                catch { }
                
                if (_conditions != null)
                {
                    DebugLogger.Info($"{_logPrefix} Loaded conditions for '{combinedKey}' - Pipes: {_conditions.OpeningTypePreferences?.Pipes ?? "null"}, RoundDucts: {_conditions.OpeningTypePreferences?.RoundDucts ?? "null"}");
                    
                    // ⚠️ DIAGNOSTIC: Log clearance values from CONDITIONS XML
                    if (_conditions.ClearanceSettings != null)
                    {
                        DebugLogger.Info($"{_logPrefix} CONDITIONS XML Clearances: RectNormal={_conditions.ClearanceSettings.RectangularNormal}mm, RectInsulated={_conditions.ClearanceSettings.RectangularInsulated}mm");
                        DebugLogger.Info($"{_logPrefix} CONDITIONS XML Clearances: RoundNormal={_conditions.ClearanceSettings.RoundNormal}mm, RoundInsulated={_conditions.ClearanceSettings.RoundInsulated}mm");
                        DebugLogger.Info($"{_logPrefix} CONDITIONS XML Clearances: PipesNormal={_conditions.ClearanceSettings.PipesNormal}mm, PipesInsulated={_conditions.ClearanceSettings.PipesInsulated}mm");
                        DebugLogger.Info($"{_logPrefix} CONDITIONS XML Clearances: CableTrayTop={_conditions.ClearanceSettings.CableTrayTop}mm, CableTrayOther={_conditions.ClearanceSettings.CableTrayOther}mm");
                    }
                    else
                    {
                        DebugLogger.Warning($"{_logPrefix} CONDITIONS XML has NULL ClearanceSettings!");
                    }
                }
                else
                {
                    DebugLogger.Warning($"{_logPrefix} No conditions found for '{combinedKey}' - using defaults");
                    _conditions = new OpeningConditions { FilterName = filterName, Category = _category };
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"{_logPrefix} Error loading conditions: {ex.Message}");
                _conditions = new OpeningConditions { FilterName = "Default", Category = _category };
            }
        }
        
        /// <summary>
        /// Filter clash zones by all 5 criteria: MEP category, host type, reference files, host files, and 3D section box
        /// </summary>
        private List<ClashZone> FilterClashZonesByAllCriteria(List<ClashZone> clashZones)
        {
            try
            {
                // ✅ CRITICAL: Create placement_debug.log even if it wasn't created during XML load
                var debugLogPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                try { File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 FILTER STARTED: {clashZones.Count} zones incoming\n"); } catch { }
                
                // 🚨 DEBUG: Direct file logging to bypass DebugLogger issues
                DebugLogger.Info($"[{DateTime.Now}] 🚨 FilterClashZonesByAllCriteria STARTED with {clashZones.Count} clash zones\n");
                
                // ✅ CRITICAL: Log sample zones before filtering
                if (clashZones.Count > 0)
                {
                    try
                    {
                        File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] Sample zones BEFORE filtering:\n");
                        foreach (var cz in clashZones.Take(3))
                        {
                            File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}]   Zone {cz.Id}: Category='{cz.MepElementCategory}', HostType='{cz.StructuralElementType}', SourceDoc='{cz.SourceDocKey}', HostDoc='{cz.StructuralElementDocumentTitle}', IP=({cz.IntersectionPointX:F3},{cz.IntersectionPointY:F3},{cz.IntersectionPointZ:F3})\n");
                        }
                    }
                    catch { }
                }
                
                // Get selected host types from UI
                var selectedHostCategories = FilterUiStateProvider.GetSelectedHostCategories?.Invoke() ?? new List<string>();
                
                if (selectedHostCategories.Count == 0)
                {
                    DebugLogger.Warning($"{_logPrefix} No host types selected in UI - placing sleeves on all host types");
                }
                
                var allowedHostCategories = new HashSet<string>(selectedHostCategories, StringComparer.OrdinalIgnoreCase);
                DebugLogger.Info($"{_logPrefix} UI selected host types: [{string.Join(", ", selectedHostCategories)}]");
                
                // Get selected reference files and host files from UI
                List<string> selectedReferenceFiles;
                List<string> selectedHostFiles;
                if (OptimizationFlags.UseRefactoredCommandServices && _uiStateProvider != null)
                {
                    selectedReferenceFiles = _uiStateProvider.GetSelectedReferenceFiles();
                    selectedHostFiles = _uiStateProvider.GetSelectedHostFiles();
                }
                else
                {
                    selectedReferenceFiles = FilterUiStateProvider.GetSelectedReferenceFiles?.Invoke() ?? new List<string>();
                    selectedHostFiles = FilterUiStateProvider.GetSelectedHostFiles?.Invoke() ?? new List<string>();
                }
                
                // 🚨 DEBUG: Direct file logging to bypass DebugLogger issues
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string universalCommandDebugLogPath = SafeFileLogger.GetLogFilePath("universal_command_debug.log");
                    System.IO.File.AppendAllText(universalCommandDebugLogPath, $"[{DateTime.Now}] 🚨 UI selected reference files: [{string.Join(", ", selectedReferenceFiles)}]\n");
                    System.IO.File.AppendAllText(universalCommandDebugLogPath, $"[{DateTime.Now}] 🚨 UI selected host files: [{string.Join(", ", selectedHostFiles)}]\n");
                }
                
                DebugLogger.Info($"{_logPrefix} UI selected reference files: [{string.Join(", ", selectedReferenceFiles)}]");
                DebugLogger.Info($"{_logPrefix} UI selected host files: [{string.Join(", ", selectedHostFiles)}]");
                
                int afterCategory = 0, afterHostType = 0, afterRefFile = 0, afterHostFile = 0, afterSection = 0;
                var filteredZones = clashZones.Where(cz =>
                {
                    // ✅ CRITICAL: Log to placement_debug.log for each zone being filtered
                    // ✅ DEPLOYMENT MODE: Skip verbose per-zone logging (major performance impact)
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        var zoneDebugPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                        try
                        {
                            File.AppendAllText(zoneDebugPath, $"[{DateTime.Now:HH:mm:ss}] Checking Zone {cz.Id}: Category='{cz.MepElementCategory}' vs Command='{_category}', HostType='{cz.StructuralElementType}', SourceDoc='{cz.SourceDocKey}', HostDoc='{cz.StructuralElementDocumentTitle}'\n");
                        }
                        catch { }
                    }
                    
                    // 🚨 TROUBLESHOOTING: Re-enabling filters one by one
                    
                    // Filter 1: MEP category filtering (safety check) - RE-ENABLED FOR TESTING
                    // ✅ CRITICAL FIX: Normalize both category names before comparison to handle variations
                    // (e.g., "Duct Curves" vs "Ducts", "Cable Tray" vs "Cable Trays")
                    string normalizedZoneCategory = MepCategoryConstants.Normalize(cz.MepElementCategory);
                    string normalizedCommandCategory = MepCategoryConstants.Normalize(_category);
                    bool categoryMatch = string.Equals(normalizedZoneCategory, normalizedCommandCategory, StringComparison.OrdinalIgnoreCase);
                    
                    if (!categoryMatch)
                    {
                        DebugLogger.Info($"{_logPrefix} Filtered out ClashZone {cz.Id}: MEP category '{cz.MepElementCategory}' (normalized: '{normalizedZoneCategory}') doesn't match command category '{_category}' (normalized: '{normalizedCommandCategory}')");
                        // ✅ DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            try { File.AppendAllText(SafeFileLogger.GetLogFilePath("placement_debug.log"), $"[{DateTime.Now:HH:mm:ss}] ❌ FILTERED: Category mismatch - Zone='{cz.MepElementCategory}' (norm: '{normalizedZoneCategory}') vs Command='{_category}' (norm: '{normalizedCommandCategory}')\n"); } catch { }
                        }
                        return false;
                    }
                    afterCategory++;
                    // ✅ DEPLOYMENT MODE: Skip file writes
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        try { File.AppendAllText(SafeFileLogger.GetLogFilePath("placement_debug.log"), $"[{DateTime.Now:HH:mm:ss}] ✅ Passed category filter\n"); } catch { }
                    }
                    
                    // Filter 2: Host type filtering - RE-ENABLED FOR FINAL TESTING
                    bool hostTypeMatch = selectedHostCategories.Count == 0 || 
                                       allowedHostCategories.Contains(cz.StructuralElementType) ||
                                       allowedHostCategories.Contains(cz.StructuralElementType + "s") ||
                                       allowedHostCategories.Any(t => t.TrimEnd('s').Equals(cz.StructuralElementType, StringComparison.OrdinalIgnoreCase));
                    
                    if (!hostTypeMatch)
                    {
                        DebugLogger.Info($"{_logPrefix} Filtered out ClashZone {cz.Id}: Host type '{cz.StructuralElementType}' not in selected types [{string.Join(", ", selectedHostCategories)}]");
                        // ✅ DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            try { File.AppendAllText(SafeFileLogger.GetLogFilePath("placement_debug.log"), $"[{DateTime.Now:HH:mm:ss}] ❌ FILTERED: Host type '{cz.StructuralElementType}' not in [{string.Join(", ", selectedHostCategories)}]\n"); } catch { }
                        }
                        return false;
                    }
                    afterHostType++;
                    // ✅ DEPLOYMENT MODE: Skip file writes
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        try { File.AppendAllText(SafeFileLogger.GetLogFilePath("placement_debug.log"), $"[{DateTime.Now:HH:mm:ss}] ✅ Passed host type filter\n"); } catch { }
                    }
                    
                    // Filter 3: Reference linked file filtering - RE-ENABLED WITH DETAILED LOGGING
                    bool referenceFileMatch = selectedReferenceFiles.Count == 0 || 
                                            IsFileInSelectedList(cz.SourceDocKey, selectedReferenceFiles);
                    
                    if (!referenceFileMatch)
                    {
                        DebugLogger.Info($"{_logPrefix} Filtered out ClashZone {cz.Id}: Reference file '{cz.SourceDocKey}' not in selected files [{string.Join(", ", selectedReferenceFiles)}]");
                        // ✅ DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            try { File.AppendAllText(SafeFileLogger.GetLogFilePath("placement_debug.log"), $"[{DateTime.Now:HH:mm:ss}] ❌ FILTERED: Reference file '{cz.SourceDocKey}' not in [{string.Join(", ", selectedReferenceFiles)}]\n"); } catch { }
                        }
                        return false;
                    }
                    afterRefFile++;
                    // ✅ DEPLOYMENT MODE: Skip file writes
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        try { File.AppendAllText(SafeFileLogger.GetLogFilePath("placement_debug.log"), $"[{DateTime.Now:HH:mm:ss}] ✅ Passed reference file filter\n"); } catch { }
                    }
                    
                    // Filter 4: Host linked file filtering - RE-ENABLED WITH DETAILED LOGGING
                    bool hostFileMatch = selectedHostFiles.Count == 0 || 
                                       IsFileInSelectedList(cz.StructuralElementDocumentTitle, selectedHostFiles);
                    
                    if (!hostFileMatch)
                    {
                        DebugLogger.Info($"{_logPrefix} Filtered out ClashZone {cz.Id}: Host file '{cz.StructuralElementDocumentTitle}' not in selected files [{string.Join(", ", selectedHostFiles)}]");
                        // ✅ DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            try { File.AppendAllText(SafeFileLogger.GetLogFilePath("placement_debug.log"), $"[{DateTime.Now:HH:mm:ss}] ❌ FILTERED: Host file '{cz.StructuralElementDocumentTitle}' not in [{string.Join(", ", selectedHostFiles)}]\n"); } catch { }
                        }
                        return false;
                    }
                    afterHostFile++;
                    // ✅ DEPLOYMENT MODE: Skip file writes
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        try { File.AppendAllText(SafeFileLogger.GetLogFilePath("placement_debug.log"), $"[{DateTime.Now:HH:mm:ss}] ✅ Passed host file filter\n"); } catch { }
                    }
                    
                    // Filter 5: 3D section box filtering - DISABLED FOR DEBUG
                    // bool sectionBoxMatch = IsClashZoneVisibleInCurrentSectionBox(cz);
                    // if (!sectionBoxMatch)
                    // {
                    //     DebugLogger.Info($"{_logPrefix} Filtered out ClashZone {cz.Id}: Not visible in current 3D section box");
                    //     return false;
                    // }
                    
                    // 🚨 TEMPORARY: Only apply 3D section box filter (most likely to be correct)
                    // Section box filtering DISABLED for placement: we already have precise clash zones
                    bool sectionBoxMatch = true; // Default to true (section box filtering disabled)
                    if (OptimizationFlags.UseRefactoredCommandServices && _sectionBoxChecker != null)
                    {
                        // Use refactored service if flag enabled (but still default to true for now)
                        // sectionBoxMatch = _sectionBoxChecker.IsClashZoneVisibleInCurrentSectionBox(_doc, cz, _logPrefix);
                    }
                    else
                    {
                        // Legacy: sectionBoxMatch = IsClashZoneVisibleInCurrentSectionBox(cz);
                    }
                    if (!sectionBoxMatch)
                    {
                        DebugLogger.Info($"{_logPrefix} Filtered out ClashZone {cz.Id}: Not visible in current 3D section box");
                        return false;
                    }
                    afterSection++;
                    
                    return true;
                }).ToList();
                
                // ✅ CRITICAL: Log filter results to placement_debug.log
                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    var finalDebugPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                    try
                    {
                        File.AppendAllText(finalDebugPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 FILTERING COMPLETED: {clashZones.Count} -> {filteredZones.Count} clash zones\n");
                        File.AppendAllText(finalDebugPath, $"[{DateTime.Now:HH:mm:ss}] Filter breakdown: afterCategory={afterCategory}, afterHostType={afterHostType}, afterRefFile={afterRefFile}, afterHostFile={afterHostFile}, afterSection={afterSection}\n");
                        File.AppendAllText(finalDebugPath, $"[{DateTime.Now:HH:mm:ss}] UI Selections: HostTypes=[{string.Join(", ", selectedHostCategories)}], RefFiles=[{string.Join(", ", selectedReferenceFiles)}], HostFiles=[{string.Join(", ", selectedHostFiles)}]\n");
                    }
                    catch { }
                }
                
                // 🚨 DEBUG: Direct file logging to bypass DebugLogger issues
                DebugLogger.Info($"[{DateTime.Now}] 🚨 FILTERING COMPLETED: {clashZones.Count} -> {filteredZones.Count} clash zones\n");
                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    try
                    {
                        string placementFilterBreakdownLogPath = SafeFileLogger.GetLogFilePath("placement_filter_breakdown.log");
                        System.IO.File.AppendAllText(placementFilterBreakdownLogPath, $"[{DateTime.Now}] Counts: afterCategory={afterCategory}, afterHostType={afterHostType}, afterRefFile={afterRefFile}, afterHostFile={afterHostFile}, afterSection={afterSection}, final={filteredZones.Count}\n");
                    }
                    catch { }
                }
                
                DebugLogger.Info($"{_logPrefix} FINAL 5-FILTER SYSTEM: MEP category + host type + reference files + host files + 3D section box filters applied: {clashZones.Count} -> {filteredZones.Count} clash zones");
                return filteredZones;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"{_logPrefix} Error filtering clash zones: {ex.Message}");
                return clashZones; // Return all zones if filtering fails
            }
        }
        
        /// <summary>
        /// Check if clash zone is visible in current 3D section box
        /// </summary>
        private bool IsClashZoneVisibleInCurrentSectionBox(ClashZone clashZone)
        {
            try
            {
                // Check if we have an active 3D view with section box
                if (!(_doc.ActiveView is View3D view3D) || !view3D.IsSectionBoxActive)
                {
                    DebugLogger.Info($"{_logPrefix} No active 3D section box - ClashZone {clashZone.Id} considered visible");
                    return true; // No section box = all visible
                }

                // Get section box bounds
                var sectionBox = Helpers.SectionBoxHelper.GetSectionBoxBounds(view3D);
                if (sectionBox == null)
                {
                    DebugLogger.Info($"{_logPrefix} Could not get section box bounds - ClashZone {clashZone.Id} considered visible");
                    return true; // Can't get bounds = all visible
                }

                // Tolerance to avoid precision misses (in feet ~ 30mm)
                const double tol = 0.1;

                // Prefer checking the clash zone's intersection bounding box if present
                var czBox = clashZone.ClashBoundingBox;
                if (czBox != null)
                {
                    // Expand section box slightly by tol
                    bool overlaps =
                        (czBox.Min.X <= sectionBox.Max.X + tol) && (czBox.Max.X >= sectionBox.Min.X - tol) &&
                        (czBox.Min.Y <= sectionBox.Max.Y + tol) && (czBox.Max.Y >= sectionBox.Min.Y - tol) &&
                        (czBox.Min.Z <= sectionBox.Max.Z + tol) && (czBox.Max.Z >= sectionBox.Min.Z - tol);

                    DebugLogger.Info($"{_logPrefix} SectionBox test (BBox) CZ={clashZone.Id} overlaps={overlaps}  czMin=({czBox.Min.X:F2},{czBox.Min.Y:F2},{czBox.Min.Z:F2}) czMax=({czBox.Max.X:F2},{czBox.Max.Y:F2},{czBox.Max.Z:F2}) sbMin=({sectionBox.Min.X:F2},{sectionBox.Min.Y:F2},{sectionBox.Min.Z:F2}) sbMax=({sectionBox.Max.X:F2},{sectionBox.Max.Y:F2},{sectionBox.Max.Z:F2})");
                    return overlaps;
                }

                // Fallback to single point test with tolerance
                var p = clashZone.IntersectionPoint;
                bool inside =
                    p.X >= sectionBox.Min.X - tol && p.X <= sectionBox.Max.X + tol &&
                    p.Y >= sectionBox.Min.Y - tol && p.Y <= sectionBox.Max.Y + tol &&
                    p.Z >= sectionBox.Min.Z - tol && p.Z <= sectionBox.Max.Z + tol;

                DebugLogger.Info($"{_logPrefix} SectionBox test (Point) CZ={clashZone.Id} inside={inside} at ({p.X:F2}, {p.Y:F2}, {p.Z:F2}) sbMin=({sectionBox.Min.X:F2},{sectionBox.Min.Y:F2},{sectionBox.Min.Z:F2}) sbMax=({sectionBox.Max.X:F2},{sectionBox.Max.Y:F2},{sectionBox.Max.Z:F2})");
                return inside;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"{_logPrefix} Error checking ClashZone {clashZone.Id} section box visibility: {ex.Message} - considering visible");
                return true; // Consider visible on error
            }
        }
        
        /// <summary>
        /// Check if a file name is in the selected files list (handles parentheses and normalization)
        /// </summary>
        private bool IsFileInSelectedList(string fileName, List<string> selectedFiles)
        {
            if (string.IsNullOrEmpty(fileName) || selectedFiles.Count == 0)
            {
                DebugLogger.Info($"{_logPrefix} IsFileInSelectedList: No filtering - fileName='{fileName}', selectedFiles.Count={selectedFiles.Count}");
                return true; // No filtering if no selection
            }
            
            // Normalize file names for comparison (remove parentheses content)
            string normalizedFileName = NormalizeFileName(fileName);
            
            DebugLogger.Info($"{_logPrefix} IsFileInSelectedList: Checking fileName='{fileName}' -> normalized='{normalizedFileName}' against selectedFiles=[{string.Join(", ", selectedFiles)}]");
            
            // Log each selected file normalization for debugging
            foreach (var selectedFile in selectedFiles)
            {
                string normalizedSelectedFile = NormalizeFileName(selectedFile);
                DebugLogger.Info($"{_logPrefix} IsFileInSelectedList: Selected file '{selectedFile}' -> normalized='{normalizedSelectedFile}'");
            }
            
            foreach (var selectedFile in selectedFiles)
            {
                string normalizedSelectedFile = NormalizeFileName(selectedFile);
                DebugLogger.Info($"{_logPrefix} IsFileInSelectedList: Comparing '{normalizedFileName}' vs '{normalizedSelectedFile}'");
                
                if (string.Equals(normalizedFileName, normalizedSelectedFile, StringComparison.OrdinalIgnoreCase))
                {
                    DebugLogger.Info($"{_logPrefix} IsFileInSelectedList: ✅ MATCH FOUND: '{normalizedFileName}' == '{normalizedSelectedFile}'");
                    return true;
                }
            }
            
            DebugLogger.Info($"{_logPrefix} IsFileInSelectedList: ❌ NO MATCH: '{normalizedFileName}' not found in selected files");
            return false;
        }
        
        /// <summary>
        /// Normalize file name by removing parentheses content, extra text, and extracting filename from path
        /// </summary>
        private string NormalizeFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return string.Empty;
            
            // Extract filename from full path (e.g., "C:\Users\...\ME-00001.rvt" -> "ME-00001.rvt")
            fileName = Path.GetFileName(fileName);
            
            // Remove file extension (e.g., "ME-00001.rvt" -> "ME-00001")
            fileName = Path.GetFileNameWithoutExtension(fileName);
            
            // Remove content in parentheses (e.g., "Building (Architectural)" -> "Building")
            var idxParen = fileName.IndexOf('(');
            if (idxParen >= 0)
            {
                fileName = fileName.Substring(0, idxParen).Trim();
            }
            
            // Remove element count suffixes (e.g., "ME-00001 (118 elements)" -> "ME-00001")
            // This handles cases where UI shows "ME-00001 (118 elements)" but clash zones store "ME-00001"
            var idxElements = fileName.IndexOf(" elements");
            if (idxElements >= 0)
            {
                fileName = fileName.Substring(0, idxElements).Trim();
            }
            
            // Remove any trailing spaces and return
            return fileName.Trim();
        }
        
        /// <summary>
        /// ✅ STANDARDIZED: Normalize category names to match clash zone file naming
        /// Converts "Duct Accessories" → "duct_accessories", "Ducts" → "ducts", etc.
        /// </summary>
        private string NormalizeCategoryName(string categoryName)
        {
            if (string.IsNullOrEmpty(categoryName))
                return "unknown";
                
            return categoryName
                .Replace(" ", "_")           // "Duct Accessories" → "Duct_Accessories"
                .ToLowerInvariant();          // "Duct_Accessories" → "duct_accessories"
        }

        /// <summary>
        /// Get selected filter names from UI state
        /// </summary>
        private List<string> GetSelectedFilterNamesFromUI()
        {
            try
            {
                // Use FilterUiStateProvider to get current UI state
                var selectedFilters = FilterUiStateProvider.GetSelectedFilterItems?.Invoke() ?? new List<string>();
                DebugLogger.Info($"{_logPrefix} UI selected filter names: {string.Join(", ", selectedFilters)}");
                return selectedFilters;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"{_logPrefix} Error getting filter names from UI: {ex.Message}");
                return new List<string>();
            }
        }
    }

    /// <summary>
    /// Failure preprocessor to auto-dismiss warnings during universal sleeve placement
    /// </summary>
    public class UniversalWarningSwallower : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(FailuresAccessor fa)
        {
            var failures = fa.GetFailureMessages();
            foreach (var f in failures)
            {
                var description = f.GetDescriptionText();
                if (f.GetSeverity() == FailureSeverity.Warning)
                {
                    fa.DeleteWarning(f);
                    DebugLogger.Info($"[UniversalWarningSwallower] Dismissed warning: {description}");
                }
                // Also dismiss duplicate-related errors to prevent transaction rollback
                else if (f.GetSeverity() == FailureSeverity.Error && 
                         (description.Contains("duplicate") || description.Contains("already exists")))
                {
                    fa.DeleteWarning(f);
                    DebugLogger.Info($"[UniversalWarningSwallower] Dismissed duplicate error: {description}");
                }
            }
            return FailureProcessingResult.Continue;
        }
    }
}
