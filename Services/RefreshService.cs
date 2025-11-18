using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service responsible for handling refresh operations and clash detection
    /// </summary>
    public class RefreshService
    {
        // Static serializers to avoid repeated XmlSerializer type metadata allocations per refresh
        private static readonly System.Xml.Serialization.XmlSerializer OpeningFilterSerializer =
            new System.Xml.Serialization.XmlSerializer(typeof(Models.OpeningFilter));
        // ⚠️ CRASH-SAFE LIMITS ⚠️
        private const int MAX_ELEMENTS_TO_PROCESS = 10000;
        private const int WARNING_THRESHOLD = 5000;
        private const int TIMEOUT_CHECK_INTERVAL = 100; // Check timeout every N elements

        private readonly Document _document;
        private readonly UIDocument _uiDocument;
        private readonly ApplicationProfileService _appProfileService;
        private readonly FilterManagementService _filterManagementService;
        private ClashZoneService _clashZoneService; // Not readonly - reinitialized during refresh with existing zones
        private readonly IntersectionDetectionService _intersectionService;
        private readonly CrashSafeExecutor _crashSafeExecutor; // ⚠️ CRITICAL: Provides timeout and crash protection
        private MemoryManager _memoryManager; // ⚠️ CRITICAL: Provides memory management and timeout protection
        private MemoryProfiler _memoryProfiler; // Memory profiling for tracking actual vs theoretical usage

        // ✅ OOP REFACTORING: Centralized flag and GUID management
        private readonly FlagManager _flagManager;
        private readonly GuidManager _guidManager;

        // ✅ OOP REFACTORING: Validation strategies
        private readonly Validation.ThreePointValidator _threePointValidator;
        private readonly ClashZonePairMatcher _pairMatcher; // ✅ OOP: Helper for pair matching

        // UI References (passed from main dialog)
        private System.Windows.Forms.Label _statusLabel;
        private System.Windows.Forms.ProgressBar _progressBar;
        private System.Windows.Forms.Button _refreshButton;

        public RefreshService(Document document, UIDocument uiDocument, ApplicationProfileService appProfileService)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _uiDocument = uiDocument ?? throw new ArgumentNullException(nameof(uiDocument));
            _appProfileService = appProfileService ?? throw new ArgumentNullException(nameof(appProfileService));
            _filterManagementService = new FilterManagementService(_document, msg =>
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info(msg);
            }, msg =>
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error(msg);
            });

            // ⚠️ CRITICAL FIX: Initialize with EMPTY storage in constructor ⚠️
            // The existing clash zones will be loaded during ExecuteRefresh from the current profile
            _clashZoneService = new ClashZoneService(new Models.ClashZoneStorage(), msg =>
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info(msg);
            });
            _intersectionService = new IntersectionDetectionService(msg =>
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info(msg);
            });

            // Initialize crash-safe executor for timeout protection
            _crashSafeExecutor = new CrashSafeExecutor();

            // ⚠️ CRITICAL: Initialize memory manager for large/unclean files (5 min timeout)
            // Memory limit is auto-calculated based on system RAM (10% of total RAM, capped at 12GB)
            // For 64GB system: ~6.4GB limit (10% of 64GB, capped at 12GB)
            _memoryManager = new MemoryManager(maxMemoryMB: null, timeoutMinutes: 5);

            // ✅ OOP REFACTORING: Initialize centralized managers
            _flagManager = new FlagManager(document);
            _guidManager = new GuidManager(document);
            _threePointValidator = new Validation.ThreePointValidator();
            _pairMatcher = new ClashZonePairMatcher(_document); // ✅ OOP: Initialize pair matcher
        }

        /// <summary>
        /// Sets UI references from the main dialog
        /// </summary>
        public void SetUIReferences(System.Windows.Forms.Label statusLabel, System.Windows.Forms.ProgressBar progressBar, System.Windows.Forms.Button refreshButton)
        {
            _statusLabel = statusLabel;
            _progressBar = progressBar;
            _refreshButton = refreshButton;
        }

        /// <summary>
        /// 🔥 CRITICAL FIX: Load existing clash zone data to preserve cluster information
        /// This prevents the Refresh process from overwriting existing cluster data
        /// </summary>
        public void LoadExistingClashZoneData()
        {
            try
            {
                var filtersDirectory = (_document != null)
                    ? ProjectPathService.GetFiltersDirectory(_document)
                    : ProjectPathService.GetFiltersDirectory(_document);

                if (!Directory.Exists(filtersDirectory))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info("[RefreshService] No existing filters directory found - starting with empty clash zones");
                    return;
                }

                // Look for existing XML files with clash zone data
                var xmlFiles = Directory.GetFiles(filtersDirectory, "*.xml");

                if (xmlFiles.Length == 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info("[RefreshService] No existing XML files found - starting with empty clash zones");
                    return;
                }

                // Load clash zones from the most recently modified XML file
                var mostRecentFile = xmlFiles
                    .OrderByDescending(f => File.GetLastWriteTime(f))
                    .First();

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[RefreshService] Loading existing clash zone data from: {mostRecentFile}");

                // Reuse a static serializer to avoid per-refresh type cache allocations
                var serializer = OpeningFilterSerializer;
                Models.OpeningFilter filter = null;
                using (var reader = new StreamReader(mostRecentFile))
                {
                    filter = (Models.OpeningFilter)serializer.Deserialize(reader);
                }

                // ✅ PHASE SQLITE-2: Load from SQLite FIRST (primary source), XML as fallback
                if (DeploymentConfiguration.UseSqliteAsPrimary)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[RefreshService] PHASE 2: Loading from SQLite (PRIMARY) for filter '{filter?.Name ?? "Unknown"}'");

                    try
                    {
                        // Extract categories from filter or load all categories from SQLite for this filter
                        var categoriesToLoad = filter?.SelectedMepCategoryNames?.ToList() ?? new List<string>();
                        
                        // If no categories in filter, try common categories
                        if (categoriesToLoad.Count == 0)
                        {
                            categoriesToLoad = new List<string> { "Ducts", "Pipes", "Cable Trays", "Duct Accessories" };
                        }

                        var sqliteZones = LoadClashZonesFromSqliteFallback(filter, categoriesToLoad);
                        if (sqliteZones != null && sqliteZones.Count > 0)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[RefreshService] ✅ PHASE 2: Loaded {sqliteZones.Count} clash zones from SQLite (PRIMARY)");

                            // Create ClashZoneStorage from SQLite zones
                            if (filter == null)
                            {
                                filter = new Models.OpeningFilter
                                {
                                    Name = Path.GetFileNameWithoutExtension(mostRecentFile),
                                    ClashZoneStorage = new Models.ClashZoneStorage()
                                };
                            }

                            filter.ClashZoneStorage ??= new Models.ClashZoneStorage();
                            filter.ClashZoneStorage.ClashZones = sqliteZones;
                        }
                        else
                        {
                            // SQLite has no zones - fallback to XML if available
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[RefreshService] PHASE 2: SQLite has no zones, falling back to XML");
                        }
                    }
                    catch (Exception sqliteEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[RefreshService] PHASE 2: SQLite load failed, falling back to XML: {sqliteEx.Message}");
                        // Fall through to use XML as fallback
                    }
                }
                else
                {
                    // Legacy mode: XML is primary, SQLite is fallback
                    if (filter?.ClashZoneStorage?.AllZones == null || filter.ClashZoneStorage.AllZones.Count == 0)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[RefreshService] XML file has no zones, attempting SQLite fallback for filter '{filter?.Name ?? "Unknown"}'");

                        try
                        {
                            var categoriesToLoad = filter?.SelectedMepCategoryNames?.ToList() ?? new List<string>();
                            if (categoriesToLoad.Count == 0)
                            {
                                categoriesToLoad = new List<string> { "Ducts", "Pipes", "Cable Trays", "Duct Accessories" };
                            }

                            var sqliteZones = LoadClashZonesFromSqliteFallback(filter, categoriesToLoad);
                            if (sqliteZones != null && sqliteZones.Count > 0)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[RefreshService] ✅ SQLite fallback loaded {sqliteZones.Count} clash zones");

                                if (filter == null)
                                {
                                    filter = new Models.OpeningFilter
                                    {
                                        Name = Path.GetFileNameWithoutExtension(mostRecentFile),
                                        ClashZoneStorage = new Models.ClashZoneStorage()
                                    };
                                }

                                filter.ClashZoneStorage ??= new Models.ClashZoneStorage();
                                filter.ClashZoneStorage.ClashZones = sqliteZones;
                            }
                        }
                        catch (Exception sqliteEx)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[RefreshService] SQLite fallback failed: {sqliteEx.Message}");
                        }
                    }
                }

                if (filter?.ClashZoneStorage?.AllZones != null && filter.ClashZoneStorage.AllZones.Count > 0)
                {
                    // ✅ GLOBAL XML FLAG MANAGEMENT: Sync flags from Global XML immediately after loading Filter XML
                    // Flags are NOT persisted to Filter XML - must sync from Global XML (single source of truth)
                    try
                    {
                        var flagManager = new FlagManager(_document);
                        var clashZonesByCategory = filter.ClashZoneStorage.AllZones
                            .GroupBy(cz => cz.MepElementCategory)
                            .ToList();

                        foreach (var categoryGroup in clashZonesByCategory)
                        {
                            var category = categoryGroup.Key;
                            var categoryClashZones = categoryGroup.ToList();

                            if (!string.IsNullOrWhiteSpace(category) && categoryClashZones.Count > 0)
                            {
                                flagManager.SyncFlagsFromGlobal(categoryClashZones, category);
                            }
                        }
                    }
                    catch (Exception flagEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[RefreshService] Failed to sync flags from Global XML: {flagEx.Message}");
                    }

                    // ✅ CRITICAL: Preserve existing clash zone data including cluster information
                    _clashZoneService = new ClashZoneService(filter.ClashZoneStorage, msg =>
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info(msg);
                    });

                        int clusterResolvedCount = filter.ClashZoneStorage.AllZones.Count(cz => cz.IsClusterResolved);
                        int individualResolvedCount = filter.ClashZoneStorage.AllZones.Count(cz => cz.IsResolved);

                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[RefreshService] ✅ Loaded existing clash zone data: {filter.ClashZoneStorage.AllZones.Count} total, {clusterResolvedCount} cluster-resolved, {individualResolvedCount} individual-resolved");
                    }
                    else
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info("[RefreshService] No clash zone data found in existing XML file");
                    }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[RefreshService] Error loading existing clash zone data: {ex.Message}");
                // Continue with empty storage if loading fails
            }
        }

        /// <summary>
        /// ✅ SQLITE FALLBACK: Load clash zones from SQLite database when XML is empty or missing
        /// </summary>
        private List<Models.ClashZone> LoadClashZonesFromSqliteFallback(Models.OpeningFilter filter, List<string> categoriesToLoad)
        {
            var allZones = new List<Models.ClashZone>();

            if (filter == null || string.IsNullOrWhiteSpace(filter.Name))
                return allZones;

            // If no categories specified, try common categories
            if (categoriesToLoad == null || categoriesToLoad.Count == 0)
            {
                categoriesToLoad = new List<string> { "Ducts", "Pipes", "Cable Trays", "Duct Accessories" };
            }

            try
            {
                using (var context = new Data.SleeveDbContext(_document, msg =>
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[RefreshService][SQLite] {msg}");
                }))
                {
                    var repository = new Data.Repositories.ClashZoneRepository(context, msg =>
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[RefreshService][SQLite] {msg}");
                    });

                    // Load zones for each category
                    foreach (var category in categoriesToLoad)
                    {
                        if (string.IsNullOrWhiteSpace(category))
                            continue;

                        try
                        {
                            // Load all zones (not just unresolved) for refresh operations
                            var categoryZones = repository.GetClashZonesByFilter(filter.Name, category, unresolvedOnly: false) ?? new List<Models.ClashZone>();
                            
                            foreach (var zone in categoryZones)
                            {
                                if (zone != null)
                                {
                                    zone.EnsureSleevePlacementPointReconstructed();
                                    zone.EnsureSleevePlacementPointActiveDocumentReconstructed();
                                    allZones.Add(zone);
                                }
                            }

                            if (categoryZones.Count > 0 && !DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[RefreshService] ✅ SQLite fallback loaded {categoryZones.Count} zones for filter '{filter.Name}', category '{category}'");
                            }
                        }
                        catch (Exception categoryEx)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[RefreshService] SQLite fallback failed for category '{category}': {categoryEx.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[RefreshService] SQLite fallback failed: {ex.Message}");
            }

            return allZones;
        }

        /// <summary>
        /// Ensure shared parameters are loaded into the project
        /// </summary>
        private void EnsureSharedParametersLoaded(Document doc)
        {
            try
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info("[SHARED_PARAMS] Ensuring shared parameters are loaded into project");

                // Path to shared parameter file
                string sharedParamFile = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Resources\Opening family shared parameter.txt";

                if (!File.Exists(sharedParamFile))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[SHARED_PARAMS] Shared parameter file not found: {sharedParamFile}");
                    return;
                }

                // Check if shared parameter file is already loaded
                var currentSharedParams = doc.Application.SharedParametersFilename;
                if (!string.IsNullOrEmpty(currentSharedParams) && currentSharedParams.Contains("Opening family shared parameter.txt"))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[SHARED_PARAMS] Shared parameters already loaded: {currentSharedParams}");
                    return;
                }

                // Load the shared parameter file
                doc.Application.SharedParametersFilename = sharedParamFile;
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[SHARED_PARAMS] Loaded shared parameter file: {sharedParamFile}");

                // Create a group for the parameters if it doesn't exist
                var groupName = "Openings";
                var sharedParams = doc.Application.OpenSharedParameterFile();

                if (sharedParams != null)
                {
                    var group = sharedParams.Groups.get_Item(groupName);
                    if (group == null)
                    {
                        // Create the group if it doesn't exist
                        group = sharedParams.Groups.Create(groupName);
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[SHARED_PARAMS] Created shared parameter group: {groupName}");
                    }
                }

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info("[SHARED_PARAMS] Shared parameters loaded successfully");
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[SHARED_PARAMS] Error loading shared parameters: {ex.Message}");
                // Continue anyway - parameters might already be loaded
            }
        }

        /// <summary>
        /// Validates that dampers in the selected linked mechanical file have required Standard and MSFD parameters
        /// </summary>
        private bool ValidateDamperParameters(List<string> selectedReferenceFiles, string refreshLogName)
        {
            try
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info("[DUCT_ACCESSORIES] Starting damper parameter validation");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DUCT_ACCESSORIES] Starting damper parameter validation\n");

                if (selectedReferenceFiles == null || selectedReferenceFiles.Count == 0)
                {
                    var result = System.Windows.Forms.MessageBox.Show(
                        "No linked mechanical files selected. Please select a linked mechanical file to validate damper parameters.",
                        "No Linked Files Selected",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);

                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning("[DUCT_ACCESSORIES] No linked files selected for damper validation");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DUCT_ACCESSORIES] WARNING: No linked files selected for damper validation\n");
                    return false;
                }

                // Find linked mechanical file
                Document linkedMechanicalDoc = null;
                foreach (var linkInstance in new FilteredElementCollector(_document)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>())
                {
                    var linkDoc = linkInstance.GetLinkDocument();
                    if (linkDoc != null && selectedReferenceFiles.Any(file => linkDoc.Title.Contains(file) || file.Contains(linkDoc.Title)))
                    {
                        // Check if this looks like a mechanical file (contains duct accessories)
                        var ductAccessories = new FilteredElementCollector(linkDoc)
                            .OfCategory(BuiltInCategory.OST_DuctAccessory)
                            .WhereElementIsNotElementType()
                            .ToElements();

                        if (ductAccessories.Count > 0)
                        {
                            linkedMechanicalDoc = linkDoc;
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[DUCT_ACCESSORIES] Found linked mechanical file: {linkDoc.Title}");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DUCT_ACCESSORIES] Found linked mechanical file: {linkDoc.Title}\n");
                            break;
                        }
                    }
                }

                if (linkedMechanicalDoc == null)
                {
                    var result = System.Windows.Forms.MessageBox.Show(
                        "No linked mechanical file found with duct accessories. Please ensure a mechanical file is linked and selected.",
                        "No Mechanical File Found",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);

                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning("[DUCT_ACCESSORIES] No linked mechanical file found with duct accessories");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DUCT_ACCESSORIES] WARNING: No linked mechanical file found with duct accessories\n");
                    return false;
                }

                // Get all duct accessories (dampers) from the linked mechanical file
                var dampers = new FilteredElementCollector(linkedMechanicalDoc)
                    .OfCategory(BuiltInCategory.OST_DuctAccessory)
                    .WhereElementIsNotElementType()
                    .ToElements();

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[DUCT_ACCESSORIES] Found {dampers.Count} dampers in linked mechanical file");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DUCT_ACCESSORIES] Found {dampers.Count} dampers in linked mechanical file\n");

                if (dampers.Count == 0)
                {
                    var result = System.Windows.Forms.MessageBox.Show(
                        "No dampers found in the linked mechanical file. Please ensure the mechanical file contains duct accessories (dampers).",
                        "No Dampers Found",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);

                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning("[DUCT_ACCESSORIES] No dampers found in linked mechanical file");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DUCT_ACCESSORIES] WARNING: No dampers found in linked mechanical file\n");
                    return false;
                }

                // Check each damper for required parameters
                var missingStandard = new List<Element>();
                var missingMSFD = new List<Element>();

                foreach (var damper in dampers)
                {
                    if (!(damper is FamilyInstance fi))
                    {
                        continue;
                    }

                    // ✅ CORRECT: Check family name for damper type indicators (not parameters)
                    string familyName = fi.Symbol?.Family?.Name ?? "";
                    string typeName = fi.Symbol?.Name ?? "";
                    string nameToCheck = familyName + " " + typeName; // Check both family and type names
                    string nameUpper = nameToCheck.Trim().ToUpperInvariant();

                    // ✅ DEBUG: Log damper family for first 5 dampers
                    if (missingStandard.Count + missingMSFD.Count < 5)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[DUCT_ACCESSORIES] Damper {damper.Id}: Family='{familyName}', Type='{typeName}', Combined='{nameUpper}'");
                    }

                    // Check if family/type name contains damper type keywords (case insensitive)
                    bool hasMSFD = nameUpper.Contains("MSFD");
                    bool hasStandard = nameUpper.Contains("STANDARD") || nameUpper.Contains("MSD") || nameUpper.Contains("MD") || nameUpper.Contains("MOTORIZED");

                    if (!hasStandard)
                    {
                        missingStandard.Add(damper);
                    }

                    if (!hasMSFD)
                    {
                        missingMSFD.Add(damper);
                    }
                }

                // Report results
                if (missingStandard.Count > 0 || missingMSFD.Count > 0)
                {
                    var message = "The following damper family name validation issues were found:\n\n";

                    if (missingStandard.Count > 0)
                    {
                        message += $"• {missingStandard.Count} dampers missing 'Standard' keyword in family name\n";
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[DUCT_ACCESSORIES] {missingStandard.Count} dampers missing 'Standard' keyword in family name");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DUCT_ACCESSORIES] WARNING: {missingStandard.Count} dampers missing 'Standard' keyword in family name\n");
                    }

                    if (missingMSFD.Count > 0)
                    {
                        message += $"• {missingMSFD.Count} dampers missing 'MSFD' keyword in family name\n";
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[DUCT_ACCESSORIES] {missingMSFD.Count} dampers missing 'MSFD' keyword in family name");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DUCT_ACCESSORIES] WARNING: {missingMSFD.Count} dampers missing 'MSFD' keyword in family name\n");
                    }

                    message += "\nPlease update the damper family names in the linked mechanical file to include 'MSFD' or 'Standard' keywords before proceeding with intersection detection.";
                    message += "\n\nDo you want to continue anyway?";

                    var result = System.Windows.Forms.MessageBox.Show(
                        message,
                        "Damper Family Name Validation Failed",
                        System.Windows.Forms.MessageBoxButtons.YesNo,
                        System.Windows.Forms.MessageBoxIcon.Warning);

                    if (result == System.Windows.Forms.DialogResult.No)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info("[DUCT_ACCESSORIES] User chose to abort due to missing damper parameters");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DUCT_ACCESSORIES] User chose to abort due to missing damper parameters\n");
                        return false;
                    }
                    else
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info("[DUCT_ACCESSORIES] User chose to continue despite missing damper parameters");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DUCT_ACCESSORIES] User chose to continue despite missing damper parameters\n");
                    }
                }
                else
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info("[DUCT_ACCESSORIES] All dampers have required Standard and MSFD parameters");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DUCT_ACCESSORIES] All dampers have required Standard and MSFD parameters\n");
                }

                return true;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[DUCT_ACCESSORIES] Error during damper parameter validation: {ex.Message}");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DUCT_ACCESSORIES] ERROR: {ex.Message}\n");

                var result = System.Windows.Forms.MessageBox.Show(
                    $"Error during damper parameter validation: {ex.Message}\n\nDo you want to continue anyway?",
                    "Validation Error",
                    System.Windows.Forms.MessageBoxButtons.YesNo,
                    System.Windows.Forms.MessageBoxIcon.Error);

                return result == System.Windows.Forms.DialogResult.Yes;
            }
        }

        /// <summary>
        /// Main refresh execution method - extracted from EmergencyMainDialog
        /// ⚠️ CRITICAL: Wrapped with crash-safe executor for timeout/memory protection
        /// </summary>
        public void ExecuteRefresh(List<string> selectedFilterItems, List<string> selectedMepCategories, List<string> selectedReferenceFiles, List<string> selectedHostFiles, Dictionary<string, double> clearanceSettings)
        {
            // ⚠️ CRITICAL: Wrap entire refresh operation with crash-safe executor
            // This ensures 5-minute timeout, graceful error handling, and prevents crashes
            var result = _crashSafeExecutor.ExecuteWithTimeout(() =>
            {
                return ExecuteRefreshInternal(selectedFilterItems, selectedMepCategories, selectedReferenceFiles, selectedHostFiles, clearanceSettings);
            }, "Refresh Operation");

            // Always ensure UI is reset even if operation failed/cancelled
            if (result != Autodesk.Revit.UI.Result.Succeeded)
            {
                try
                {
                    _progressBar.Visible = false;
                    _refreshButton.Enabled = true;
                    if (result == Autodesk.Revit.UI.Result.Failed)
                    {
                        _statusLabel.Text = "Refresh operation failed. Check logs for details.";
                    }
                    else if (result == Autodesk.Revit.UI.Result.Cancelled)
                    {
                        _statusLabel.Text = "Refresh operation cancelled due to timeout or resource limits.";
                    }
                }
                catch { } // Fail silently on UI updates
            }
        }
        /// <summary>
        /// Internal refresh implementation - called by crash-safe wrapper
        /// </summary>
        private Autodesk.Revit.UI.Result ExecuteRefreshInternal(List<string> selectedFilterItems, List<string> selectedMepCategories, List<string> selectedReferenceFiles, List<string> selectedHostFiles, Dictionary<string, double> clearanceSettings)
        {
            // ✅ METHOD-LEVEL: Load settings once at the start - reuse throughout method
            var settings = _appProfileService?.GetCurrentSettings();
            bool enableThreePointValidation = settings?.EnableThreePointValidation ?? true;

            // ✅ FIX 2: Clear geometry cache BEFORE starting to prevent memory accumulation
            MepIntersectionService.ClearGeometryCache();
            MepIntersectionService.ClearTransformCache();

            try
            {
                // ✅ MEMORY OPTIMIZATION: Declare batched logger early for proper scope management
                BatchedLogger batchedLogger = null;
                // ✅ CRASH-SAFE: Create timestamped refresh log file using SafeFileLogger
                // SafeFileLogger automatically creates directories and handles missing paths gracefully
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                string refreshLogName = $"Refresh_{timestamp}.log";
                var __refreshStart = DateTime.Now;

                // ✅ INVESTIGATION: Get log directory info for debugging
                string actualLogDir = SafeFileLogger.GetLogDirectory();
                string fullRefreshLogPath = SafeFileLogger.GetLogFilePath(refreshLogName);

                // ✅ INVESTIGATION: Log to SafeFileLogger to see if it's working
                SafeFileLogger.SafeAppendText(refreshLogName, $"=== REFRESH METHOD STARTED ===\n");
                SafeFileLogger.SafeAppendText(refreshLogName, $"Log file: {fullRefreshLogPath}\n");
                SafeFileLogger.SafeAppendText(refreshLogName, $"Log directory: {actualLogDir}\n");
                SafeFileLogger.SafeAppendText(refreshLogName, $"DeploymentMode={DeploymentConfiguration.DeploymentMode}, DiagnosticMode={OptimizationFlags.UseDiagnosticMode}\n");

                // ✅ DIAGNOSTIC: Write all diagnostic info to refresh log file
                SafeFileLogger.SafeAppendText(refreshLogName, $"[REFRESH] === REFRESH METHOD STARTED AT {DateTime.Now} ===\n");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[REFRESH] Timestamp: {timestamp}\n");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[REFRESH] Log file will be: {refreshLogName}\n");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[REFRESH] Full log path: {fullRefreshLogPath}\n");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[REFRESH] Log directory: {actualLogDir}\n");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[REFRESH] DeploymentMode={DeploymentConfiguration.DeploymentMode}\n");

                // ✅ CRASH-SAFE: Use SafeFileLogger - no need for manual directory creation or try-catch
                SafeFileLogger.SafeAppendText("performance.log", $"REFRESH_START {__refreshStart:O}");

                // Local memory diagnostics helpers (scoped to this refresh run)
                long __lastManagedBytes = 0;
                Action<string, IList<Models.ClashZone>, IList<Models.ClashZone>, IList<Models.ClashZone>> __logPhaseMem = (phase, listA, listB, listC) =>
                {
                    try
                    {
                        GC.Collect(2, GCCollectionMode.Forced, true);
                        GC.WaitForPendingFinalizers();
                        GC.Collect(2, GCCollectionMode.Forced, true);
                        long managed = GC.GetTotalMemory(false);

                        int a = listA?.Count ?? 0;
                        int b = listB?.Count ?? 0;
                        int c = listC?.Count ?? 0;
                        int totalListsCount = a + b + c;

                        // Duplication analysis by Guid
                        var unique = new HashSet<Guid>();
                        if (listA != null) foreach (var x in listA) if (x != null) unique.Add(x.Id);
                        if (listB != null) foreach (var x in listB) if (x != null) unique.Add(x.Id);
                        if (listC != null) foreach (var x in listC) if (x != null) unique.Add(x.Id);
                        int uniqueCount = unique.Count;
                        double dupFactor = uniqueCount == 0 ? 0.0 : (double)totalListsCount / uniqueCount;

                        SafeFileLogger.SafeAppendText(refreshLogName,
                            $"[{DateTime.Now}] [MEMORY_DEBUG] PHASE={phase}: managed={managed / 1024.0 / 1024.0:F2} MB, counts: A={a}, B={b}, C={c}, totalEntries={totalListsCount}, uniqueZones={uniqueCount}, duplicationFactor={dupFactor:F2}\n");

                        if (uniqueCount > 0)
                        {
                            double kbPerZone = (managed - __lastManagedBytes) / 1024.0 / uniqueCount;
                            SafeFileLogger.SafeAppendText(refreshLogName,
                                $"[{DateTime.Now}] [MEMORY_DEBUG] PHASE={phase}: delta since last phase = {(managed - __lastManagedBytes) / 1024.0:F1} KB total, ≈ {kbPerZone:F1} KB/zone\n");
                        }
                        __lastManagedBytes = managed;
                    }
                    catch { /* non-fatal diagnostics */ }
                };
                // Initial memory baseline
                __logPhaseMem("START", null, null, null);

                // Write to the main debug logger file that user can see
                // ✅ DEPLOYMENT MODE: ConditionalAppendAllText now checks DeploymentMode automatically
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("logger_debug.txt"), $"[{DateTime.Now}] === REFRESH STARTED === Timestamp: {timestamp}\n");

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info("=== REFRESH METHOD STARTED ===");
                SafeFileLogger.SafeAppendText(refreshLogName, "[REFRESH] DebugLogger.Info called\n");

                // ⚠️ CRITICAL: Reset memory manager timer for this refresh operation
                if (_memoryManager != null)
                {
                    // Dispose old instance and create new one for this refresh
                    _memoryManager?.Dispose();
                    _memoryManager = new MemoryManager(maxMemoryMB: null, timeoutMinutes: 5); // Auto-detect optimal limit
                                                                                              // Log after initialization so we can access CurrentMemoryMB
                    if (_memoryManager != null)
                    {
                        SafeFileLogger.SafeAppendText("refresh_memory.log",
                            $"Memory manager reset for new refresh operation. Limit: {_memoryManager.CurrentMemoryMB / 1024}GB (auto-calculated from system RAM), Timeout: 5 minutes");
                    }
                }

                // ✅ CRITICAL: Force AppData Logs directory creation for deployment scenarios
                // This ensures the directory exists even if project directory takes priority
                SafeFileLogger.EnsureAppDataLogsDirectory();

                // ✅ TEST: Write a test file to AppData Logs directory to verify it's writable
                string appDataLogsPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "JSE_MEP_Openings",
                    "Logs"
                );
                try
                {
                    if (!Directory.Exists(appDataLogsPath))
                    {
                        Directory.CreateDirectory(appDataLogsPath);
                    }
                    string testFile = Path.Combine(appDataLogsPath, "test_write_check.txt");
                    File.WriteAllText(testFile, $"Test write successful at {DateTime.Now:yyyy-MM-dd HH:mm:ss}\n");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[REFRESH] ✅ Test file written to: {testFile}\n");
                }
                catch (Exception testEx)
                {
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[REFRESH] ⚠️ Failed to write test file to AppData Logs: {testEx.Message}\n");
                }

                // Log the actual log directory being used (for troubleshooting)
                // ✅ NOTE: actualLogDir already set above, but refresh it in case directory changed
                actualLogDir = SafeFileLogger.GetLogDirectory();
                string logDirInfo = SafeFileLogger.GetLogDirectoryInfo();

                // ✅ DIAGNOSTIC: Log directory info and write first diagnostic entry
                SafeFileLogger.SafeAppendText(refreshLogName, $"═══ LOG DIRECTORY INFO ═══\n");
                SafeFileLogger.SafeAppendText(refreshLogName, $"{logDirInfo}\n");
                SafeFileLogger.SafeAppendText(refreshLogName, $"Memory profiling log: refresh_memory_profiling_{timestamp}.log\n");
                SafeFileLogger.SafeAppendText(refreshLogName, $"All logs will be written to: {actualLogDir}\n");
                SafeFileLogger.SafeAppendText(refreshLogName, $"Full refresh log path: {fullRefreshLogPath}\n");
                SafeFileLogger.SafeAppendText(refreshLogName, $"════════════════════════════════════\n");

                if (DeploymentConfiguration.EnableGlobalIndexDedupe &&
                    selectedMepCategories != null &&
                    selectedMepCategories.Count > 0)
                {
                    var dedupeLogBuilder = new StringBuilder();
                    var categoriesToDedupe = selectedMepCategories
                        .Where(cat => !string.IsNullOrWhiteSpace(cat))
                        .Select(cat => cat.Trim())
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    foreach (var category in categoriesToDedupe)
                    {
                        try
                        {
                            dedupeLogBuilder.AppendLine($"[{DateTime.Now}] [DEDUPER] Category='{category}' DryRun={DeploymentConfiguration.GlobalIndexDedupeDryRun}");
                            var dedupeResult = GlobalIndexMaintenance.Deduplicate(
                                _document,
                                category,
                                DeploymentConfiguration.GlobalIndexDedupeDryRun,
                                dedupeLogBuilder);

                            dedupeLogBuilder.AppendLine(
                                $"[{DateTime.Now}] [DEDUPER]   FiltersRemoved={dedupeResult.FiltersRemoved}, CombosMerged={dedupeResult.CombosMerged}, EntriesMerged={dedupeResult.EntriesMerged}, Applied={dedupeResult.ChangesApplied}, Error='{dedupeResult.ErrorMessage}'");
                        }
                        catch (Exception dedupeEx)
                        {
                            dedupeLogBuilder.AppendLine($"[{DateTime.Now}] [DEDUPER] ERROR for category '{category}': {dedupeEx}");
                        }
                    }

                    if (dedupeLogBuilder.Length > 0)
                    {
                        SafeFileLogger.SafeAppendText(refreshLogName, dedupeLogBuilder.ToString());
                        SafeFileLogger.SafeAppendText("filters_dedupe.log", dedupeLogBuilder.ToString());
                    }
                }

                // ✅ DIAGNOSTIC: Write diagnostic info to refresh log file
                SafeFileLogger.SafeAppendText(refreshLogName, $"[REFRESH] Log directory: {actualLogDir}\n");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[REFRESH] Full log path: {fullRefreshLogPath}\n");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[REFRESH] File exists after write: {File.Exists(fullRefreshLogPath)}\n");

                // ✅ DIAGNOSTIC: Verify file was created by reading it back
                try
                {
                    if (File.Exists(fullRefreshLogPath))
                    {
                        var content = File.ReadAllText(fullRefreshLogPath);
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[REFRESH] ✅ Log file exists! Size: {content.Length} bytes, First 200 chars: {content.Substring(0, Math.Min(200, content.Length))}\n");
                    }
                    else
                    {
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[REFRESH] ❌ Log file does NOT exist: {fullRefreshLogPath}\n");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[REFRESH] Directory exists: {Directory.Exists(actualLogDir)}\n");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[REFRESH] DeploymentMode: {DeploymentConfiguration.DeploymentMode}\n");
                    }
                }
                catch (Exception diagEx)
                {
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[REFRESH] ❌ Error checking log file: {diagEx.Message}\n");
                }
                // Extra memory snapshot to pinpoint early allocations
                try { _memoryProfiler.TakeSnapshot("AFTER_LOG_INFO", 0); } catch { }

                // ✅ REMOVED: Log directory prompt (user knows where logs are now)
                // Log directory info is still written to refresh log file for reference

                // Show user-friendly message in status label
                _statusLabel.Text = "Starting refresh...";

                // ✅ MEMORY PROFILING: Initialize profiler to track actual memory usage per clash zone
                // ✅ FIX: Initialize profiler even when deployment mode is OFF (for diagnostic purposes)
                _memoryProfiler = new MemoryProfiler($"refresh_memory_profiling_{timestamp}.log");
                _memoryProfiler.TakeSnapshot("REFRESH_START", 0);

                // ✅ INVESTIGATION: Log profiler initialization using SafeFileLogger
                SafeFileLogger.SafeAppendText(refreshLogName, $"[MEMORY_PROFILER] ✅ Memory profiler initialized: refresh_memory_profiling_{timestamp}.log\n");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[MEMORY_PROFILER] DeploymentMode={DeploymentConfiguration.DeploymentMode}, DiagnosticMode={OptimizationFlags.UseDiagnosticMode}\n");

                // Step 1: Use passed filter selections from UI (optional - can work without filters)
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] Selected filter items: {string.Join(", ", selectedFilterItems)}");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Selected filter items: {string.Join(", ", selectedFilterItems)}\n");

                // ⚠️ CRITICAL: Require filter selection - don't proceed without a filter ⚠️
                if (selectedFilterItems.Count == 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error("[CLASH_DEBUG] ERROR: No filters selected - cannot proceed with refresh");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ERROR: No filters selected - cannot proceed with refresh\n");

                    // Prompt user to select a filter
                    System.Windows.Forms.MessageBox.Show(
                        "Please select at least one filter before running Refresh.\n\nFilters are listed in the left panel.",
                        "No Filter Selected",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);

                    _statusLabel.Text = "Refresh cancelled - No filter selected";
                    _progressBar.Value = 0;
                    return Autodesk.Revit.UI.Result.Cancelled; // STOP - don't proceed without a filter
                }

                // ✅ UI STATE VALIDATION: Check all 4 required UI selections before proceeding
                var validationErrors = new List<string>();

                // 1. Check MEP Categories
                if (selectedMepCategories == null || selectedMepCategories.Count == 0)
                {
                    validationErrors.Add("MEP Categories");
                }

                // 2. Check Host Categories (Walls, Structural Framing, Floors, Ceilings)
                // Note: "Host Categories" and "Host Element Types" both read from the same UI panel,
                // so we only need to check one of them.
                var selectedHostCategories = FilterUiStateProvider.GetSelectedHostCategories?.Invoke() ?? new List<string>();
                if (selectedHostCategories == null || selectedHostCategories.Count == 0)
                {
                    validationErrors.Add("Host Categories");
                }

                // 3. Check MEP Linked Files (Reference Files)
                if (selectedReferenceFiles == null || selectedReferenceFiles.Count == 0)
                {
                    validationErrors.Add("MEP Linked Files (Reference Files)");
                }

                // 4. Check Host Linked Files
                if (selectedHostFiles == null || selectedHostFiles.Count == 0)
                {
                    validationErrors.Add("Host Linked Files");
                }

                // If any validation failed, show prompt and stop
                if (validationErrors.Count > 0)
                {
                    string errorMessage = "Please select the following before running Refresh:\n\n";
                    errorMessage += string.Join("\n", validationErrors.Select(e => $"• {e}"));
                    errorMessage += "\n\nThese selections are required for clash detection to work correctly.";

                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[CLASH_DEBUG] ERROR: Missing UI selections: {string.Join(", ", validationErrors)}");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ERROR: Missing UI selections: {string.Join(", ", validationErrors)}\n");

                    System.Windows.Forms.MessageBox.Show(
                        errorMessage,
                        "Missing Required Selections",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);

                    _statusLabel.Text = $"Refresh cancelled - Missing: {string.Join(", ", validationErrors)}";
                    _progressBar.Value = 0;
                    return Autodesk.Revit.UI.Result.Cancelled;
                }

                // ✅ All validations passed - log selections
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] ✅ UI State Validation PASSED:");
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG]   - MEP Categories: {selectedMepCategories.Count} selected");
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG]   - Host Categories: {selectedHostCategories.Count} selected");
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG]   - MEP Linked Files: {selectedReferenceFiles.Count} selected");
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG]   - Host Linked Files: {selectedHostFiles.Count} selected");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ✅ UI State Validation PASSED\n");

                // Step 2: Process filters and detect intersections
                var filtersToProcess = new List<Models.OpeningFilter>();

                // Try to load actual filters first
                foreach (var filterName in selectedFilterItems)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] Attempting to load filter: '{filterName}'");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Attempting to load filter: '{filterName}'\n");

                    var filter = _filterManagementService.LoadFilterAuto(filterName);

                    if (filter != null)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CLASH_DEBUG] ✓ Successfully loaded filter: '{filterName}'");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ✓ Successfully loaded filter: '{filterName}'\n");
                        filtersToProcess.Add(filter);
                    }
                    else
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[CLASH_DEBUG] ✗ Failed to load filter: '{filterName}' - file may not exist or deserialization failed");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ✗ Failed to load filter: '{filterName}' - file may not exist or deserialization failed\n");
                    }
                }

                // ✅ FIX: If no filters found, ERROR - don't auto-create fallback files
                if (filtersToProcess.Count == 0)
                {
                    string errorMessage = "No filter found. Please create a filter first before running Refresh.";
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[CLASH_DEBUG] {errorMessage}");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ERROR: {errorMessage}\n");

                    // ✅ FIXED: Use UIHelpers.ShowTaskDialogOnTop to ensure dialog appears in front of main UI
                    UIHelpers.ShowTaskDialogOnTop(
                        "Filter Required",
                        "No filter found!",
                        "Please create a filter in the Filter Management section before running Refresh.\n\nFilters define which MEP categories and reference files to process.");

                    return Autodesk.Revit.UI.Result.Failed; // Stop refresh - don't proceed without a filter
                }

                // ⚠️ CRITICAL: Ensure shared parameters are loaded into the project before any parameter operations
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info("[SHARED_PARAMS] Ensuring shared parameters are loaded into project before refresh operations");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [SHARED_PARAMS] Ensuring shared parameters are loaded into project\n");
                EnsureSharedParametersLoaded(_document);

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] MEP filters to process: {filtersToProcess.Count}");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] MEP filters to process: {filtersToProcess.Count}\n");

                foreach (var filter in filtersToProcess)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] Filter: {filter.Name} - Category: {filter.Category} - Enabled: {filter.IsEnabled}");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Filter: {filter.Name} - Category: {filter.Category} - Enabled: {filter.IsEnabled}\n");
                }

                // Step 3: Get current profile and initialize clash zone service
                var currentProfile = _appProfileService.GetCurrentProfile();
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] Current profile: {currentProfile?.Name ?? "NULL"}");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Profile loaded: {currentProfile?.Name ?? "NULL"}\n");

                if (currentProfile?.Configuration?.ClashZoneStorage == null)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info("[CLASH_DEBUG] No clash zone storage in profile - proceeding with direct clash detection");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] INFO: No clash zone storage - proceeding with direct clash detection\n");
                }

                // Step 4: Check for active document
                if (_document == null)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error("[CLASH_DEBUG] No active Revit document found");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ERROR: No active Revit document found\n");

                    _statusLabel.Text = "No active document";
                    _progressBar.Visible = false;
                    _refreshButton.Enabled = true;
                    return Autodesk.Revit.UI.Result.Failed;
                }

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] Document: {_document.Title}");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Document: {_document.Title}\n");

                // Step 5: Check for Duct Accessories specific validation
                if (selectedMepCategories != null && selectedMepCategories.Count == 1 && selectedMepCategories.Contains("Duct Accessories"))
                {
                    _statusLabel.Text = "Validating damper parameters...";
                    _progressBar.Visible = true;
                    _progressBar.Value = 10;

                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info("[DUCT_ACCESSORIES] Only Duct Accessories selected - validating damper parameters");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DUCT_ACCESSORIES] Only Duct Accessories selected - validating damper parameters\n");

                    // Validate damper parameters in selected linked mechanical file
                    if (!ValidateDamperParameters(selectedReferenceFiles, refreshLogName))
                    {
                        // Validation failed - user was prompted, exit refresh
                        _statusLabel.Text = "Damper parameter validation failed";
                        _progressBar.Visible = false;
                        _refreshButton.Enabled = true;
                        return Autodesk.Revit.UI.Result.Failed;
                    }
                }

                // Step 5.5: Replace mode shortcut (flag management only, no Filter XML access)
                // ✅ CRITICAL FIX: REPLACE-MODE should ONLY activate if:
                // 1. "Adopt to modified document" is OFF (enableThreePointValidation == false)
                // 2. AND all selected file combos are already processed in Global XML
                // If there are NEW file combos (e.g., AR+WALL), REPLACE-MODE should NOT activate - run normal detection instead
                bool hasNewFileCombos = false;
                if (selectedMepCategories != null && selectedMepCategories.Count > 0 &&
                    selectedReferenceFiles != null && selectedReferenceFiles.Count > 0 &&
                    selectedHostFiles != null && selectedHostFiles.Count > 0)
                {
                    // Build all file combos from UI selections
                    var allFileCombos = new List<(string LinkedFile, string HostFile)>();
                    foreach (var refFile in selectedReferenceFiles)
                    {
                        foreach (var hostFile in selectedHostFiles)
                        {
                            allFileCombos.Add((refFile, hostFile));
                        }
                    }

                    // Check if ANY combo is NEW (not processed) for ANY selected category
                    foreach (var category in selectedMepCategories)
                    {
                        if (string.IsNullOrWhiteSpace(category))
                            continue;

                        var processedKeys = GlobalIndexService.GetProcessedFileComboKeys(_document, category) ?? Enumerable.Empty<string>();
                        var comboKeysSet = new HashSet<string>(processedKeys, StringComparer.OrdinalIgnoreCase);

                        foreach (var combo in allFileCombos)
                        {
                            var comboNormalized = new ProcessedFileCombo { LinkedFile = combo.LinkedFile, HostFile = combo.HostFile };
                            var comboKey = comboNormalized.GetNormalizedKey();

                            if (!comboKeysSet.Contains(comboKey))
                            {
                                hasNewFileCombos = true;
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[REPLACE-MODE-CHECK] ✅ NEW FILE COMBO DETECTED: Category='{category}', Linked='{combo.LinkedFile}', Host='{combo.HostFile}', NormalizedKey='{comboKey}'");
                                }
                                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [REPLACE-MODE-CHECK] ✅ NEW FILE COMBO: Category='{category}', Linked='{combo.LinkedFile}', Host='{combo.HostFile}'\n");
                                break;
                            }
                        }

                        if (hasNewFileCombos)
                            break;
                    }
                }

                // ✅ REPLACE-MODE only activates if BOTH conditions are true:
                // 1. "Adopt to modified document" is OFF
                // 2. NO new file combos detected (all combos already processed)
                bool isReplaceMode = !enableThreePointValidation && !hasNewFileCombos;
                
                if (isReplaceMode)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[REPLACE-MODE] ✅ REPLACE-MODE ACTIVATED: enableThreePointValidation={enableThreePointValidation}, hasNewFileCombos={hasNewFileCombos}");
                    }
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [REPLACE-MODE] ✅ ACTIVATED: enableThreePointValidation={enableThreePointValidation}, hasNewFileCombos={hasNewFileCombos}\n");
                }
                else if (!enableThreePointValidation && hasNewFileCombos)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[REPLACE-MODE] ❌ REPLACE-MODE SKIPPED: New file combos detected - running normal detection instead");
                    }
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [REPLACE-MODE] ❌ SKIPPED: New file combos detected - running normal detection\n");
                }
                
                if (isReplaceMode)
                {
                    _statusLabel.Text = "Replace mode: syncing Global XML flags…";
                    _progressBar.Visible = true;
                    _progressBar.Value = 30;

                    SafeFileLogger.SafeAppendText(refreshLogName,
                        $"[{DateTime.Now}] [REPLACE-MODE] Activated replace mode – skipping Filter XML and detection. Running flag reset only.\n");

                    var categoriesForReset = selectedMepCategories?
                        .Where(cat => !string.IsNullOrWhiteSpace(cat))
                        .Select(cat => cat.Trim())
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList() ?? new List<string>();

                    int resetCount = 0;
                    if (categoriesForReset.Count == 0)
                    {
                        SafeFileLogger.SafeAppendText(refreshLogName,
                            $"[{DateTime.Now}] [REPLACE-MODE] WARNING: No MEP categories supplied – nothing to reset.\n");
                    }
                    else
                    {
                        try
                        {
                            resetCount = _flagManager.ResetFlagsForDeletedSleeves(categoriesForReset, null, refreshLogName);
                            SafeFileLogger.SafeAppendText(refreshLogName,
                                $"[{DateTime.Now}] [REPLACE-MODE] Flag reset completed. Categories={string.Join(", ", categoriesForReset)} → Resets={resetCount}\n");
                        }
                        catch (Exception resetEx)
                        {
                            SafeFileLogger.SafeAppendText(refreshLogName,
                                $"[{DateTime.Now}] [REPLACE-MODE] ERROR during flag reset: {resetEx}\n");
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Error($"[REPLACE-MODE] Flag reset failed: {resetEx.Message}");
                            _statusLabel.Text = $"Replace mode failed: {resetEx.Message}";
                            _progressBar.Visible = false;
                            _refreshButton.Enabled = true;
                            batchedLogger?.Dispose();
                            batchedLogger = null;
                            _memoryProfiler?.TakeSnapshot("REPLACE_MODE_ERROR", 0);
                            _memoryProfiler = null;
                            return Autodesk.Revit.UI.Result.Failed;
                        }
                    }

                    int totalEntries = 0;
                    int unresolvedEntries = 0;
                    var summaryBuilder = new StringBuilder();

                    foreach (var category in categoriesForReset)
                    {
                        try
                        {
                            var globalIndex = GlobalIndexService.LoadOrCreate(_document, category);
                            var entries = GlobalIndexService.GetAllEntries(globalIndex).ToList();
                            int categoryTotal = entries.Count;
                            int categoryUnresolved = entries.Count(e => !e.IsResolved && !e.IsClusterResolved);

                            totalEntries += categoryTotal;
                            unresolvedEntries += categoryUnresolved;

                            summaryBuilder.AppendLine(
                                $"[{DateTime.Now}] [REPLACE-MODE] Category '{category}': total={categoryTotal}, unresolved={categoryUnresolved}");
                        }
                        catch (Exception summaryEx)
                        {
                            summaryBuilder.AppendLine(
                                $"[{DateTime.Now}] [REPLACE-MODE] ERROR reading Global XML for '{category}': {summaryEx.Message}");
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[REPLACE-MODE] Failed to read Global XML for '{category}': {summaryEx.Message}");
                        }
                    }

                    if (summaryBuilder.Length > 0)
                        SafeFileLogger.SafeAppendText(refreshLogName, summaryBuilder.ToString());

                    _progressBar.Value = 100;
                    _statusLabel.Text = $"Replace mode complete. Reset {resetCount} flags. Unresolved {unresolvedEntries}/{totalEntries}.";
                    _refreshButton.Enabled = true;

                    SafeFileLogger.SafeAppendText(refreshLogName,
                        $"[{DateTime.Now}] [REPLACE-MODE] Completed replace mode refresh. TotalResets={resetCount}, Unresolved={unresolvedEntries}, Total={totalEntries}\n");

                    _memoryProfiler?.TakeSnapshot("REPLACE_MODE_END", 0);
                    _memoryProfiler = null;

                    batchedLogger?.Dispose();
                    batchedLogger = null;

                    return Autodesk.Revit.UI.Result.Succeeded;
                }

                // Step 6: Intersection detection will happen later (after existingClashZones is loaded and optimization service is created)

                // Step 6: Process clash zones
                _progressBar.Value = 40;
                _statusLabel.Text = "Processing clash zones...";

                // ✅ FIX: Load existing clash zones from Filter XML files (NOT profile)
                // This preserves IsResolved and IsClusterResolved flags from previous placement/clustering
                // Ensure per-project filters directory exists
                try { ProjectPathService.EnsureFiltersDirectory(_document); } catch { }
                // ✅ Get resolved GUID sets from Global XML (for filtering - no Revit API calls)
                // Note: Flag validation/reset happens in:
                // ✅ CRITICAL FIX: Load existing clash zones FIRST to get ALL categories (not just selectedMepCategories)
                // This ensures we check Global XML for ALL categories that have clash zones, even if:
                // - New categories were added to the filter
                // - New linked files were added (which might have clash zones from different categories)
                var existingClashZones = LoadExistingClashZonesFromFilterXml(selectedFilterItems, selectedMepCategories);

                // ✅ CRITICAL FIX: Build __globalsResolved from ALL categories that exist in clash zones
                // Uses OOP method GlobalIndexService.GetResolvedGuidsForCategories() instead of inline logic
                // Not just selectedMepCategories - this handles new categories and new linked files
                var __globalsResolved = new HashSet<Guid>();
                try
                {
                    // Step 1: Collect all categories to check
                    var categoriesToCheck = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    // Add selected categories from UI
                    if (selectedMepCategories != null)
                    {
                        foreach (var cat in selectedMepCategories)
                        {
                            if (!string.IsNullOrWhiteSpace(cat))
                                categoriesToCheck.Add(cat);
                        }
                    }

                    // Add categories from existing clash zones (handles new categories in filter)
                    if (existingClashZones?.ClashZones != null && existingClashZones.ClashZones.Count > 0)
                    {
                        var allCategoriesInExistingZones = existingClashZones.ClashZones
                            .Where(cz => !string.IsNullOrEmpty(cz.MepElementCategory))
                            .Select(cz => cz.MepElementCategory)
                            .Distinct(StringComparer.OrdinalIgnoreCase);

                        foreach (var cat in allCategoriesInExistingZones)
                        {
                            categoriesToCheck.Add(cat);
                        }

                        if (!DeploymentConfiguration.DeploymentMode && allCategoriesInExistingZones.Any())
                        {
                            var newCategories = allCategoriesInExistingZones
                                .Where(cat => selectedMepCategories == null || !selectedMepCategories.Contains(cat, StringComparer.OrdinalIgnoreCase))
                                .ToList();

                            if (newCategories.Count > 0)
                            {
                                DebugLogger.Info($"[REFRESH-GLOBAL-SYNC] Found {newCategories.Count} additional categories in existing clash zones: {string.Join(", ", newCategories)}");
                                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [REFRESH-GLOBAL-SYNC] Found {newCategories.Count} additional categories in existing clash zones\n");
                            }
                        }
                    }

                    // ✅ CRITICAL FIX: Also include ALL categories that have Global XML files
                    // This ensures categories from previous filter runs are checked even if they're not in the current Filter XML
                    // Example: User runs "Plumbing" filter with "Pipes" category → Places sleeves → Flags saved to Pipes_global.xml
                    // Then user runs "All" filter → "All" Filter XML might not have "Pipes" clash zones yet
                    // But Pipes_global.xml exists → Should be included in check to prevent duplicates
                    var allCategoriesWithGlobalXml = GlobalIndexService.GetAllCategoriesWithGlobalXml(_document);
                    foreach (var cat in allCategoriesWithGlobalXml)
                    {
                        categoriesToCheck.Add(cat);
                    }

                    if (!DeploymentConfiguration.DeploymentMode && allCategoriesWithGlobalXml.Count > 0)
                    {
                        var additionalCategories = allCategoriesWithGlobalXml
                            .Where(cat => !categoriesToCheck.Contains(cat, StringComparer.OrdinalIgnoreCase) ||
                                          (selectedMepCategories != null && selectedMepCategories.Contains(cat, StringComparer.OrdinalIgnoreCase)))
                            .ToList();

                        if (additionalCategories.Count > 0)
                        {
                            DebugLogger.Info($"[REFRESH-GLOBAL-SYNC] Found {allCategoriesWithGlobalXml.Count} categories with Global XML files: {string.Join(", ", allCategoriesWithGlobalXml)}");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [REFRESH-GLOBAL-SYNC] Found {allCategoriesWithGlobalXml.Count} categories with Global XML files\n");
                        }
                    }

                    // Step 2: Use OOP method to get resolved GUIDs for all categories
                    __globalsResolved = GlobalIndexService.GetResolvedGuidsForCategories(_document, categoriesToCheck);

                    if (!DeploymentConfiguration.DeploymentMode && __globalsResolved.Count > 0)
                    {
                        DebugLogger.Info($"[REFRESH-GLOBAL-SYNC] Loaded {__globalsResolved.Count} resolved GUIDs from Global XML for {categoriesToCheck.Count} categories");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [REFRESH-GLOBAL-SYNC] Loaded {__globalsResolved.Count} resolved GUIDs from {categoriesToCheck.Count} categories\n");
                    }
                }
                catch (Exception globalEx)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Error($"[REFRESH-GLOBAL-SYNC] Error building __globalsResolved: {globalEx.Message}");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [REFRESH-GLOBAL-SYNC] ERROR building __globalsResolved: {globalEx.Message}\n");
                    }
                }
                __logPhaseMem("AFTER_XML_LOAD", existingClashZones?.ClashZones, null, null);
                var existingCount = existingClashZones?.ClashZones?.Count ?? 0;

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] Loaded {existingCount} existing clash zones from Filter XML files");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Loaded {existingCount} existing clash zones from Filter XML files (selected filter + categories)\n");

                // ✅ OOP OPTIMIZATION: Create optimization service after existingClashZones is loaded
                // - If "Adopt to modified document" is CHECKED → Run full intersection detection (detect movements/modifications)
                // - If "Adopt to modified document" is UNCHECKED → Skip geometry checks for known valid pairs (user trusts model unchanged)
                var optimizationService = new IntersectionOptimizationService(existingClashZones);
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[PERFORMANCE] {optimizationService.GetOptimizationStatus()}");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [PERFORMANCE] {optimizationService.GetOptimizationStatus()}\n");

                // ✅ OOP REFACTORING: Use FlagManager for flag syncing from Global XML
                // Clash zones must remain in Filter XML even if resolved (for OK button logic and history)
                // Only update their flags to match Global XML state
                if (existingClashZones?.ClashZones != null && existingClashZones.ClashZones.Count > 0)
                {
                    try
                    {
                        // Group by category and sync flags for each category
                        var clashZonesByCategory = existingClashZones.ClashZones
                            .GroupBy(cz => cz.MepElementCategory)
                            .ToList();

                        foreach (var categoryGroup in clashZonesByCategory)
                        {
                            var category = categoryGroup.Key;
                            var categoryClashZones = categoryGroup.ToList();

                            _flagManager.SyncFlagsFromGlobal(categoryClashZones, category);

                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[REFRESH-GLOBAL-SYNC] Synced flags from Global XML for {categoryClashZones.Count} clash zones in category '{category}'");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [REFRESH-GLOBAL-SYNC] Synced {categoryClashZones.Count} clash zones from Global XML for category '{category}'\n");
                        }
                    }
                    catch (Exception syncEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error($"[REFRESH-GLOBAL-SYNC] Error syncing flags from Global XML: {syncEx.Message}");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [REFRESH-GLOBAL-SYNC] ERROR: {syncEx.Message}\n");
                    }

                    // ✅ CRITICAL FIX: Check for deleted sleeves IMMEDIATELY after syncing flags
                    // This must happen BEFORE deciding whether to skip intersection detection
                    // If sleeves are deleted, flags will be reset, making zones unresolved
                    // ✅ UNIFIED RESET: Single method handles everything - Global XML entries + optional Filter XML sync
                    int totalResetCount = 0;
                    try
                    {
                        // ✅ REUSE: Settings already loaded at method level (line ~464) - reuse here
                        // No need to reload settings - use enableThreePointValidation from method level

                        // Always check for deleted sleeves (regardless of 3-point validation setting)
                        // 3-point validation only affects geometry validation, not sleeve existence checks
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[FLAG-MANAGER] ===== UNIFIED RESET: CHECKING FOR DELETED SLEEVES (after flag sync) =====");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [FLAG-MANAGER] ===== UNIFIED RESET: CHECKING FOR DELETED SLEEVES =====\n");

                        // ✅ UNIFIED APPROACH: Build categories list and optional Filter XML clash zones dictionary
                        var categoriesToCheck = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                        // Add categories from Filter XML clash zones (if they exist)
                        if (existingClashZones?.ClashZones != null && existingClashZones.ClashZones.Count > 0)
                        {
                            var clashZonesByCategoryForReset = existingClashZones.ClashZones
                                .GroupBy(cz => cz.MepElementCategory)
                                .ToList();

                            foreach (var categoryGroup in clashZonesByCategoryForReset)
                            {
                                if (!string.IsNullOrWhiteSpace(categoryGroup.Key))
                                    categoriesToCheck.Add(categoryGroup.Key);
                            }
                        }

                        // Add selected categories from UI (may have Global XML entries even if not in Filter XML)
                        if (selectedMepCategories != null)
                        {
                            foreach (var cat in selectedMepCategories)
                            {
                                if (!string.IsNullOrWhiteSpace(cat))
                                    categoriesToCheck.Add(cat);
                            }
                        }

                        // ✅ OPTIONAL: Build Filter XML clash zones dictionary for sync-back (if Filter XML exists)
                        Dictionary<string, List<ClashZone>> clashZonesByCategory = null;
                        if (existingClashZones?.ClashZones != null && existingClashZones.ClashZones.Count > 0)
                        {
                            clashZonesByCategory = existingClashZones.ClashZones
                                .GroupBy(cz => cz.MepElementCategory)
                                .ToDictionary(g => g.Key, g => g.ToList());
                        }

                        // ✅ UNIFIED RESET: Single method call handles everything
                        if (categoriesToCheck.Count > 0)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[FLAG-MANAGER] Calling unified reset for {categoriesToCheck.Count} categories: {string.Join(", ", categoriesToCheck)}");

                            totalResetCount = _flagManager.ResetFlagsForDeletedSleeves(categoriesToCheck.ToList(), clashZonesByCategory, refreshLogName);

                            if (totalResetCount > 0)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[FLAG-MANAGER] ✅ Total {totalResetCount} clash zones reset due to deleted sleeves - these zones are now unresolved");
                                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [FLAG-MANAGER] ✅ Total {totalResetCount} clash zones reset due to deleted sleeves\n");
                            }
                            else
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[FLAG-MANAGER] All sleeves exist - no flags reset");
                                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [FLAG-MANAGER] All sleeves exist - no flags reset\n");
                            }
                            
                            // ✅ CRITICAL FIX: Re-sync flags from Global XML after reset completes
                            // This ensures zones loaded from SQLite get the updated (reset) flags from Global XML
                            // Global XML is the source of truth, and reset just updated it
                            // Always re-sync, even if no resets occurred, to ensure zones have latest flags
                            if (clashZonesByCategory != null && categoriesToCheck.Count > 0)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[FLAG-MANAGER] Re-syncing flags from Global XML after reset for {categoriesToCheck.Count} categories");
                                
                                foreach (var category in categoriesToCheck)
                                {
                                    if (string.IsNullOrWhiteSpace(category))
                                        continue;
                                    
                                    if (clashZonesByCategory.TryGetValue(category, out var categoryZones) && categoryZones != null && categoryZones.Count > 0)
                                    {
                                        _flagManager.SyncFlagsFromGlobal(categoryZones, category);
                                        var categoryUnresolvedCount = categoryZones.Count(z => !z.IsResolved && !z.IsClusterResolved);
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[FLAG-MANAGER] ✅ Re-synced flags for {categoryZones.Count} zones in category '{category}' from Global XML ({categoryUnresolvedCount} unresolved)");
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[FLAG-MANAGER] ⚠️ No categories to check - categoriesToCheck is empty. Filter XML clash zones: {existingClashZones?.ClashZones?.Count ?? 0}, Selected categories: {selectedMepCategories?.Count ?? 0}");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [FLAG-MANAGER] ⚠️ WARNING: No categories to check - reset skipped\n");
                        }
                    }
                    catch (Exception resetEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error($"[FLAG-MANAGER] Error checking for deleted sleeves: {resetEx.Message}");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [FLAG-MANAGER] ERROR: {resetEx.Message}\n");
                    }
                }
                // ✅ CRITICAL FIX: Section box filtering should only filter for PROCESSING, not permanently remove from storage
                // Keep original list for saving to Filter XML (append-only principle)
                // Zones outside section box should remain in Filter XML for future use
                var sectionBoxFilteredZones = existingClashZones?.ClashZones ?? new List<Models.ClashZone>();
                if (_document.ActiveView is View3D secView && secView != null)
                {
                    try
                    {
                        var secBox = secView.GetSectionBox();
                        var inv = secBox.Transform.Inverse;
                        var kept = new List<Models.ClashZone>();
                        foreach (var cz in existingClashZones?.ClashZones ?? new List<Models.ClashZone>())
                        {
                            var pt = new XYZ(cz.IntersectionPointX, cz.IntersectionPointY, cz.IntersectionPointZ);
                            var local = inv.OfPoint(pt);
                            bool inside = local.X >= secBox.Min.X && local.X <= secBox.Max.X &&
                                          local.Y >= secBox.Min.Y && local.Y <= secBox.Max.Y &&
                                          local.Z >= secBox.Min.Z && local.Z <= secBox.Max.Z;
                            if (inside) kept.Add(cz);
                        }
                        sectionBoxFilteredZones = kept; // Use filtered list for processing only
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CLASH_DEBUG] Section-box filtered existing zones for processing: {kept.Count} (original: {existingClashZones?.ClashZones?.Count ?? 0})");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Section-box filtered zones: {kept.Count} for processing (keeping all {existingClashZones?.ClashZones?.Count ?? 0} in storage)\n");
                    }
                    catch (Exception ex)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[CLASH_DEBUG] Section-box prefilter failed: {ex.Message}");
                    }
                }

                // ✅ CRITICAL: 3-Point Validation - Validate clash zone validity before processing
                // Validation Criteria: 1) MEP Element exists, 2) Structural Element exists, 3) Elements still intersect
                if (existingClashZones?.ClashZones != null && existingClashZones.ClashZones.Count > 0)
                {
                    // ✅ REUSE: Settings already loaded at method level (line ~464) - reuse enableThreePointValidation
                    // enableThreePointValidation is already declared at method level

                    if (enableThreePointValidation)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CLASH_DEBUG] ===== 3-POINT VALIDATION ENABLED FOR {existingClashZones.ClashZones.Count} EXISTING CLASH ZONES =====");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ===== 3-POINT VALIDATION ENABLED FOR {existingClashZones.ClashZones.Count} EXISTING CLASH ZONES =====\n");
                    }
                    else
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CLASH_DEBUG] ===== 3-POINT VALIDATION DISABLED - SKIPPING VALIDATION, ONLY FLAG CHECKS =====");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ===== 3-POINT VALIDATION DISABLED - SKIPPING VALIDATION, ONLY FLAG CHECKS =====\n");
                    }

                    var validClashZones = new List<Models.ClashZone>();
                    var invalidClashZones = new List<Models.ClashZone>();
                    int removedCount = 0;

                    // ✅ OOP REFACTORING: Use ThreePointValidator for validation (replaces inline validation logic)
                    // ✅ CONFIGURATION: Only run validation if enabled in settings
                    foreach (var existingZone in existingClashZones.ClashZones)
                    {
                        try
                        {
                            if (enableThreePointValidation)
                            {
                                // ✅ PERFORM 3-POINT VALIDATION (costly but thorough)
                                var validationResult = _threePointValidator.Validate(existingZone, _document);

                                if (!validationResult.IsValid)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[3-POINT-VALIDATION] ❌ INVALID: Zone {existingZone.Id} - {validationResult.FailureReason}");
                                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [3-POINT-VALIDATION] ❌ INVALID: Zone {existingZone.Id} - {validationResult.FailureReason}\n");
                                    invalidClashZones.Add(existingZone);
                                    continue;
                                }

                                // ✅ VALID - Keep clash zone
                                validClashZones.Add(existingZone);

                                // Update intersection point if it changed
                                if (validationResult.UpdatedIntersectionPoint != null && validationResult.IntersectionPointMovement.HasValue)
                                {
                                    // ✅ CRITICAL: Update intersection point FIRST before deletion
                                    // This ensures FlagManager uses the correct NEW intersection point when updating Global XML
                                    // Update intersection point (linked coordinates for individual sleeve placement)
                                    existingZone.IntersectionPointX = validationResult.UpdatedIntersectionPoint.X;
                                    existingZone.IntersectionPointY = validationResult.UpdatedIntersectionPoint.Y;
                                    existingZone.IntersectionPointZ = validationResult.UpdatedIntersectionPoint.Z;

                                    // ✅ CRITICAL: Update active document coordinates for proximity calculation
                                    if (validationResult.UpdatedActiveDocumentPoint != null)
                                    {
                                        existingZone.SleevePlacementPointActiveDocumentX = validationResult.UpdatedActiveDocumentPoint.X;
                                        existingZone.SleevePlacementPointActiveDocumentY = validationResult.UpdatedActiveDocumentPoint.Y;
                                        existingZone.SleevePlacementPointActiveDocumentZ = validationResult.UpdatedActiveDocumentPoint.Z;
                                    }

                                    // ✅ OOP HELPER: Delete old sleeve if intersection point moved beyond tolerance
                                    // Only delete if "Adopt to modified document" is enabled (enableThreePointValidation == true)
                                    // This allows new sleeve to be placed at the new intersection point
                                    // NOTE: Intersection point already updated above, so FlagManager will use correct new coordinates
                                    if (enableThreePointValidation && _flagManager != null && validationResult.IntersectionPointMovement.Value > 0.1)
                                    {
                                        try
                                        {
                                            _flagManager.DeleteSleeveForIntersectionPointChange(
                                                existingZone,
                                                existingZone.MepElementCategory,
                                                validationResult.IntersectionPointMovement.Value);
                                        }
                                        catch (Exception deleteEx)
                                        {
                                            if (!DeploymentConfiguration.DeploymentMode)
                                                DebugLogger.Warning($"[3-POINT-VALIDATION] Error deleting sleeve for intersection point change: {deleteEx.Message}");
                                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [3-POINT-VALIDATION] ERROR deleting sleeve: {deleteEx.Message}\n");
                                        }
                                    }

                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[3-POINT-VALIDATION] ✅ VALID + UPDATED: Zone {existingZone.Id} - Intersection point updated (Δ={validationResult.IntersectionPointMovement.Value:F3}ft)");
                                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [3-POINT-VALIDATION] ✅ VALID + UPDATED: Zone {existingZone.Id} - Intersection updated (Δ={validationResult.IntersectionPointMovement.Value:F3}ft)\n");
                                }
                                else
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[3-POINT-VALIDATION] ✅ VALID: Zone {existingZone.Id} - All 3 points valid, no update needed");
                                }
                            }
                            else
                            {
                                // ✅ SKIP 3-POINT VALIDATION - Only flag checks (faster)
                                // All zones are considered valid if 3-point validation is disabled
                                validClashZones.Add(existingZone);
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[FLAG-ONLY] ✅ Zone {existingZone.Id} - 3-point validation skipped, relying on flag checks only");
                            }
                        }
                        catch (Exception ex)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[3-POINT-VALIDATION] ❌ ERROR: Zone {existingZone.Id} - Exception during validation: {ex.Message}");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [3-POINT-VALIDATION] ❌ ERROR: Zone {existingZone.Id} - {ex.Message}\n");
                            invalidClashZones.Add(existingZone); // Treat as invalid on error
                        }
                    }

                    // ✅ DELETED SLEEVES CHECK ALREADY COMPLETED above (right after SyncFlagsFromGlobal)
                    // No need to check again here - it was done before validation to ensure flags are correct

                    // ✅ CRITICAL FIX: Only remove invalid clash zones if "Adopt Document" is enabled AND validation failed
                    // If disabled, keep ALL existing zones (append-only principle)
                    if (enableThreePointValidation && invalidClashZones.Count > 0)
                    {
                        removedCount = invalidClashZones.Count;

                        // ✅ OOP REFACTORING: Use GuidManager to remove invalid clash zones from Global XML
                        foreach (var invalidClashZone in invalidClashZones)
                        {
                            try
                            {
                                _guidManager.RemoveFromGlobalXml(invalidClashZone.Id, invalidClashZone.MepElementCategory);
                            }
                            catch (Exception globalEx)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Warning($"[3-POINT-VALIDATION] Error removing ClashZone {invalidClashZone.Id} from Global XML: {globalEx.Message}");
                            }
                        }

                        if (invalidClashZones.Count > 0)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[3-POINT-VALIDATION] Removed {invalidClashZones.Count} invalid clash zones from Global XML");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [3-POINT-VALIDATION] Removed {invalidClashZones.Count} invalid clash zones from Global XML\n");
                        }

                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[3-POINT-VALIDATION] Removed {removedCount} invalid clash zones (MEP/Structural deleted or no longer intersect)");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [3-POINT-VALIDATION] Removed {removedCount} invalid clash zones\n");
                    }
                    else if (!enableThreePointValidation && invalidClashZones.Count > 0)
                    {
                        // ✅ CRITICAL FIX: If "Adopt Document" is disabled, keep ALL zones (including invalid ones)
                        // This preserves placement data even if elements are temporarily missing
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[3-POINT-VALIDATION] Skipping removal of {invalidClashZones.Count} invalid zones (Adopt Document disabled - append-only mode)");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [3-POINT-VALIDATION] Keeping {invalidClashZones.Count} invalid zones (Adopt Document disabled)\n");
                        // Add invalid zones back to valid list (they'll be kept in Filter XML)
                        validClashZones.AddRange(invalidClashZones);
                    }

                    // ✅ CRITICAL FIX: Only replace with validated zones if "Adopt Document" is enabled
                    // If disabled, keep ALL existing zones (append-only)
                    if (enableThreePointValidation)
                    {
                        // Replace with validated clash zones only (invalid zones removed)
                        existingClashZones.ClashZones = validClashZones;
                    }
                    // else: Keep existingClashZones.ClashZones unchanged (all zones preserved)

                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[3-POINT-VALIDATION] Validation complete: {validClashZones.Count} valid, {removedCount} removed");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [3-POINT-VALIDATION] Result: {validClashZones.Count} valid, {removedCount} removed\n");
                }

                // ✅ ARCHITECTURE: Intersection detection is ALWAYS required
                // Even if all existing zones are resolved, new clashes may exist from:
                // - New MEP elements added to the model
                // - New structural elements added to the model
                // - Modified elements that now intersect
                // Performance optimization happens AFTER detection by filtering known intersections

                // ✅ NOTE: Intersection points already updated during 3-point validation above
                // Continue with orientation calculation for validated clash zones
                if (existingClashZones?.ClashZones != null && existingClashZones.ClashZones.Count > 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] ===== CALCULATING ORIENTATION FOR {existingClashZones.ClashZones.Count} VALIDATED CLASH ZONES =====");

                    int orientationUpdatedCount = 0;

                    // Calculate orientation for validated clash zones
                    foreach (var existingZone in existingClashZones.ClashZones)
                    {
                        try
                        {
                            // Get the MEP and structural elements (already validated above)
                            var mepElement = ElementRetrievalService.GetElementFromDocumentOrLinked(_document, existingZone.MepElementId, enableLogging: false);
                            var structuralElement = ElementRetrievalService.GetElementFromDocumentOrLinked(_document, existingZone.StructuralElementId, enableLogging: false);

                            if (mepElement == null || structuralElement == null)
                                continue; // Skip if elements not found (shouldn't happen after validation)

                            // ✅ ORIENTATION: Cache minimal data (no heavy API retention) and compute orientation post-API
                            try
                            {
                                // ✅ OOP REFACTORING: Use centralized WallDirectionService (eliminates code duplication)
                                // This ensures consistent orientation calculation across all services
                                string hostOrientation = WallDirectionService.GetHostOrientation(structuralElement);

                                // Cache structural normal for other uses
                                var hostNormal = WallDirectionService.GetWallNormal(structuralElement);
                                if (hostNormal != null)
                                {
                                    existingZone.StructuralElementNormalX = hostNormal.X;
                                    existingZone.StructuralElementNormalY = hostNormal.Y;
                                    existingZone.StructuralElementNormalZ = hostNormal.Z;
                                }

                                // MEP orientation vector (projection to XY)
                                XYZ mepDir = null;
                                if (mepElement.Location is LocationCurve mepLc)
                                {
                                    mepDir = (mepLc.Curve.GetEndPoint(1) - mepLc.Curve.GetEndPoint(0)).Normalize();
                                }

                                if (mepDir != null)
                                {
                                    existingZone.MepElementOrientationX = mepDir.X;
                                    existingZone.MepElementOrientationY = mepDir.Y;
                                    existingZone.MepElementOrientationZ = mepDir.Z;
                                }

                                // ✅ Set HostOrientation (now calculated from direction, matching CreateClashZone)
                                existingZone.HostOrientation = hostOrientation;
                            }
                            catch { /* non-fatal */ }
                        }
                        catch (Exception ex)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[CLASH_DEBUG] ❌ Zone {existingZone.Id}: Error calculating orientation - {ex.Message}");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ❌ Zone {existingZone.Id}: Orientation error - {ex.Message}\n");
                        }
                    }

                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] Orientation calculation complete for {existingClashZones.ClashZones.Count} validated clash zones");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Orientation calculation complete\n");

                    // ✅ MEMORY OPTIMIZATION: Clear Revit API objects after recalculation to free memory.
                    //    IMPORTANT: this loop is intentionally placed *after* every persistence call.
                    //    Clearing before we serialize would wipe placement points / bounding boxes from
                    //    the ClashZone instances and the XML save would fall back to host centroids.
                    foreach (var existingZone in existingClashZones.ClashZones)
                    {
                        existingZone.ClearRevitApiObjects();
                    }
                }

                // Step 6: Perform intersection detection (AFTER existingClashZones loaded and optimization service created)
                // ✅ TWO PATHS: Based on "Adopt to modified document" setting
                // ✅ GLOBAL XML DECISION: Use Global XML's ProcessedFileCombos to decide if intersection detection is needed
                // Filter XML is ONLY used for storing clash zone data (placement points, dimensions) for sleeve placement
                // Decision logic:
                // - If all file combos are already processed in Global XML → Skip detection, use Filter XML for placement data
                // - If there are unprocessed file combos → Run detection
                // - "Adopt Document" setting still controls 3-point validation for existing zones

                List<(Element, Element, BoundingBoxXYZ, XYZ)> currentIntersections;

                // ✅ DECLARE VARIABLES OUTSIDE IF/ELSE BLOCK: These are used in shared code after the if/else
                var filteredIntersections = new List<(Element, Element, BoundingBoxXYZ, XYZ)>();
                int skippedResolvedCount = 0; // Track how many intersections were filtered out because resolved
                var knownIntersectionsMap = new Dictionary<(int mepId, int hostId, string pointKey), ClashZone>();
                var resolvedIntersectionPointsFromGlobalXml = new Dictionary<(int mepId, int hostId, string pointKey), CategoryGlobalIndexEntry>();
                List<string> filesToProcessForRef = null;
                List<string> filesToProcessForHost = null;
                // ✅ NEW FILE DETECTION VARIABLES: Also declared outside if/else block for use in shared code
                bool hasNewFilesAdded = false;
                List<string> newReferenceFiles = null; // Only NEW reference files (not already processed)
                List<string> newHostFiles = null; // Only NEW host files (not already processed)
                HashSet<string> savedRefFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                HashSet<string> savedHostFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                // ✅ CRITICAL: Declare view3D outside if/else block so it's accessible in shared code (section box filtering)
                View3D view3D = null;

                // ✅ GLOBAL XML DECISION: Check if all file combos are already processed in Global XML
                // This replaces the Filter XML existence check - Global XML is the single source of truth for detection decisions
                bool allFileCombosProcessed = false;
                bool filterHasPlacementData = false;
                bool hasNewSelectionsInUi = false;
                if (selectedMepCategories != null && selectedMepCategories.Count > 0 &&
                    selectedReferenceFiles != null && selectedReferenceFiles.Count > 0 &&
                    selectedHostFiles != null && selectedHostFiles.Count > 0)
                {
                    // Build all file combos from UI selections
                    var allFileCombos = new List<(string LinkedFile, string HostFile)>();
                    foreach (var refFile in selectedReferenceFiles)
                    {
                        foreach (var hostFile in selectedHostFiles)
                        {
                            allFileCombos.Add((refFile, hostFile));
                        }
                    }

                    // Check if ALL combos are processed for ALL selected categories
                    bool allProcessed = true;
                    foreach (var category in selectedMepCategories)
                    {
                        if (string.IsNullOrWhiteSpace(category))
                            continue;

                        var processedKeys = GlobalIndexService.GetProcessedFileComboKeys(_document, category) ?? Enumerable.Empty<string>();
                        var comboKeysSet = new HashSet<string>(processedKeys);

                        // Check if all combos are processed for this category
                        Func<string, string> norm = s =>
                        {
                            if (string.IsNullOrWhiteSpace(s)) return string.Empty;
                            var trimmed = s;
                            var idxParen = trimmed.IndexOf('(');
                            if (idxParen >= 0) trimmed = trimmed.Substring(0, idxParen);
                            trimmed = System.IO.Path.GetFileNameWithoutExtension(trimmed);
                            trimmed = trimmed.ToLowerInvariant().Replace("_detached", "");
                            trimmed = trimmed.Replace('_', ' ').Replace('-', ' ');
                            trimmed = System.Text.RegularExpressions.Regex.Replace(trimmed, "\\s+", " ");
                            return trimmed.Trim();
                        };

                        foreach (var combo in allFileCombos)
                        {
                            var comboNormalized = new ProcessedFileCombo { LinkedFile = combo.LinkedFile, HostFile = combo.HostFile };
                            var comboKey = comboNormalized.GetNormalizedKey();

                            if (!comboKeysSet.Contains(comboKey))
                            {
                                allProcessed = false;
                                break;
                            }
                        }

                        if (!allProcessed)
                            break;
                    }

                    allFileCombosProcessed = allProcessed && allFileCombos.Count > 0;

                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[GLOBAL-XML-DECISION] All file combos processed: {allFileCombosProcessed} (Total combos: {allFileCombos.Count}, Categories: {selectedMepCategories.Count})");
                    }
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-XML-DECISION] All file combos processed: {allFileCombosProcessed}\n");

                    // ✅ FILTER XML CHECK: Do we already have placement data in the filter for these categories?
                    if (existingClashZones?.ClashZones != null && existingClashZones.ClashZones.Count > 0)
                    {
                        var allowedCategories = new HashSet<string>(selectedMepCategories, StringComparer.OrdinalIgnoreCase);
                        filterHasPlacementData = existingClashZones.ClashZones.Any(cz =>
                            cz != null &&
                            !string.IsNullOrWhiteSpace(cz.MepElementCategory) &&
                            allowedCategories.Contains(cz.MepElementCategory));

                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[GLOBAL-XML-DECISION] Filter XML has placement data for selected categories: {filterHasPlacementData}");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-XML-DECISION] Filter XML has placement data: {filterHasPlacementData}\n");
                    }

                    var enabledFilterForComparison = filtersToProcess.FirstOrDefault(f => f.IsEnabled);
                    hasNewSelectionsInUi = DetectNewUiSelections(
                        _document,
                        enabledFilterForComparison,
                        selectedReferenceFiles,
                        selectedHostFiles,
                        selectedMepCategories,
                        selectedHostCategories);

                    if (hasNewSelectionsInUi && !DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info("[GLOBAL-XML-DECISION] UI selections include new linked files/categories → detection required");
                }

                // ✅ CRITICAL: Check if unresolved zones exist in Global XML
                // If unresolved zones exist, skip detection and use Filter XML for placement (even if combo not processed)
                bool hasUnresolvedZonesInGlobalXml = false;
                if (selectedMepCategories != null && selectedMepCategories.Count > 0)
                {
                    foreach (var category in selectedMepCategories)
                    {
                        if (string.IsNullOrWhiteSpace(category))
                            continue;

                        try
                        {
                            var globalIndex = GlobalIndexService.LoadOrCreate(_document, category);

                            // ✅ CRITICAL FIX: Use GetAllEntries to get entries from BOTH hierarchical and flat structures
                            var allEntries = GlobalIndexService.GetAllEntries(globalIndex).ToList();

                            if (allEntries != null && allEntries.Count > 0)
                            {
                                // Check if there are any unresolved entries (flags are false)
                                int categoryUnresolvedCount = allEntries.Count(e => !e.IsResolved && !e.IsClusterResolved);
                                if (categoryUnresolvedCount > 0)
                                {
                                    hasUnresolvedZonesInGlobalXml = true;
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        DebugLogger.Info($"[GLOBAL-XML-DECISION] Found {categoryUnresolvedCount} unresolved zones in Global XML for category '{category}' - will skip detection, use Filter XML for placement");
                                    }
                                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-XML-DECISION] Category '{category}': {categoryUnresolvedCount} unresolved zones - skip detection\n");
                                    break; // Found unresolved zones, no need to check other categories
                                }
                            }
                        }
                        catch (Exception globalEx)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[GLOBAL-XML-DECISION] Error checking Global XML for unresolved zones in category '{category}': {globalEx.Message}");
                        }
                    }
                }

                // ✅ PATH DECISION LOGIC:
                // 1. If unresolved zones exist AND combo IS processed → Skip detection, use Filter XML (Filter XML has data from previous detection)
                // 2. If all combos processed AND no unresolved zones → Skip detection, use Filter XML (all resolved)
                // 3. If unresolved zones exist BUT combo NOT processed → Run detection (Filter XML has no data yet)
                // 4. If combo NOT processed AND no unresolved zones → Run detection (new combo, need to detect)
                // Filter XML is ONLY used for placement data, not for detection decisions
                // CRITICAL: Can only skip detection if Filter XML has data (combo was processed before)
                bool canUseFilterXmlFromUnresolved = hasUnresolvedZonesInGlobalXml && filterHasPlacementData;
                // ✅ NEW: When "Adopt to modified document" is OFF (enableThreePointValidation == false) and
                // we already have placement data in the filter, skip intersection detection even if the combo
                // was never recorded as processed in Global XML. This honours the replay-only expectation.
                bool canUseFilterXmlFromSnapshot = !enableThreePointValidation && filterHasPlacementData;
                bool detectionRan = false;

                if (!enableThreePointValidation && !hasNewSelectionsInUi && (allFileCombosProcessed || canUseFilterXmlFromUnresolved || canUseFilterXmlFromSnapshot))
                {
                    // ✅ PATH 1: Unresolved zones exist AND combo processed → Filter XML has data, skip detection
                    // OR: All combos processed AND all resolved → Skip detection
                    // OR: Replay-only mode (Adopt OFF) with saved placement data available
                    string path1Reason;
                    if (allFileCombosProcessed)
                    {
                        path1Reason = hasUnresolvedZonesInGlobalXml
                            ? "Unresolved zones exist in Global XML (combo processed, Filter XML has data)"
                            : "All file combos processed and all resolved";
                    }
                    else if (canUseFilterXmlFromUnresolved)
                    {
                        path1Reason = "Unresolved zones exist and Filter XML already has placement data";
                    }
                    else
                    {
                        path1Reason = "Replay-only refresh (Adopt OFF) with existing placement snapshot";
                    }

                    _statusLabel.Text = $"Using Filter XML for placement data ({path1Reason})...";
                    _progressBar.Value = 25;

                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] ═══ PATH 1: {path1Reason} - Skipping intersection detection, using Filter XML for placement data ═══");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ═══ PATH 1: {path1Reason} - Skipping detection, using Filter XML for placement data ═══\n");

                    // ✅ CRITICAL: Reconstruct SleevePlacementPoint and IntersectionPoint from XML-serializable properties
                    // Filter XML contains placement data needed for sleeve placement (from previous detection)
                    if (existingClashZones?.ClashZones != null)
                    {
                        foreach (var cz in existingClashZones.ClashZones)
                        {
                            // Ensure SleevePlacementPoint is reconstructed from XML properties
                            cz.EnsureSleevePlacementPointReconstructed();

                            // Reconstruct IntersectionPoint if needed
                            if (cz.IntersectionPoint == null && (Math.Abs(cz.IntersectionPointX) > 1e-9 || Math.Abs(cz.IntersectionPointY) > 1e-9 || Math.Abs(cz.IntersectionPointZ) > 1e-9))
                            {
                                cz.IntersectionPoint = new XYZ(cz.IntersectionPointX, cz.IntersectionPointY, cz.IntersectionPointZ);
                            }
                        }
                    }

                    // Set empty intersections list (will use existing clash zones directly)
                    currentIntersections = new List<(Element, Element, BoundingBoxXYZ, XYZ)>();

                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] PATH 1: Reconstructed {existingClashZones?.ClashZones?.Count ?? 0} clash zones from Filter XML - ready for placement using saved placement data");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] PATH 1: Reconstructed {existingClashZones?.ClashZones?.Count ?? 0} clash zones - will use Filter XML placement data\n");
                }
                else
                {
                    // ✅ PATH 2: Run intersection detection when:
                    // 1. Combo NOT processed (Filter XML has no data yet) - even if unresolved zones exist
                    // 2. Combo NOT processed AND no unresolved zones (new combo)
                    // 3. "Adopt Document" enabled (always run detection for 3-point validation)
                    //
                    // ✅ COMPLETE FLOW FOR PATH 2 (combo NOT processed):
                    // Step 1: Run intersection detection → Get new intersections
                    // Step 2: Create clash zones → Save to Filter XML (placement data: coordinates, dimensions, host type, family name)
                    // Step 3: Ensure Global XML entries for NEW zones (flags=false) → Save to Global XML
                    // Step 4: Mark file combo as processed in Global XML only (ProcessedFileCombos list)
                    // Step 5: Check unresolved zones in Global XML → Log count for user
                    // Step 6: USER ACTION: Place Sleeves button → 
                    //         - Load Global XML (get unresolved zones with flags=false)
                    //         - Load Filter XML (get placement data: coordinates, dimensions, host type)
                    //         - Place sleeves using Filter XML data
                    //         - Update flags to true in Global XML after successful placement
                    string path2Reason = "";
                    if (!enableThreePointValidation)
                    {
                        if (hasNewSelectionsInUi)
                            path2Reason = "UI selections include new linked files or categories";
                        else if (!allFileCombosProcessed && hasUnresolvedZonesInGlobalXml)
                            path2Reason = "Combo NOT processed yet (Filter XML has no data, need to detect first)";
                        else if (!allFileCombosProcessed && !hasUnresolvedZonesInGlobalXml)
                            path2Reason = "Combo NOT processed (new combo, need detection)";
                    }
                    else
                    {
                        path2Reason = "Adopt Document enabled (3-point validation)";
                    }

                    detectionRan = true;

                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] ═══ PATH 2: {path2Reason} - Running intersection detection ═══");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ═══ PATH 2: {path2Reason} - Running detection ═══\n");

                    _statusLabel.Text = "Detecting intersections...";
                    _progressBar.Visible = true;
                    _progressBar.Value = 20;

                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] ═══ PATH 2: Running intersection detection (Optimization: {optimizationService.ShouldSkipKnownPairsGeometryCheck}, Known pairs: {optimizationService.KnownValidPairs.Count}) ═══");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ═══ PATH 2: Running intersection detection (Optimization: {optimizationService.ShouldSkipKnownPairsGeometryCheck}) ═══\n");

                    // Build allowed host types from current UI (UI PRECEDENCE)
                    var allowedHostTypesUI = new HashSet<string>(
                        FilterUiStateProvider.GetSelectedHostCategories?.Invoke() ?? new List<string>(),
                        StringComparer.OrdinalIgnoreCase);
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] UI Selected MEP categories: {string.Join(", ", selectedMepCategories ?? new List<string>())}");
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] UI Selected Host types: {string.Join(", ", allowedHostTypesUI)}");
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] UI Selected Host files: {string.Join(", ", selectedHostFiles ?? new List<string>())}");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] UI Selected MEP categories: {string.Join(", ", selectedMepCategories ?? new List<string>())}\n");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] UI Selected Host types: {string.Join(", ", allowedHostTypesUI)}\n");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] UI Selected Host files: {string.Join(", ", selectedHostFiles ?? new List<string>())}\n");

                    // Route intersection logs into both debug and refresh log files
                    var intersectionService = new IntersectionDetectionService(msg =>
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info(msg);
                        try { SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] {msg}"); } catch { }
                    });

                    // ⚠️ CRITICAL: Check memory/timeout limits before heavy operation
                    try
                    {
                        _memoryManager.CheckLimits();
                    }
                    catch (TimeoutException ex)
                    {
                        SafeFileLogger.SafeAppendText("refresh_timeouts.log", $"Timeout before intersection detection: {ex.Message}");
                        _statusLabel.Text = $"Operation timeout: {ex.Message}";
                        _progressBar.Visible = false;
                        _refreshButton.Enabled = true;
                        System.Windows.Forms.MessageBox.Show(
                            ex.Message,
                            "Operation Timeout",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Warning);
                        return Autodesk.Revit.UI.Result.Cancelled;
                    }
                    catch (OutOfMemoryException ex)
                    {
                        SafeFileLogger.SafeAppendText("refresh_memory.log", $"Memory limit before intersection detection: {ex.Message}");
                        _statusLabel.Text = $"Memory limit exceeded: Please try with a smaller file";
                        _progressBar.Visible = false;
                        _refreshButton.Enabled = true;
                        System.Windows.Forms.MessageBox.Show(
                            ex.Message,
                            "Memory Limit Exceeded",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Warning);
                        return Autodesk.Revit.UI.Result.Failed;
                    }

                    // ✅ CRITICAL FIX: Get 3D view or find/create one if ActiveView is not 3D
                    // ✅ REUSE: view3D is already declared outside if/else block - initialize here
                    view3D = _document.ActiveView as View3D;
                    if (view3D == null)
                    {
                        // Try to find any existing 3D view
                        var all3DViews = new FilteredElementCollector(_document)
                            .OfClass(typeof(View3D))
                            .Cast<View3D>()
                            .Where(v => !v.IsTemplate && v.CanBePrinted)
                            .FirstOrDefault();

                        if (all3DViews != null)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[CLASH_DEBUG] Active view is not 3D. Using existing 3D view: {all3DViews.Name}");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Active view is not 3D. Using existing 3D view: {all3DViews.Name}\n");
                            view3D = all3DViews;
                        }
                        else
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Error("[CLASH_DEBUG] ERROR: No 3D view found in document. Please create a 3D view and try again.");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] ERROR: No 3D view available for intersection detection\n");
                            _statusLabel.Text = "Error: No 3D view found. Please create a 3D view.";
                            _progressBar.Value = 0;
                            return Autodesk.Revit.UI.Result.Failed;
                        }
                    }

                    // ✅ CRITICAL FIX: Check if UI state has changed (new files added) - MUST BE BEFORE using these variables
                    // If new files added → Don't filter by resolved status (process all intersections for new files)
                    // Then filter out resolved zones from Global XML for existing zones
                    // ✅ FIX: Check UI state even if no existing clash zones (for new filters)
                    // ✅ OPTIMIZATION: Identify NEW files to avoid redundant intersection detection for already-processed files
                    // ✅ NOTE: Variables hasNewFilesAdded, newReferenceFiles, newHostFiles, savedRefFiles, savedHostFiles 
                    // are already declared outside if/else block - initializing/clearing here
                    hasNewFilesAdded = false;
                    newReferenceFiles = null;
                    newHostFiles = null;
                    savedRefFiles.Clear();
                    savedHostFiles.Clear();

                    if (filtersToProcess != null && filtersToProcess.Count > 0)
                    {
                        // Compare current UI files with saved files from filters
                        Func<string, string> norm = s =>
                        {
                            if (string.IsNullOrWhiteSpace(s)) return string.Empty;
                            var trimmed = s;
                            var idxParen = trimmed.IndexOf('(');
                            if (idxParen >= 0) trimmed = trimmed.Substring(0, idxParen);
                            trimmed = System.IO.Path.GetFileNameWithoutExtension(trimmed);
                            trimmed = trimmed.ToLowerInvariant().Replace("_detached", "");
                            trimmed = trimmed.Replace('_', ' ').Replace('-', ' ');
                            trimmed = System.Text.RegularExpressions.Regex.Replace(trimmed, "\\s+", " ");
                            return trimmed.Trim();
                        };

                        // Get saved files from filter XML (if available)
                        // Try to get saved files from filtersToProcess (loaded earlier)
                        foreach (var filter in filtersToProcess)
                        {
                            if (filter.SelectedReferenceFiles != null)
                            {
                                foreach (var f in filter.SelectedReferenceFiles.Select(norm))
                                    savedRefFiles.Add(f);
                            }
                            if (filter.SelectedHostFiles != null)
                            {
                                foreach (var f in filter.SelectedHostFiles.Select(norm))
                                    savedHostFiles.Add(f);
                            }
                        }

                        // Get current UI files (with original format for intersection detection)
                        var currentRefFilesNormalized = new HashSet<string>(
                            (selectedReferenceFiles ?? new List<string>()).Select(f => norm(f)),
                            StringComparer.OrdinalIgnoreCase);
                        var currentHostFilesNormalized = new HashSet<string>(
                            (selectedHostFiles ?? new List<string>()).Select(f => norm(f)),
                            StringComparer.OrdinalIgnoreCase);

                        // ✅ CRITICAL FIX: Check which files have clash zones (files that have been processed)
                        // Even if filter was saved with new file, if that file has NO clash zones, it's effectively "new"
                        // This handles user behavior: user can save filter before/after adding files - code will detect correctly
                        var filesWithClashZonesRef = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        var filesWithClashZonesHost = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                        if (existingClashZones?.ClashZones != null && existingClashZones.ClashZones.Count > 0)
                        {
                            // Extract files from clash zones using SourceDocKey or MepElementDocumentTitle
                            foreach (var cz in existingClashZones.ClashZones)
                            {
                                if (cz == null) continue;

                                // Check MEP element source file (reference file)
                                if (!string.IsNullOrEmpty(cz.SourceDocKey))
                                {
                                    var refFileName = System.IO.Path.GetFileNameWithoutExtension(cz.SourceDocKey);
                                    filesWithClashZonesRef.Add(norm(refFileName));
                                }

                                // Check host element source file (host file)
                                if (!string.IsNullOrEmpty(cz.StructuralElementDocumentTitle))
                                {
                                    var hostFileName = System.IO.Path.GetFileNameWithoutExtension(cz.StructuralElementDocumentTitle);
                                    filesWithClashZonesHost.Add(norm(hostFileName));
                                }
                            }

                            if (!DeploymentConfiguration.DeploymentMode && (filesWithClashZonesRef.Count > 0 || filesWithClashZonesHost.Count > 0))
                            {
                                DebugLogger.Info($"[UI-STATE-CHECK] Files with clash zones: Ref={filesWithClashZonesRef.Count} ({string.Join(", ", filesWithClashZonesRef)}), Host={filesWithClashZonesHost.Count} ({string.Join(", ", filesWithClashZonesHost)})");
                            }
                        }

                        // ✅ OPTIMIZATION: Identify NEW files - files that DON'T have clash zones yet
                        // This works even if filter was saved with new file - if no clash zones, it's new
                        newReferenceFiles = (selectedReferenceFiles ?? new List<string>())
                            .Where(f =>
                            {
                                var normalized = norm(f);
                                // If no clash zones exist at all, treat all files as new (first run)
                                if (filesWithClashZonesRef.Count == 0)
                                    return true;
                                // If file doesn't have clash zones, it's new
                                return !filesWithClashZonesRef.Contains(normalized);
                            })
                            .ToList();

                        newHostFiles = (selectedHostFiles ?? new List<string>())
                            .Where(f =>
                            {
                                var normalized = norm(f);
                                // If no clash zones exist at all, treat all files as new (first run)
                                if (filesWithClashZonesHost.Count == 0)
                                    return true;
                                // If file doesn't have clash zones, it's new
                                return !filesWithClashZonesHost.Contains(normalized);
                            })
                            .ToList();

                        // Check if new files were detected (files without clash zones)
                        bool hasNewRefFiles = newReferenceFiles.Count > 0;
                        bool hasNewHostFiles = newHostFiles.Count > 0;

                        // ✅ FALLBACK: Also check if files exist in current UI but not in saved XML (original logic)
                        // This handles edge case where clash zones might not have SourceDocKey set
                        if (!hasNewRefFiles && savedRefFiles.Count > 0)
                        {
                            hasNewRefFiles = currentRefFilesNormalized.Any(f => !savedRefFiles.Contains(f));
                        }
                        if (!hasNewHostFiles && savedHostFiles.Count > 0)
                        {
                            hasNewHostFiles = currentHostFilesNormalized.Any(f => !savedHostFiles.Contains(f));
                        }

                        hasNewFilesAdded = hasNewRefFiles || hasNewHostFiles;
                    }

                    // ✅ CRITICAL FIX: BEFORE intersection detection, mark new file combos in Global XML
                    // This ensures the intersection service knows which combos need processing
                    // Step 1: Detect new files from UI
                    // Step 2: Mark new file combos in Global XML (they're implicitly "new" by not being in ProcessedFileCombos)
                    // Step 3: Intersection service checks Global XML and processes only combos NOT in ProcessedFileCombos
                    // Step 4: After detection, mark combos as processed
                    if (hasNewFilesAdded && (newReferenceFiles != null && newReferenceFiles.Count > 0 || newHostFiles != null && newHostFiles.Count > 0))
                    {
                        if (selectedMepCategories != null && selectedMepCategories.Count > 0)
                        {
                            // Build new file combos from UI selections
                            var newFileCombos = new List<(string LinkedFile, string HostFile)>();
                            var refFilesForCombos = newReferenceFiles ?? selectedReferenceFiles ?? new List<string>();
                            var hostFilesForCombos = newHostFiles ?? selectedHostFiles ?? new List<string>();

                            foreach (var refFile in refFilesForCombos)
                            {
                                foreach (var hostFile in hostFilesForCombos)
                                {
                                    newFileCombos.Add((refFile, hostFile));
                                }
                            }

                            // Mark new file combos in Global XML for each category BEFORE detection
                            // They're "new" by not being in ProcessedFileCombos - but we log this explicitly
                            foreach (var category in selectedMepCategories)
                            {
                                if (string.IsNullOrWhiteSpace(category))
                                    continue;

                                try
                                {
                                    var globalIndex = GlobalIndexService.LoadOrCreate(_document, category);
                                    var processedKeys = GlobalIndexService.GetProcessedFileComboKeys(_document, category);
                                    var processedKeysSet = new HashSet<string>(processedKeys, StringComparer.OrdinalIgnoreCase);

                                    // Check which combos are truly new (not in ProcessedFileCombos)
                                    var trulyNewCombos = newFileCombos.Where(combo =>
                                    {
                                        var comboObj = new ProcessedFileCombo { LinkedFile = combo.LinkedFile, HostFile = combo.HostFile };
                                        var key = comboObj.GetNormalizedKey();
                                        return !processedKeysSet.Contains(key);
                                    }).ToList();

                                    if (trulyNewCombos.Count > 0)
                                    {
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            DebugLogger.Info($"[GLOBAL-INDEX] ✅ STEP 1: Marked {trulyNewCombos.Count} NEW file combos for category '{category}' BEFORE intersection detection:");
                                            foreach (var combo in trulyNewCombos.Take(5))
                                            {
                                                DebugLogger.Info($"[GLOBAL-INDEX]   New combo: LinkedFile='{combo.LinkedFile}', HostFile='{combo.HostFile}'");
                                            }
                                        }
                                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] ✅ STEP 1: Marked {trulyNewCombos.Count} NEW file combos for '{category}' BEFORE detection\n");

                                        // Note: These combos are NOT in ProcessedFileCombos, so they're implicitly "new"
                                        // The intersection service will check ProcessedFileCombos and process only combos NOT in the list
                                    }
                                }
                                catch (Exception comboEx)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Warning($"[GLOBAL-INDEX] Error checking new file combos for category '{category}': {comboEx.Message}");
                                }
                            }
                        }
                    }

                    // Use passed MEP categories, reference files, and host filters from UI (UI PRECEDENCE)
                    // ✅ OOP: Pass optimization service to intersection detection (dependency injection)
                    // ✅ OPTIMIZATION: Only pass NEW files to intersection detection (skip already-processed files)
                    // If new files detected, only process those files - avoids redundant detection for existing files
                    // ✅ NOTE: filesToProcessForRef and filesToProcessForHost are declared outside if/else block - assigning here
                    filesToProcessForRef = hasNewFilesAdded && newReferenceFiles != null && newReferenceFiles.Count > 0
                        ? newReferenceFiles  // Only NEW reference files
                        : selectedReferenceFiles; // All files (no new files or first run)

                    filesToProcessForHost = hasNewFilesAdded && newHostFiles != null && newHostFiles.Count > 0
                        ? newHostFiles  // Only NEW host files
                        : selectedHostFiles; // All files (no new files or first run)

                    // ✅ OPTIMIZATION (OOP): Filter out already-processed file combos from Global XML
                    // Intersection service will check ProcessedFileCombos and process only combos NOT in the list
                    // For each category, check which (LinkedFile, HostFile) combos are already processed
                    // This ensures that even if a file is "new" to this filter, if it was processed in another filter for the same category, we skip it
                    if (selectedMepCategories != null && selectedMepCategories.Count > 0 &&
                        (filesToProcessForRef?.Count > 0 || filesToProcessForHost?.Count > 0))
                    {
                        Func<string, string> norm = s =>
                        {
                            if (string.IsNullOrWhiteSpace(s)) return string.Empty;
                            var trimmed = s;
                            var idxParen = trimmed.IndexOf('(');
                            if (idxParen >= 0) trimmed = trimmed.Substring(0, idxParen);
                            trimmed = System.IO.Path.GetFileNameWithoutExtension(trimmed);
                            trimmed = trimmed.ToLowerInvariant().Replace("_detached", "");
                            trimmed = trimmed.Replace('_', ' ').Replace('-', ' ');
                            trimmed = System.Text.RegularExpressions.Regex.Replace(trimmed, "\\s+", " ");
                            return trimmed.Trim();
                        };

                        // Build list of file combos to process (all combinations of ref × host files)
                        var fileCombosToProcess = new List<(string LinkedFile, string HostFile)>();
                        var refFilesToProcess = filesToProcessForRef ?? new List<string>();
                        var hostFilesToProcess = filesToProcessForHost ?? new List<string>();

                        // If no ref files specified, use all selected ref files
                        if (refFilesToProcess.Count == 0)
                            refFilesToProcess = selectedReferenceFiles ?? new List<string>();
                        // If no host files specified, use all selected host files
                        if (hostFilesToProcess.Count == 0)
                            hostFilesToProcess = selectedHostFiles ?? new List<string>();

                        // Generate all combinations of ref × host files
                        foreach (var refFile in refFilesToProcess)
                        {
                            foreach (var hostFile in hostFilesToProcess)
                            {
                                fileCombosToProcess.Add((refFile, hostFile));
                            }
                        }

                        // Filter out already-processed file combos per category
                        var filteredFileCombos = new List<(string LinkedFile, string HostFile)>();
                        var processedComboKeys = new Dictionary<string, HashSet<string>>(); // category -> set of processed combo keys

                        foreach (var category in selectedMepCategories)
                        {
                            if (string.IsNullOrWhiteSpace(category))
                                continue;

                            // Get processed file combo keys for this category from Global XML (OOP method)
                            var processedKeys = GlobalIndexService.GetProcessedFileComboKeys(_document, category);
                            processedComboKeys[category] = processedKeys;

                            if (!DeploymentConfiguration.DeploymentMode && processedKeys.Count > 0)
                            {
                                DebugLogger.Info($"[GLOBAL-INDEX] Category '{category}': Found {processedKeys.Count} already-processed file combos in Global XML");
                                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] Category '{category}': {processedKeys.Count} processed file combos\n");
                            }
                        }

                        // Filter file combos: keep only those NOT already processed for ANY category
                        foreach (var combo in fileCombosToProcess)
                        {
                            var comboNormalized = new ProcessedFileCombo { LinkedFile = combo.LinkedFile, HostFile = combo.HostFile };
                            var comboKey = comboNormalized.GetNormalizedKey();

                            // Check if this combo is already processed for ANY selected category
                            bool isAlreadyProcessed = false;
                            foreach (var category in selectedMepCategories)
                            {
                                if (string.IsNullOrWhiteSpace(category))
                                    continue;

                                if (processedComboKeys.ContainsKey(category) && processedComboKeys[category].Contains(comboKey))
                                {
                                    isAlreadyProcessed = true;
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        DebugLogger.Info($"[GLOBAL-INDEX] Skipping already-processed file combo for category '{category}': LinkedFile='{combo.LinkedFile}', HostFile='{combo.HostFile}'");
                                    }
                                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] Skipping combo for '{category}': {combo.LinkedFile} × {combo.HostFile}\n");
                                    break;
                                }
                            }

                            if (!isAlreadyProcessed)
                            {
                                filteredFileCombos.Add(combo);
                            }
                        }

                        // Extract unique ref and host files from filtered combos
                        var filteredRefFiles = filteredFileCombos.Select(c => c.LinkedFile).Distinct().ToList();
                        var filteredHostFiles = filteredFileCombos.Select(c => c.HostFile).Distinct().ToList();

                        // Update files to process if any combos were filtered out
                        if (filteredFileCombos.Count < fileCombosToProcess.Count)
                        {
                            int skippedCount = fileCombosToProcess.Count - filteredFileCombos.Count;
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[GLOBAL-INDEX] ⚡ OPTIMIZATION: Filtered out {skippedCount} already-processed file combos (out of {fileCombosToProcess.Count} total)");
                                DebugLogger.Info($"[GLOBAL-INDEX]   Before: {refFilesToProcess.Count} ref files × {hostFilesToProcess.Count} host files = {fileCombosToProcess.Count} combos");
                                DebugLogger.Info($"[GLOBAL-INDEX]   After: {filteredRefFiles.Count} ref files × {filteredHostFiles.Count} host files = {filteredFileCombos.Count} combos");
                                DebugLogger.Info($"[GLOBAL-INDEX]   ✅ Note: Even though detection is skipped for processed combos, Global XML will still be checked for unresolved zones");
                            }
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] ⚡ Skipped {skippedCount} processed combos: {filteredFileCombos.Count}/{fileCombosToProcess.Count} combos remaining\n");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] ✅ Note: Global XML will be checked for unresolved zones from processed combos\n");

                            filesToProcessForRef = filteredRefFiles;
                            filesToProcessForHost = filteredHostFiles;
                        }

                        // ✅ CRITICAL: If ALL file combos are already processed (filteredFileCombos.Count == 0),
                        // we skip intersection detection but Global XML will still be checked for unresolved zones
                        // This ensures that even if a file combo was processed before, unresolved zones from that combo
                        // (e.g., after sleeve deletion) will still be detected and processed
                        if (filteredFileCombos.Count == 0 && fileCombosToProcess.Count > 0)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[GLOBAL-INDEX] ⚡ ALL file combos already processed - skipping intersection detection, but will check Global XML for unresolved zones");
                            }
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] ⚡ ALL combos processed - skipping detection, checking Global XML for unresolved zones\n");
                        }
                    }

                    if (hasNewFilesAdded && (newReferenceFiles != null && newReferenceFiles.Count > 0 || newHostFiles != null && newHostFiles.Count > 0))
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[OPTIMIZATION] ⚡ Running intersection detection ONLY for NEW files:");
                            DebugLogger.Info($"[OPTIMIZATION]   Reference files: {filesToProcessForRef?.Count ?? 0} (saved: {savedRefFiles?.Count ?? 0}, current: {selectedReferenceFiles?.Count ?? 0})");
                            DebugLogger.Info($"[OPTIMIZATION]   Host files: {filesToProcessForHost?.Count ?? 0} (saved: {savedHostFiles?.Count ?? 0}, current: {selectedHostFiles?.Count ?? 0})");
                            DebugLogger.Info($"[OPTIMIZATION]   Skipping intersection detection for already-processed files - will merge with existing clash zones");
                        }
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [OPTIMIZATION] ⚡ Running intersection detection ONLY for NEW files (ref: {filesToProcessForRef?.Count ?? 0}, host: {filesToProcessForHost?.Count ?? 0})\n");
                    }

                    currentIntersections = intersectionService.FindIntersections(
                        _document,
                        view3D,
                        selectedMepCategories,
                        filesToProcessForRef,  // ✅ OPTIMIZATION: Filtered by processed file combos
                        filesToProcessForHost,  // ✅ OPTIMIZATION: Filtered by processed file combos
                        allowedHostTypesUI.ToList(),
                        optimizationService); // ✅ OOP: Pass optimization service instead of raw data

                    // ⚠️ CRITICAL: Check limits after intersection detection
                    try
                    {
                        _memoryManager.CheckLimits();
                    }
                    catch (TimeoutException ex)
                    {
                        SafeFileLogger.SafeAppendText("refresh_timeouts.log", $"Timeout after intersection detection: {ex.Message}");
                        _statusLabel.Text = $"Operation timeout after finding {currentIntersections.Count} intersections";
                        _progressBar.Visible = false;
                        _refreshButton.Enabled = true;
                        System.Windows.Forms.MessageBox.Show(
                            $"Found {currentIntersections.Count} intersections but operation timed out.\n\n{ex.Message}",
                            "Operation Timeout",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Warning);
                        return Autodesk.Revit.UI.Result.Cancelled;
                    }
                    catch (OutOfMemoryException ex)
                    {
                        SafeFileLogger.SafeAppendText("refresh_memory.log", $"Memory limit after intersection detection: {ex.Message}");
                        _statusLabel.Text = $"Memory limit exceeded after finding {currentIntersections.Count} intersections";
                        _progressBar.Visible = false;
                        _refreshButton.Enabled = true;
                        System.Windows.Forms.MessageBox.Show(
                            $"Found {currentIntersections.Count} intersections but memory limit exceeded.\n\n{ex.Message}",
                            "Memory Limit Exceeded",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Warning);
                        return Autodesk.Revit.UI.Result.Failed;
                    }

                    // Log intersection details after detection
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] INTERSECTION DETECTION COMPLETE: {currentIntersections.Count} intersections found");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] INTERSECTION DETECTION COMPLETE: {currentIntersections.Count} intersections found\n");
                    if (currentIntersections.Count > 0)
                    {
                        if (OptimizationFlags.UseDiagnosticMode)
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] INTERSECTION DETECTION RESULTS:\n");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] Total intersections found: {currentIntersections.Count}\n");
                        if (OptimizationFlags.UseDiagnosticMode)
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] Intersection details:\n");

                        // Log first 20 intersections
                        int logCount = Math.Min(20, currentIntersections.Count);
                        for (int i = 0; i < logCount; i++)
                        {
                            var (mepElement, structuralElement, intersectionBBox, intersectionPoint) = currentIntersections[i];
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}]   Intersection {i + 1}:\n");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}]     MEP Element ID: {mepElement.Id}\n");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}]     MEP Element Name: {mepElement.Name}\n");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}]     MEP Category: {mepElement.Category?.Name ?? "Unknown"}\n");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}]     Structural Element ID: {structuralElement.Id}\n");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}]     Structural Element Name: {structuralElement.Name}\n");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}]     Structural Category: {structuralElement.Category?.Name ?? "Unknown"}\n");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}]     Intersection Point: ({intersectionPoint.X:F2}, {intersectionPoint.Y:F2}, {intersectionPoint.Z:F2})\n");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}]     Intersection BBox: Min({intersectionBBox.Min.X:F2}, {intersectionBBox.Min.Y:F2}, {intersectionBBox.Min.Z:F2}) Max({intersectionBBox.Max.X:F2}, {intersectionBBox.Max.Y:F2}, {intersectionBBox.Max.Z:F2})\n");
                        }

                        if (currentIntersections.Count > 20)
                        {
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}]   ... and {currentIntersections.Count - 20} more intersections\n");
                        }
                    }

                    // ✅ MEMORY OPTIMIZATION: Clear Revit API objects for all loaded existing clash zones
                    // XML deserialization may create XYZ objects when setters are accessed
                    if (existingClashZones?.ClashZones != null)
                    {
                        foreach (var cz in existingClashZones.ClashZones)
                        {
                            cz.ClearRevitApiObjects();
                        }
                    }

                    // ✅ CRITICAL FIX: ClashZoneService logging must write to refresh log file
                    // ✅ MEMORY OPTIMIZATION: Use batched logger to reduce string allocations and I/O
                    var baseLogger = new Action<string>(msg =>
                    {
                        // Only write to Refresh_* log; avoid DebugLogger to keep logging limited
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] {msg}\n");
                    });

                    // Flush each message immediately to avoid buffered string buildup
                    batchedLogger = new BatchedLogger(baseLogger, batchSize: 1, flushIntervalSeconds: 1, logFileName: refreshLogName);
                    // ✅ OOP: Pass GuidManager to ClashZoneService for GUID lookup from Revit sleeves
                    _clashZoneService = new ClashZoneService(existingClashZones, msg => batchedLogger.Log(msg), _flagManager, _guidManager);

                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] ClashZoneService reinitialized with {existingCount} existing zones from XML");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ClashZoneService reinitialized with {existingCount} existing zones from XML\n");

                    // ✅ STEP 3a: Create Known Intersections Map from valid clash zones (PERFORMANCE OPTIMIZATION)
                    // This allows filtering out already-known intersections AFTER detection, so we only process NEW ones
                    // ✅ FLAG MANAGEMENT: Only include zones that passed 3-point validation AND are NOT cluster resolved
                    // Cluster resolved zones are excluded from the map (flags already set by FlagManager.SyncFlagsFromGlobal)
                    // ✅ NOTE: knownIntersectionsMap is declared outside if/else block - clearing and initializing here
                    knownIntersectionsMap.Clear();
                    var validatedClashZones = existingClashZones?.ClashZones ?? new List<Models.ClashZone>();

                    // Filter out cluster resolved zones (don't need to check them - flags already managed by FlagManager)
                    var zonesForIntersectionDetection = validatedClashZones
                        .Where(cz => cz != null && !cz.IsClusterResolved)
                        .ToList();
                    int clusterResolvedSkippedCount = validatedClashZones.Count - zonesForIntersectionDetection.Count;

                    // ✅ ALWAYS create the map (intersection detection already completed above)
                    if (zonesForIntersectionDetection.Count > 0)
                    {
                        foreach (var cz in zonesForIntersectionDetection)
                        {
                            int mepId = cz.MepElementId?.IntegerValue ?? cz.MepElementIdValue;
                            int hostId = cz.StructuralElementId?.IntegerValue ?? cz.StructuralElementIdValue;

                            // Create point key with tolerance (round to 0.1ft for matching)
                            double tolerance = 0.1;
                            string pointKey = $"{Math.Round(cz.IntersectionPointX / tolerance) * tolerance:F1}," +
                                             $"{Math.Round(cz.IntersectionPointY / tolerance) * tolerance:F1}," +
                                             $"{Math.Round(cz.IntersectionPointZ / tolerance) * tolerance:F1}";

                            var key = (mepId, hostId, pointKey);
                            if (!knownIntersectionsMap.ContainsKey(key))
                            {
                                knownIntersectionsMap[key] = cz;
                            }
                        }
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[PERFORMANCE] Created known intersections map: {knownIntersectionsMap.Count} entries (skipping re-detection), {clusterResolvedSkippedCount} cluster resolved zones skipped");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [PERFORMANCE] Known intersections map: {knownIntersectionsMap.Count} entries, {clusterResolvedSkippedCount} cluster resolved skipped\n");
                    }

                    // Step 7: Filter and detect new clash zones
                    _progressBar.Value = 50;
                    _statusLabel.Text = "Detecting new clash zones...";

                    // ✅ PERFORMANCE OPTIMIZATION: Filter out already-known intersections AFTER detection
                    // This skips expensive clash zone creation for unchanged intersections
                    // BUT still allows detection of NEW clashes from new/modified elements

                    // ✅ CRITICAL FIX: Check Global XML for resolved clash zones by MEP+Host+Point BEFORE creating clash zones
                    // This prevents creating clash zones for intersections that are already resolved in Global XML
                    // Uses O(1) Dictionary lookup by MEP+Host+Point (not GUID) - enables cross-filter matching
                    // ✅ UI STATE CHECK: If new files added (host or reference), skip resolved filtering to process all intersections
                    // ✅ NOTE: resolvedIntersectionPointsFromGlobalXml is declared outside if/else block - clearing and initializing here
                    resolvedIntersectionPointsFromGlobalXml.Clear();

                    // ✅ NOTE: Variables hasNewFilesAdded, newReferenceFiles, newHostFiles, savedRefFiles, savedHostFiles 
                    // are already declared earlier (around line 1544) - using those declarations here

                    if (hasNewFilesAdded)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[UI-STATE-CHECK] ✅ NEW FILES DETECTED - Processing ALL intersections from new files (skip resolved filtering)");
                        }
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [UI-STATE-CHECK] ✅ NEW FILES DETECTED: Processing ALL intersections from new files\n");
                    }
                    else
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[UI-STATE-CHECK] No new files added - will filter resolved zones from Global XML");
                    }

                    try
                    {
                        // ✅ ONLY build resolved index if NO new files added
                        // If new files added → Process all intersections (don't filter by resolved status)
                        if (!hasNewFilesAdded)
                        {
                            // ✅ Build index from ALL Global XML entries with resolved flags (not just ones matching Filter XML)
                            // This handles case where Filter B detects clash zone that Filter A already placed sleeves for
                            foreach (var category in selectedMepCategories ?? new List<string>())
                            {
                                try
                                {
                                    var globalIndex = GlobalIndexService.BuildMepHostPointIndex(_document, category);

                                    // ✅ SIMPLE LOGIC: Only include entries with resolved flags = true
                                    // If flags are false, entry won't be in resolvedIntersectionPointsFromGlobalXml, so clash zone will be processed
                                    // ResetFlagsForDeletedSleeves already checks sleeve existence and sets flags to false when sleeves are deleted
                                    foreach (var kvp in globalIndex)
                                    {
                                        var entry = kvp.Value;

                                        // ✅ CRITICAL LOGIC: Skip ONLY if BOTH flags are true
                                        // Process if ANY flag is false (IsResolved == false OR IsClusterResolved == false)
                                        if (entry.IsResolved && entry.IsClusterResolved)
                                        {
                                            resolvedIntersectionPointsFromGlobalXml[kvp.Key] = entry; // Add to skip list
                                        }
                                        // If any flag is false → don't add to skip list → process clash zone
                                    }
                                }
                                catch (Exception categoryEx)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        DebugLogger.Warning($"[REFRESH-GLOBAL-FILTER] Error building index for category '{category}': {categoryEx.Message}");
                                    }
                                }
                            }

                            if (resolvedIntersectionPointsFromGlobalXml.Count > 0 && !DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[REFRESH-GLOBAL-FILTER] Built resolved intersections index: {resolvedIntersectionPointsFromGlobalXml.Count} (MEP+Host+Point) entries already resolved in Global XML");
                                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [REFRESH-GLOBAL-FILTER] Built resolved index: {resolvedIntersectionPointsFromGlobalXml.Count} entries\n");
                            }
                        }
                        else
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[REFRESH-GLOBAL-FILTER] Skipping resolved filtering - new files added, will process all intersections");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [REFRESH-GLOBAL-FILTER] Skipping resolved filtering - new files added\n");
                        }
                    }
                    catch (Exception globalFilterEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[REFRESH-GLOBAL-FILTER] Error building resolved intersections index: {globalFilterEx.Message}");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [REFRESH-GLOBAL-FILTER] ERROR: {globalFilterEx.Message}\n");
                        }
                    }

                    // ✅ PATH 2 OPTIMIZATION: When "Adopt Document" is enabled, check unresolved zones against detected intersections
                    // If 3 points match (MEP ID, Host ID, Intersection Point) → update GUID/data and skip re-detection
                    // If NOT identical → process as new clash zone
                    if (enableThreePointValidation && currentIntersections.Count > 0 && existingClashZones?.ClashZones != null)
                    {
                        var unresolvedZones = existingClashZones.ClashZones
                            .Where(cz => cz != null && !cz.IsClusterResolved && !cz.IsResolved)
                            .ToList();

                        if (unresolvedZones.Count > 0)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[ADOPT-DOCUMENT] Checking {unresolvedZones.Count} unresolved zones against {currentIntersections.Count} detected intersections");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [ADOPT-DOCUMENT] Checking {unresolvedZones.Count} unresolved zones against {currentIntersections.Count} detected intersections\n");

                            double tolerance = 0.1;
                            int updatedCount = 0;
                            int processedAsNewCount = 0;
                            var intersectionsToRemove = new HashSet<(Element, Element, BoundingBoxXYZ, XYZ)>();

                            foreach (var unresolvedZone in unresolvedZones)
                            {
                                try
                                {
                                    int mepId = unresolvedZone.MepElementId?.IntegerValue ?? unresolvedZone.MepElementIdValue;
                                    int hostId = unresolvedZone.StructuralElementId?.IntegerValue ?? unresolvedZone.StructuralElementIdValue;

                                    // Create point key from existing zone's intersection point
                                    string existingPointKey = $"{Math.Round(unresolvedZone.IntersectionPointX / tolerance) * tolerance:F1}," +
                                                             $"{Math.Round(unresolvedZone.IntersectionPointY / tolerance) * tolerance:F1}," +
                                                             $"{Math.Round(unresolvedZone.IntersectionPointZ / tolerance) * tolerance:F1}";

                                    // Find matching intersection in detected intersections
                                    var matchingIntersection = currentIntersections.FirstOrDefault(i =>
                                    {
                                        int iMepId = i.Item1.Id.IntegerValue;
                                        int iHostId = i.Item2.Id.IntegerValue;
                                        XYZ iPoint = i.Item4;

                                        // Check MEP ID and Host ID match
                                        if (iMepId != mepId || iHostId != hostId)
                                            return false;

                                        // Check intersection point matches (with tolerance)
                                        string iPointKey = $"{Math.Round(iPoint.X / tolerance) * tolerance:F1}," +
                                                          $"{Math.Round(iPoint.Y / tolerance) * tolerance:F1}," +
                                                          $"{Math.Round(iPoint.Z / tolerance) * tolerance:F1}";

                                        return iPointKey == existingPointKey;
                                    });

                                    if (matchingIntersection.Item1 != null)
                                    {
                                        // ✅ 3 POINTS IDENTICAL: Update GUID and 3-point data, skip re-detection
                                        var (mepElement, structuralElement, bbox, point) = matchingIntersection;

                                        // Update intersection point if it changed slightly (within tolerance)
                                        if (Math.Abs(unresolvedZone.IntersectionPointX - point.X) > 1e-6 ||
                                            Math.Abs(unresolvedZone.IntersectionPointY - point.Y) > 1e-6 ||
                                            Math.Abs(unresolvedZone.IntersectionPointZ - point.Z) > 1e-6)
                                        {
                                            unresolvedZone.IntersectionPointX = point.X;
                                            unresolvedZone.IntersectionPointY = point.Y;
                                            unresolvedZone.IntersectionPointZ = point.Z;
                                            unresolvedZone.IntersectionPoint = new XYZ(point.X, point.Y, point.Z);

                                            if (!DeploymentConfiguration.DeploymentMode)
                                                DebugLogger.Info($"[ADOPT-DOCUMENT] ✅ Updated intersection point for unresolved zone {unresolvedZone.Id}: ({point.X:F3},{point.Y:F3},{point.Z:F3})");
                                        }

                                        // Update Global XML with GUID and 3-point data
                                        var category = unresolvedZone.MepElementCategory;
                                        if (!string.IsNullOrWhiteSpace(category))
                                        {
                                            try
                                            {
                                                // ✅ CRITICAL: Pass filter name so Global XML knows which Filter XML file contains placement data
                                                // Find the filter that contains this unresolved zone
                                                var containingFilter = filtersToProcess.FirstOrDefault(f =>
                                                    f.ClashZoneStorage?.AllZones?.Any(cz => cz.Id == unresolvedZone.Id) == true);
                                                // ✅ CRITICAL FIX: Use containingFilter or fallback to enabled filter (targetFilter not in scope here)
                                                // ✅ CRITICAL FIX: Add category suffix to filter name format: "FilterName_category.xml"
                                                var baseFilterName = containingFilter?.Name ?? filtersToProcess.FirstOrDefault(f => f.IsEnabled)?.Name ?? string.Empty;
                                                var filterName = !string.IsNullOrEmpty(baseFilterName) && !string.IsNullOrWhiteSpace(category)
                                                    ? $"{baseFilterName}_{category.ToLower()}.xml"
                                                    : baseFilterName;
                                                _guidManager.EnsureGlobalXmlEntry(unresolvedZone, category, filterName);

                                                if (!DeploymentConfiguration.DeploymentMode)
                                                    DebugLogger.Info($"[ADOPT-DOCUMENT] ✅ Updated Global XML entry for unresolved zone {unresolvedZone.Id} (MEP={mepId}, Host={hostId}, Point=({point.X:F3},{point.Y:F3},{point.Z:F3}), Filter='{filterName ?? "N/A"}')");
                                                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [ADOPT-DOCUMENT] ✅ Updated Global XML for zone {unresolvedZone.Id}\n");
                                            }
                                            catch (Exception globalEx)
                                            {
                                                if (!DeploymentConfiguration.DeploymentMode)
                                                    DebugLogger.Warning($"[ADOPT-DOCUMENT] Error updating Global XML for zone {unresolvedZone.Id}: {globalEx.Message}");
                                            }
                                        }

                                        // Mark intersection to be removed from currentIntersections (skip re-detection)
                                        intersectionsToRemove.Add(matchingIntersection);
                                        updatedCount++;
                                    }
                                    else
                                    {
                                        // ✅ 3 POINTS NOT IDENTICAL: Process as new clash zone
                                        // The intersection will go through normal detection flow
                                        processedAsNewCount++;

                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[ADOPT-DOCUMENT] ⚠️ Unresolved zone {unresolvedZone.Id} (MEP={mepId}, Host={hostId}) - 3 points changed, will process as new clash zone");
                                    }
                                }
                                catch (Exception zoneEx)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Warning($"[ADOPT-DOCUMENT] Error checking unresolved zone {unresolvedZone?.Id}: {zoneEx.Message}");
                                }
                            }

                            // Remove matched intersections from currentIntersections (skip re-detection for these)
                            if (intersectionsToRemove.Count > 0)
                            {
                                currentIntersections = currentIntersections
                                    .Where(i => !intersectionsToRemove.Contains(i))
                                    .ToList();
                            }

                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[ADOPT-DOCUMENT] Result: {updatedCount} unresolved zones updated (3 points identical), {processedAsNewCount} will process as new (3 points changed), {currentIntersections.Count} intersections remaining");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [ADOPT-DOCUMENT] Result: {updatedCount} updated, {processedAsNewCount} as new, {currentIntersections.Count} remaining\n");
                        }
                    }
                }

                // ✅ PATH HANDLING: If PATH 1 (no intersection detection), currentIntersections is empty
                // In this case, we'll use existing clash zones directly without creating new ones
                if (currentIntersections.Count > 0)
                {
                    // ✅ PATH 2: Filter detected intersections
                    if (knownIntersectionsMap.Count > 0 || resolvedIntersectionPointsFromGlobalXml.Count > 0)
                    {
                        int skippedCount = 0;
                        foreach (var intersection in currentIntersections)
                        {
                            var (mepElement, structuralElement, bbox, point) = intersection;
                            int mepId = mepElement.Id.IntegerValue;
                            int hostId = structuralElement.Id.IntegerValue;

                            // ✅ CRITICAL FIX: Check if intersection point matches resolved clash zone in Global XML
                            // This handles case where Filter B detects clash zone that Filter A already placed sleeves for
                            // Uses O(1) Dictionary lookup by MEP+Host+Point (not GUID) - enables cross-filter matching
                            double tolerance = 0.1;
                            string pointKey = $"{Math.Round(point.X / tolerance) * tolerance:F1}," +
                                             $"{Math.Round(point.Y / tolerance) * tolerance:F1}," +
                                             $"{Math.Round(point.Z / tolerance) * tolerance:F1}";
                            var key = (mepId, hostId, pointKey);

                            if (resolvedIntersectionPointsFromGlobalXml.ContainsKey(key))
                            {
                                // Intersection point matches resolved clash zone in Global XML - skip creating new clash zone
                                skippedResolvedCount++;
                                continue; // Skip this intersection - already resolved in Global XML with matching point
                            }

                            // Then check known intersections map (performance optimization)
                            if (knownIntersectionsMap.ContainsKey(key))
                            {
                                skippedCount++;
                                continue; // Skip this intersection - already known (performance optimization)
                            }

                            filteredIntersections.Add(intersection); // NEW intersection - keep it for clash zone creation
                        }
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[PERFORMANCE] Filtered intersections: {currentIntersections.Count} total, {skippedCount} skipped (known), {skippedResolvedCount} skipped (resolved in Global XML), {filteredIntersections.Count} NEW clashes to process");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [PERFORMANCE] Filtered: {currentIntersections.Count} total, {skippedCount} skipped (known), {skippedResolvedCount} skipped (resolved), {filteredIntersections.Count} NEW\n");
                    }
                    else
                    {
                        // No known intersections - process all (first run or all zones invalidated)
                        filteredIntersections = currentIntersections;
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[PERFORMANCE] No known intersections map - processing all {currentIntersections.Count} intersections as NEW clashes");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [PERFORMANCE] No known intersections - processing all {currentIntersections.Count} as NEW\n");
                    }
                }
                else
                {
                    // ✅ PATH 1: No intersection detection - filteredIntersections is empty, will use existing clash zones directly
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[PERFORMANCE] PATH 1: No intersection detection - will use {existingClashZones?.ClashZones?.Count ?? 0} existing clash zones from Filter XML");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [PERFORMANCE] PATH 1: No intersection detection - using {existingClashZones?.ClashZones?.Count ?? 0} existing clash zones\n");
                }

                // ✅ CRITICAL FIX: Track if all intersections were filtered out because resolved
                // This will be used later to show "all resolved" message when totalCount == 0
                bool allIntersectionsFilteredAsResolved = currentIntersections.Count > 0 &&
                                                           filteredIntersections.Count == 0 &&
                                                           skippedResolvedCount > 0;

                // Use passed file selections and clearance settings from UI

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info("[CLASH_DEBUG] Skipping cleanup - no cleanup needed");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Skipping cleanup - no cleanup needed\n");

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] Using loaded existing clash zones - Reference files: {selectedReferenceFiles.Count}, Clearance settings: {clearanceSettings.Count}");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Using loaded existing clash zones - Reference files: {selectedReferenceFiles.Count}, Clearance settings: {clearanceSettings.Count}\n");

                // ✅ CRITICAL FIX: DO NOT call FilterClashZonesByCurrentSelection here!
                // It will overwrite the loaded clash zones with empty storage
                // Instead, use the loaded existing clash zones directly
                var filteredClashZones = existingClashZones;

                // ✅ CRITICAL: Filter existing zones by CURRENT UI selections (categories, host types, reference files, host files)
                // This ensures when user changes UI selections, only matching clash zones are processed
                try
                {
                    var allowedMepCats = new HashSet<string>(selectedMepCategories ?? new List<string>(), StringComparer.OrdinalIgnoreCase);
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] Selected MEP categories from UI: [{string.Join(", ", selectedMepCategories ?? new List<string>())}]");
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] Allowed MEP categories: [{string.Join(", ", allowedMepCats)}]");
                    var allowedHostCategories = new HashSet<string>(
                        FilterUiStateProvider.GetSelectedHostCategories?.Invoke() ?? new List<string>(),
                        StringComparer.OrdinalIgnoreCase);

                    // Normalize file names for comparison
                    Func<string, string> norm = s =>
                    {
                        if (string.IsNullOrWhiteSpace(s)) return string.Empty;
                        var trimmed = s;
                        var idxParen = trimmed.IndexOf('(');
                        if (idxParen >= 0) trimmed = trimmed.Substring(0, idxParen);
                        trimmed = System.IO.Path.GetFileNameWithoutExtension(trimmed);
                        trimmed = trimmed.ToLowerInvariant().Replace("_detached", "");
                        trimmed = trimmed.Replace('_', ' ').Replace('-', ' ');
                        trimmed = System.Text.RegularExpressions.Regex.Replace(trimmed, "\\s+", " ");
                        return trimmed.Trim();
                    };

                    var allowedRefFiles = new HashSet<string>(
                        (selectedReferenceFiles ?? new List<string>()).Select(f => norm(f)),
                        StringComparer.OrdinalIgnoreCase);
                    var allowedHostFiles = new HashSet<string>(
                        (selectedHostFiles ?? new List<string>()).Select(f => norm(f)),
                        StringComparer.OrdinalIgnoreCase);

                    if (existingClashZones?.ClashZones != null)
                    {
                        var before = existingClashZones.ClashZones.Count;

                        // 🔥 CRITICAL DEBUG: Log flags BEFORE filtering
                        int clusterResolvedBefore = existingClashZones.ClashZones.Count(cz => cz.IsClusterResolved);
                        int individualResolvedBefore = existingClashZones.ClashZones.Count(cz => cz.IsResolved);
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CLASH_DEBUG] BEFORE FILTERING: {before} clash zones, IsClusterResolved=True: {clusterResolvedBefore}, IsResolved=True: {individualResolvedBefore}");

                        existingClashZones.ClashZones = existingClashZones.ClashZones.Where(cz =>
                        {
                            // Filter by MEP category
                            bool categoryMatch = allowedMepCats.Count == 0 || allowedMepCats.Contains(cz.MepElementCategory);

                            // Filter by host type - Handle plural/singular mismatch: "Walls" (UI) vs "Wall" (Revit)
                            bool hostTypeMatch = allowedHostCategories.Count == 0 ||
                                               allowedHostCategories.Contains(cz.StructuralElementType) ||
                                               allowedHostCategories.Contains(cz.StructuralElementType + "s") ||
                                               allowedHostCategories.Any(t => t.TrimEnd('s').Equals(cz.StructuralElementType, StringComparison.OrdinalIgnoreCase));

                            // 🔥 CRITICAL DEBUG: Log individual clash zone filtering
                            if (!categoryMatch)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[CLASH_DEBUG] FILTERED OUT: ClashZone {cz.Id} - Category mismatch: '{cz.MepElementCategory}' not in [{string.Join(", ", allowedMepCats)}]");
                                // ✅ DEPLOYMENT MODE: Skip hardcoded log writes
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    string refresh_debug_logLogPath = SafeFileLogger.GetLogFilePath("refresh_debug.log");
                                    File.AppendAllText(refresh_debug_logLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] FILTERED OUT: ClashZone {cz.Id} - Category mismatch: '{cz.MepElementCategory}' not in [{string.Join(", ", allowedMepCats)}]\n");
                                }
                            }
                            if (!hostTypeMatch)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[CLASH_DEBUG] FILTERED OUT: ClashZone {cz.Id} - Host type mismatch: '{cz.StructuralElementType}' not in [{string.Join(", ", allowedHostCategories)}]");
                                // ✅ DEPLOYMENT MODE: Skip hardcoded log writes
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    string refresh_debug_logLogPath = SafeFileLogger.GetLogFilePath("refresh_debug.log");
                                    File.AppendAllText(refresh_debug_logLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] FILTERED OUT: ClashZone {cz.Id} - Host type mismatch: '{cz.StructuralElementType}' not in [{string.Join(", ", allowedHostCategories)}]\n");
                                }
                            }

                            // Filter by reference file (MEP source) - only if reference files are specified
                            // ✅ CRITICAL FIX: Extract filename from full path for comparison
                            string refFileName = string.IsNullOrEmpty(cz.SourceDocKey) ? "" : Path.GetFileNameWithoutExtension(cz.SourceDocKey);
                            bool refFileMatch = allowedRefFiles.Count == 0 || allowedRefFiles.Contains(norm(refFileName));

                            // Filter by host file (structural source) - only if host files are specified
                            // ✅ CRITICAL FIX: Extract filename from full path for comparison
                            string hostFileName = string.IsNullOrEmpty(cz.StructuralElementDocumentTitle) ? "" : Path.GetFileNameWithoutExtension(cz.StructuralElementDocumentTitle);
                            bool hostFileMatch = allowedHostFiles.Count == 0 || allowedHostFiles.Contains(norm(hostFileName));

                            // 🔥 CRITICAL DEBUG: Log file filtering
                            if (!refFileMatch)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[CLASH_DEBUG] FILTERED OUT: ClashZone {cz.Id} - Reference file mismatch: '{cz.SourceDocKey}' -> filename: '{refFileName}' -> normalized: '{norm(refFileName)}' not in [{string.Join(", ", allowedRefFiles)}]");
                                // ✅ DEPLOYMENT MODE: Skip hardcoded log writes
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    string refresh_debug_logLogPath = SafeFileLogger.GetLogFilePath("refresh_debug.log");
                                    File.AppendAllText(refresh_debug_logLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] FILTERED OUT: ClashZone {cz.Id} - Reference file mismatch: '{refFileName}' not in [{string.Join(", ", allowedRefFiles)}]\n");
                                }
                            }
                            if (!hostFileMatch)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[CLASH_DEBUG] FILTERED OUT: ClashZone {cz.Id} - Host file mismatch: '{cz.StructuralElementDocumentTitle}' -> filename: '{hostFileName}' -> normalized: '{norm(hostFileName)}' not in [{string.Join(", ", allowedHostFiles)}]");
                                // ✅ DEPLOYMENT MODE: Skip hardcoded log writes
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    string refresh_debug_logLogPath = SafeFileLogger.GetLogFilePath("refresh_debug.log");
                                    File.AppendAllText(refresh_debug_logLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] FILTERED OUT: ClashZone {cz.Id} - Host file mismatch: '{hostFileName}' not in [{string.Join(", ", allowedHostFiles)}]\n");
                                }
                            }

                            return categoryMatch && hostTypeMatch && refFileMatch && hostFileMatch;
                        }).ToList();

                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CLASH_DEBUG] Current UI filter: MEP cats={allowedMepCats.Count}, Host types={allowedHostCategories.Count}, Ref files={allowedRefFiles.Count}, Host files={allowedHostFiles.Count}");
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CLASH_DEBUG] Filtered existing zones by current UI: {before} -> {existingClashZones.ClashZones.Count}");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Filtered existing zones by current UI: {before} -> {existingClashZones.ClashZones.Count}\n");
                        __logPhaseMem("AFTER_UI_FILTER_EXISTING", existingClashZones?.ClashZones, null, null);

                        // 🔥 CRITICAL DEBUG: Log flags AFTER filtering
                        int clusterResolvedAfter = existingClashZones.ClashZones.Count(cz => cz.IsClusterResolved);
                        int individualResolvedAfter = existingClashZones.ClashZones.Count(cz => cz.IsResolved);
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CLASH_DEBUG] AFTER FILTERING: {existingClashZones.ClashZones.Count} clash zones, IsClusterResolved=True: {clusterResolvedAfter}, IsResolved=True: {individualResolvedAfter}");
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CLASH_DEBUG] [{DateTime.Now}]  AFTER FILTERING: {existingClashZones.ClashZones.Count} clash zones, IsClusterResolved=True: {clusterResolvedAfter}, IsResolved=True: {individualResolvedAfter}\n");
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[CLASH_DEBUG] Current-selection filter failed: {ex.Message}");
                }

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] Calling DetectNewClashZones with {filteredIntersections.Count} filtered intersections (from {currentIntersections.Count} total)...");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Calling DetectNewClashZones with {filteredIntersections.Count} filtered intersections (from {currentIntersections.Count} total)...\n");

                // ✅ CRITICAL FIX: Filter intersections by section box BEFORE creating clash zones
                // This ensures ONLY clash zones within section box are created, not all in the model
                // ✅ REUSE: view3D is already declared at line 1717 - check if section box is active
                if (view3D != null && view3D.IsSectionBoxActive)
                {
                    try
                    {
                        var secBox = view3D.GetSectionBox();
                        if (secBox != null && secBox.Min != null && secBox.Max != null)
                        {
                            var inv = secBox.Transform.Inverse;
                            int beforeCount = filteredIntersections.Count;

                            filteredIntersections = filteredIntersections
                                .Where(intersection =>
                                {
                                    var intersectionPoint = intersection.Item4;
                                    var local = inv.OfPoint(intersectionPoint);
                                    bool inside = local.X >= secBox.Min.X && local.X <= secBox.Max.X &&
                                                  local.Y >= secBox.Min.Y && local.Y <= secBox.Max.Y &&
                                                  local.Z >= secBox.Min.Z && local.Z <= secBox.Max.Z;
                                    return inside;
                                })
                                .ToList();

                            int afterCount = filteredIntersections.Count;
                            int filteredOut = beforeCount - afterCount;

                            if (filteredOut > 0)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[SECTION-BOX-FILTER] ✅ Filtered {filteredOut} intersections outside section box ({beforeCount} → {afterCount}) - ONLY zones within section box will be created");
                                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [SECTION-BOX-FILTER] ✅ Filtered {filteredOut} intersections outside section box ({beforeCount} → {afterCount})\n");
                            }
                            else
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[SECTION-BOX-FILTER] ✅ All {afterCount} intersections are within section box - proceeding with clash zone creation");
                            }
                        }
                    }
                    catch (Exception secBoxEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[SECTION-BOX-FILTER] Error filtering intersections by section box: {secBoxEx.Message} - proceeding with all intersections");
                    }
                }
                else
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[SECTION-BOX-FILTER] ⚠️ No section box active - will create clash zones for ALL intersections in model");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [SECTION-BOX-FILTER] ⚠️ No section box active - creating zones for all intersections\n");
                }

                // DEBUG: Log intersection breakdown by category
                var intersectionBreakdown = currentIntersections.GroupBy(i => GetElementCategory(i.Item1))
                    .ToDictionary(g => g.Key, g => g.Count());
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] Intersection breakdown by category: {string.Join(", ", intersectionBreakdown.Select(kv => $"{kv.Key}={kv.Value}"))}");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Intersection breakdown by category: {string.Join(", ", intersectionBreakdown.Select(kv => $"{kv.Key}={kv.Value}"))}\n");

                // ✅ DEBUG: Log Duct-Wall intersections BEFORE optimization
                var ductWallIntersections = currentIntersections.Where(i =>
                {
                    var mepCat = GetElementCategory(i.Item1);
                    var structType = i.Item2.Category?.Name ?? "";
                    return (mepCat == "Ducts" || mepCat == "Duct Curves" || mepCat == "Duct Accessories")
                        && (structType == "Walls" || structType == "Wall");
                }).ToList();
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] ═══ BEFORE OPTIMIZATION: {ductWallIntersections.Count} Duct-Wall intersections found ═══");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ═══ BEFORE OPTIMIZATION: {ductWallIntersections.Count} Duct-Wall intersections (Total intersections: {currentIntersections.Count}) ═══\n");

                // ✅ CRITICAL PROTECTED LOGIC: Check if (MEP ID, Host ID) pairs already exist BEFORE clash zone creation
                // ⚠️ PROTECTION: This logic prevents duplicate clash zone creation with new GUIDs
                // DO NOT MODIFY THIS LOGIC WITHOUT USER CONSENT - Tested and working
                // 
                // LOGIC:
                // 1. If 3-point validation is NOT enabled AND pair exists → skip clash zone creation
                // 2. DO NOT touch flags or parameters of existing clash zone (preserve state)
                // 3. Existing clash zones are already checked by FlagManager above (line ~1021) for sleeve existence
                //    - FlagManager checks Global XML vs Revit API for sleeve existence
                //    - FlagManager follows flag hierarchy (cluster first, then individual)
                //    - If sleeve exists → flags remain true
                //    - If sleeve doesn't exist → flags reset to false
                // 4. Only NEW pairs proceed to clash zone creation
                var existingClashZonesList = existingClashZones?.ClashZones?.ToList() ?? new List<Models.ClashZone>();
                // ✅ REUSE: Settings and enableThreePointValidation already loaded at method level (line ~464)

                // ✅ CRITICAL FIX: First, update intersection points for existing zones with zero values from detected intersections
                // This ensures existing zones have valid intersection points before pair matching
                if (existingClashZonesList.Count > 0 && filteredIntersections.Count > 0)
                {
                    int updatedCount = 0;
                    foreach (var existingZone in existingClashZonesList)
                    {
                        // Check if zone has zero intersection points
                        bool hasZeroPoints = Math.Abs(existingZone.IntersectionPointX) < 1e-9 &&
                                           Math.Abs(existingZone.IntersectionPointY) < 1e-9 &&
                                           Math.Abs(existingZone.IntersectionPointZ) < 1e-9;

                        if (hasZeroPoints)
                        {
                            int mepId = existingZone.MepElementId?.IntegerValue ?? existingZone.MepElementIdValue;
                            int hostId = existingZone.StructuralElementId?.IntegerValue ?? existingZone.StructuralElementIdValue;

                            // Find matching intersection from detected intersections
                            var matchingIntersection = filteredIntersections.FirstOrDefault(i =>
                            {
                                int iMepId = i.Item1.Id.IntegerValue;
                                int iHostId = i.Item2.Id.IntegerValue;
                                return iMepId == mepId && iHostId == hostId;
                            });

                            if (matchingIntersection.Item1 != null)
                            {
                                // Update intersection point from detected intersection
                                var intersectionPoint = matchingIntersection.Item4;
                                existingZone.IntersectionPointX = intersectionPoint.X;
                                existingZone.IntersectionPointY = intersectionPoint.Y;
                                existingZone.IntersectionPointZ = intersectionPoint.Z;
                                existingZone.IntersectionPoint = new XYZ(intersectionPoint.X, intersectionPoint.Y, intersectionPoint.Z);
                                updatedCount++;
                            }
                        }
                    }

                    if (updatedCount > 0)
                    {
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] FIXED: Updated intersection points for {updatedCount} existing zones with zero values from detected intersections\n");
                    }
                }

                // ✅ CRITICAL FIX: Always use intersection point matching (regardless of 3-point validation setting)
                // This prevents duplicate clash zone creation and wrong matches when same MEP crosses multiple hosts
                // Uses knownIntersectionsMap which already has efficient O(1) Dictionary lookup with point key
                // 3-point validation fast-path is cheap (element existence + bounding box check), so we can always do pair matching
                if (existingClashZonesList.Count > 0 && knownIntersectionsMap.Count > 0)
                {
                    int skippedPairExistsCount = 0;
                    var filteredIntersectionsBeforeCreation = new List<(Element, Element, BoundingBoxXYZ, XYZ)>();

                    foreach (var intersection in filteredIntersections)
                    {
                        var (mepElement, structuralElement, bbox, point) = intersection;
                        int mepId = mepElement.Id.IntegerValue;
                        int hostId = structuralElement.Id.IntegerValue;

                        // ✅ USE INTERSECTION POINT: Match by (MEP+Host+Point) using knownIntersectionsMap
                        // This is faster (Dictionary lookup) and more accurate than ClashZonePairMatcher
                        // 3-point validation fast-path is cheap (element existence + bounding box check), so we can always do pair matching
                        double tolerance = 0.1;
                        string pointKey = $"{Math.Round(point.X / tolerance) * tolerance:F1}," +
                                         $"{Math.Round(point.Y / tolerance) * tolerance:F1}," +
                                         $"{Math.Round(point.Z / tolerance) * tolerance:F1}";
                        var key = (mepId, hostId, pointKey);

                        if (knownIntersectionsMap.ContainsKey(key))
                        {
                            // ✅ PROTECTED LOGIC: Pair with matching intersection point exists - skip clash zone creation
                            // DO NOT touch flags or parameters - preserve existing clash zone state
                            // FlagManager already checked this clash zone above (line ~1021) for sleeve existence
                            skippedPairExistsCount++;
                            continue; // Skip this intersection - pair with matching point already exists
                        }

                        filteredIntersectionsBeforeCreation.Add(intersection); // NEW pair - keep for clash zone creation
                    }

                    if (skippedPairExistsCount > 0 && !DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[PAIR-MATCHER] Filtered {skippedPairExistsCount} intersections - pairs with matching intersection points already exist. Existing clash zones were already checked by FlagManager for sleeve existence.");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [PAIR-MATCHER] Filtered {skippedPairExistsCount} intersections - pairs with matching points exist\n");
                    }

                    // Update filteredIntersections to only include NEW pairs
                    filteredIntersections = filteredIntersectionsBeforeCreation;
                }

                // ✅ CRITICAL FIX: Wrap DetectNewClashZones in try-catch to handle any exceptions
                List<Models.ClashZone> newClashZones = null;
                try
                {
                    // ⚠️ CRITICAL: Check memory/timeout before heavy processing
                    _memoryManager.CheckLimits();

                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] About to call DetectNewClashZones...");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] About to call DetectNewClashZones...\n");

                    // ⚠️ CRITICAL: Pass memory manager to clash zone service for timeout/memory protection
                    _clashZoneService.SetMemoryManager(_memoryManager);

                    // ✅ MEMORY PROFILING: Pass profiler to track memory during clash zone detection
                    _memoryProfiler?.TakeSnapshot("BEFORE_DETECT_NEW_CLASH_ZONES", 0);
                    _clashZoneService.SetMemoryProfiler(_memoryProfiler);

                    newClashZones = _clashZoneService.DetectNewClashZones(filteredIntersections, _document, clearanceSettings, selectedMepCategories);

                    // ✅ MEMORY PROFILING: Record snapshot after clash zone detection
                    _memoryProfiler?.TakeSnapshot("AFTER_DETECT_NEW_CLASH_ZONES", newClashZones?.Count ?? 0);

                    // ✅ MEMORY OPTIMIZATION: Flush batched logs after clash zone detection completes
                    if (batchedLogger != null)
                    {
                        batchedLogger.Flush();
                    }

                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] ✅ DetectNewClashZones RETURNED: {newClashZones?.Count ?? 0} clash zones");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ✅ DetectNewClashZones RETURNED: {newClashZones?.Count ?? 0} clash zones\n");

                    // ✅ DEBUG: Log file combos from detected clash zones to diagnose missing combos in global XML
                    if (newClashZones != null && newClashZones.Count > 0)
                    {
                        var combosFromZones = newClashZones
                            .GroupBy(cz => new { 
                                LinkedFile = cz.SourceDocKey ?? cz.DocumentPath ?? "unknown",
                                HostFile = cz.HostDocKey ?? cz.StructuralElementDocumentTitle ?? "unknown",
                                HostType = cz.StructuralElementType ?? "unknown"
                            })
                            .ToList();

                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[CLASH-DEBUG] Detected clash zones grouped by file combo: {combosFromZones.Count} combos");
                            foreach (var combo in combosFromZones.Take(10))
                            {
                                DebugLogger.Info($"[CLASH-DEBUG]   Combo: Linked='{combo.Key.LinkedFile}', Host='{combo.Key.HostFile}', HostType='{combo.Key.HostType}', Zones={combo.Count()}");
                            }
                        }
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH-DEBUG] Detected clash zones grouped by file combo: {combosFromZones.Count} combos\n");
                        foreach (var combo in combosFromZones.Take(10))
                        {
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH-DEBUG]   Combo: Linked='{combo.Key.LinkedFile}', Host='{combo.Key.HostFile}', HostType='{combo.Key.HostType}', Zones={combo.Count()}\n");
                        }
                    }
                    // Skip zones flagged resolved in per-category globals
                    if (__globalsResolved != null && __globalsResolved.Count > 0 && newClashZones != null)
                    {
                        int before = newClashZones.Count;
                        newClashZones = newClashZones.Where(cz => !__globalsResolved.Contains(cz.Id)).ToList();
                        int filtered = before - newClashZones.Count;
                        if (filtered > 0)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[CLASH_DEBUG] Skipped {filtered} zones due to global resolved flags");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Skipped {filtered} zones due to global resolved flags\n");
                        }
                    }
                    __logPhaseMem("AFTER_DETECT", newClashZones, existingClashZones?.ClashZones, null);
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] DetectNewClashZones completed successfully\n");
                }
                catch (TimeoutException timeoutEx)
                {
                    SafeFileLogger.SafeAppendText("refresh_timeouts.log", $"Timeout during DetectNewClashZones: {timeoutEx.Message}");
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[CLASH_DEBUG] ⏱ TIMEOUT in DetectNewClashZones: {timeoutEx.Message}");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ⏱ TIMEOUT in DetectNewClashZones: {timeoutEx.Message}\n");

                    _statusLabel.Text = "Operation timed out during clash zone detection. Try with smaller file.";
                    _progressBar.Visible = false;
                    _refreshButton.Enabled = true;

                    System.Windows.Forms.MessageBox.Show(
                        $"Operation timed out after 5 minutes during clash zone detection.\n\n" +
                        $"Found {currentIntersections.Count} intersections but could not process all.\n\n" +
                        $"Please:\n• Try with a smaller file or section box\n• Process fewer categories at once\n• Contact support if issue persists",
                        "Operation Timeout",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);

                    // Cleanup and return gracefully
                    _memoryManager?.ForceCleanup();
                    return Autodesk.Revit.UI.Result.Cancelled;
                }
                catch (OutOfMemoryException memEx)
                {
                    SafeFileLogger.SafeAppendText("refresh_memory.log", $"Memory limit during DetectNewClashZones: {memEx.Message}");
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[CLASH_DEBUG] 💾 MEMORY LIMIT in DetectNewClashZones: {memEx.Message}");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] 💾 MEMORY LIMIT in DetectNewClashZones: {memEx.Message}\n");

                    _statusLabel.Text = "Memory limit exceeded. Try with smaller file.";
                    _progressBar.Visible = false;
                    _refreshButton.Enabled = true;

                    System.Windows.Forms.MessageBox.Show(
                        $"Memory limit exceeded during clash zone detection.\n\n" +
                        $"Found {currentIntersections.Count} intersections but memory is insufficient.\n\n" +
                        $"Please:\n• Close other applications\n• Try with a smaller file\n• Process fewer categories\n• Contact support if issue persists",
                        "Memory Limit Exceeded",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);

                    // Cleanup and return gracefully
                    _memoryManager?.ForceCleanup();
                    return Autodesk.Revit.UI.Result.Cancelled;
                }
                catch (Exception detectEx)
                {
                    SafeFileLogger.SafeAppendText("refresh_errors.log", $"Error in DetectNewClashZones: {detectEx.Message}\n{detectEx.StackTrace}");
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[CLASH_DEBUG] ❌ ERROR in DetectNewClashZones: {detectEx.Message}");
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[CLASH_DEBUG] Stack trace: {detectEx.StackTrace}");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ❌ ERROR in DetectNewClashZones: {detectEx.Message}\n");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Stack trace: {detectEx.StackTrace}\n");

                    // Show user-friendly error instead of crashing
                    _statusLabel.Text = $"Error during clash detection: {detectEx.Message}";
                    _progressBar.Visible = false;
                    _refreshButton.Enabled = true;

                    System.Windows.Forms.MessageBox.Show(
                        $"An error occurred during clash zone detection.\n\n" +
                        $"Error: {detectEx.Message}\n\n" +
                        $"Please check the log file for details or contact support.",
                        "Detection Error",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Error);

                    newClashZones = new List<Models.ClashZone>(); // Set to empty list to continue

                    // Graceful exit instead of continuing with empty results
                    return Autodesk.Revit.UI.Result.Failed;
                }

                // ✅ CRITICAL DEBUG: Log immediately after DetectNewClashZones returns (or exception)
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] After DetectNewClashZones: {newClashZones?.Count ?? 0} clash zones");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] After DetectNewClashZones: {newClashZones?.Count ?? 0} clash zones\n");

                // ✅ DEBUG: Log Duct-Wall clash zones AFTER optimization
                if (newClashZones != null && newClashZones.Count > 0)
                {
                    var ductWallClashZones = newClashZones.Where(cz =>
                    {
                        var mepCat = cz.MepElementCategory ?? "";
                        var structType = cz.StructuralElementType ?? "";
                        return (mepCat == "Ducts" || mepCat == "Duct Curves" || mepCat == "Duct Accessories")
                            && (structType == "Wall");
                    }).Count();
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] ═══ AFTER OPTIMIZATION: {ductWallClashZones} Duct-Wall clash zones created (from {ductWallIntersections.Count} intersections) ═══");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ═══ AFTER OPTIMIZATION: {ductWallClashZones} Duct-Wall clash zones created (from {ductWallIntersections.Count} intersections) ═══\n");
                }

                // === Parameter Snapshot (whitelist-based) ===
                try
                {
                    // ✅ MEMORY PROFILING: Track memory before parameter snapshots
                    _memoryProfiler?.TakeSnapshot("BEFORE_PARAMETER_SNAPSHOTS", newClashZones?.Count ?? 0);

                    var snapshotService = new ParameterSnapshotService();

                    // Build whitelist from existing storage + curated keys
                    var tempStorage = existingClashZones ?? new Models.ClashZoneStorage();
                    var whitelist = snapshotService.BuildWhitelist(tempStorage, currentIntersections.Select(t => (t.Item1, t.Item2)));

                    // Cache per element to avoid repeat lookups
                    var cache = new Dictionary<(string docKey, int id), List<Models.SerializableKeyValue>>();

                    foreach (var cz in newClashZones ?? new List<Models.ClashZone>())
                    {
                        // Resolve elements from document or links
                        var mep = ElementRetrievalService.GetElementFromDocumentOrLinked(_document, cz.MepElementId, enableLogging: false);
                        var host = ElementRetrievalService.GetElementFromDocumentOrLinked(_document, cz.StructuralElementId, enableLogging: false);
                        if (mep == null || host == null) continue;

                        var mepKey = (snapshotService.GetDocKey(mep), cz.MepElementId?.IntegerValue ?? -1);
                        var hostKey = (snapshotService.GetDocKey(host), cz.StructuralElementId?.IntegerValue ?? -1);

                        if (!cache.TryGetValue(mepKey, out var mepBag))
                        {
                            mepBag = snapshotService.CaptureParams(mep, whitelist);
                            cache[mepKey] = mepBag;
                        }
                        if (!cache.TryGetValue(hostKey, out var hostBag))
                        {
                            hostBag = snapshotService.CaptureParams(host, whitelist);
                            cache[hostKey] = hostBag;
                        }

                        cz.SourceDocKey = mepKey.Item1;
                        cz.HostDocKey = hostKey.Item1;
                        cz.MepParameterValues = mepBag;
                        cz.HostParameterValues = hostBag;

                        // ✅ MEMORY OPTIMIZATION: Clear Revit API objects after coordinates are extracted
                        // This releases ~0.3 KB per clash zone (6 XYZ objects + 1 BoundingBoxXYZ)
                        cz.ClearRevitApiObjects();

                        // ✅ MEMORY TRACKING: Log parameter counts for debugging (always log for memory analysis)
                        // Note: Even in deployment mode, we want memory debug logs to analyze memory usage
                        if (newClashZones != null && newClashZones.Count % 100 == 0)
                        {
                            int totalParams = (mepBag?.Count ?? 0) + (hostBag?.Count ?? 0);
                            int paramBytes = totalParams * 150; // Rough estimate: 150 bytes per parameter
                            SafeFileLogger.SafeAppendText(refreshLogName,
                                $"[{DateTime.Now}] [MEMORY_DEBUG] ClashZone {cz.Id}: MEP params={mepBag?.Count ?? 0}, Host params={hostBag?.Count ?? 0}, Total={totalParams}, Est. size={paramBytes} bytes\n");
                        }
                    }

                    // ✅ FIX 4: Optimize string storage with interning - share strings across ALL clash zones
                    // Combine new and existing clash zones for interning
                    var allClashZones = (newClashZones ?? Enumerable.Empty<Models.ClashZone>()).Concat(existingClashZones?.ClashZones ?? Enumerable.Empty<Models.ClashZone>()).ToList();
                    foreach (var cz in allClashZones)
                    {
                        // Intern parameter keys and values
                        if (cz.MepParameterValues != null)
                        {
                            foreach (var kv in cz.MepParameterValues)
                            {
                                kv.Key = string.Intern(kv.Key ?? string.Empty);
                                kv.Value = string.Intern(kv.Value ?? string.Empty);
                            }
                        }

                        if (cz.HostParameterValues != null)
                        {
                            foreach (var kv in cz.HostParameterValues)
                            {
                                kv.Key = string.Intern(kv.Key ?? string.Empty);
                                kv.Value = string.Intern(kv.Value ?? string.Empty);
                            }
                        }

                        // Intern other repeated string properties
                        if (cz.MepElementCategory != null)
                            cz.MepElementCategory = string.Intern(cz.MepElementCategory);
                        if (cz.StructuralElementType != null)
                            cz.StructuralElementType = string.Intern(cz.StructuralElementType);
                        if (cz.DocumentPath != null)
                            cz.DocumentPath = string.Intern(cz.DocumentPath);
                        if (cz.SourceDocKey != null)
                            cz.SourceDocKey = string.Intern(cz.SourceDocKey);
                        if (cz.HostDocKey != null)
                            cz.HostDocKey = string.Intern(cz.HostDocKey);
                    }

                    // ✅ FIX 5: Add memory profiling to identify culprits
                    // ✅ FIX: Always run memory analysis when deployment mode is OFF OR diagnostic mode is ON
                    if (newClashZones?.Count > 0 && (!DeploymentConfiguration.DeploymentMode || OptimizationFlags.UseDiagnosticMode))
                    {
                        AnalyzeClashZoneMemory(newClashZones, refreshLogName);
                    }

                    // ✅ MEMORY PROFILING: Track memory after parameter snapshots
                    _memoryProfiler?.TakeSnapshot("AFTER_PARAMETER_SNAPSHOTS", newClashZones?.Count ?? 0);
                    __logPhaseMem("AFTER_PARAM_SNAPSHOTS", newClashZones, existingClashZones?.ClashZones, null);

                    // ✅ MEMORY DEBUG: Log average parameter counts across all clash zones
                    // Note: Always log memory debug info even in deployment mode for analysis
                    if (newClashZones?.Count > 0)
                    {
                        int totalMepParams = newClashZones.Sum(cz => cz.MepParameterValues?.Count ?? 0);
                        int totalHostParams = newClashZones.Sum(cz => cz.HostParameterValues?.Count ?? 0);
                        double avgMepParams = (double)totalMepParams / newClashZones.Count;
                        double avgHostParams = (double)totalHostParams / newClashZones.Count;
                        SafeFileLogger.SafeAppendText(refreshLogName,
                            $"[{DateTime.Now}] [MEMORY_DEBUG] Average parameters per clash zone: MEP={avgMepParams:F1}, Host={avgHostParams:F1}, Total={avgMepParams + avgHostParams:F1}\n");
                    }

                    // ✅ MEMORY OPTIMIZATION: Force FULL GC to release temporary objects from parameter capture
                    GC.Collect(2, GCCollectionMode.Forced, true);
                    GC.WaitForPendingFinalizers();
                    GC.Collect(2, GCCollectionMode.Forced, true);
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[PARAM_SNAPSHOT] Non-fatal: {ex.Message}");
                }

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] DetectNewClashZones completed - {newClashZones?.Count ?? 0} new clash zones detected");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] DetectNewClashZones completed - {newClashZones?.Count ?? 0} new clash zones detected\n");

                // ✅ MEMORY DEBUG: Analyze existing clash zones' parameter counts (for memory analysis even when no new zones)
                if (existingClashZones?.ClashZones != null && existingClashZones.ClashZones.Count > 0)
                {
                    int totalExistingMepParams = existingClashZones.ClashZones.Sum(cz => cz.MepParameterValues?.Count ?? 0);
                    int totalExistingHostParams = existingClashZones.ClashZones.Sum(cz => cz.HostParameterValues?.Count ?? 0);
                    double avgExistingMepParams = (double)totalExistingMepParams / existingClashZones.ClashZones.Count;
                    double avgExistingHostParams = (double)totalExistingHostParams / existingClashZones.ClashZones.Count;

                    // Calculate total string size from all properties
                    int totalStringChars = existingClashZones.ClashZones.Sum(cz =>
                        (cz.MepElementUniqueId?.Length ?? 0) +
                        (cz.MepElementCategory?.Length ?? 0) +
                        (cz.DuctShape?.Length ?? 0) +
                        (cz.InsulationType?.Length ?? 0) +
                        (cz.MepElementFormattedSize?.Length ?? 0) +
                        (cz.MepElementSystemAbbreviation?.Length ?? 0) +
                        (cz.DamperConnectorSide?.Length ?? 0) +
                        (cz.DocumentPath?.Length ?? 0) +
                        (cz.StructuralElementDocumentTitle?.Length ?? 0) +
                        (cz.StructuralElementType?.Length ?? 0) +
                        (cz.HostOrientation?.Length ?? 0) +
                        (cz.WallDirectionType?.Length ?? 0) +
                        (cz.MepElementOrientationDirection?.Length ?? 0) +
                        (cz.PipeOpeningType?.Length ?? 0) +
                        (cz.MepElementLevelName?.Length ?? 0) +
                        (cz.SleeveFamilyName?.Length ?? 0) +
                        (cz.SourceDocKey?.Length ?? 0) +
                        (cz.HostDocKey?.Length ?? 0) +
                        (cz.MepElementGeometryHash?.Length ?? 0) +
                        (cz.StructuralElementGeometryHash?.Length ?? 0));
                    int stringSizeBytes = totalStringChars * 2; // UTF-16 = 2 bytes per char

                    // Sample a few clash zones to see parameter details
                    var sampleZones = existingClashZones.ClashZones.Take(5).ToList();
                    SafeFileLogger.SafeAppendText(refreshLogName,
                        $"[{DateTime.Now}] [MEMORY_DEBUG] ===== EXISTING CLASH ZONES ANALYSIS (from XML) =====\n");
                    SafeFileLogger.SafeAppendText(refreshLogName,
                        $"[{DateTime.Now}] [MEMORY_DEBUG] Total existing clash zones: {existingClashZones.ClashZones.Count}\n");
                    SafeFileLogger.SafeAppendText(refreshLogName,
                        $"[{DateTime.Now}] [MEMORY_DEBUG] Average parameters per EXISTING clash zone: MEP={avgExistingMepParams:F1}, Host={avgExistingHostParams:F1}, Total={avgExistingMepParams + avgExistingHostParams:F1}\n");
                    SafeFileLogger.SafeAppendText(refreshLogName,
                        $"[{DateTime.Now}] [MEMORY_DEBUG] Total string data: {totalStringChars} chars = {stringSizeBytes / 1024.0:F1} KB (avg {stringSizeBytes / existingClashZones.ClashZones.Count / 1024.0:F2} KB per clash zone)\n");

                    foreach (var cz in sampleZones)
                    {
                        int totalParams = (cz.MepParameterValues?.Count ?? 0) + (cz.HostParameterValues?.Count ?? 0);
                        int paramBytes = totalParams * 150; // Rough estimate
                        int stringBytes = ((cz.MepElementUniqueId?.Length ?? 0) + (cz.MepElementCategory?.Length ?? 0) +
                            (cz.MepElementFormattedSize?.Length ?? 0) + (cz.SourceDocKey?.Length ?? 0) +
                            (cz.HostDocKey?.Length ?? 0)) * 2; // UTF-16
                        SafeFileLogger.SafeAppendText(refreshLogName,
                            $"[{DateTime.Now}] [MEMORY_DEBUG] Sample ClashZone {cz.Id}: MEP params={cz.MepParameterValues?.Count ?? 0}, Host params={cz.HostParameterValues?.Count ?? 0}, Total={totalParams}, Param bytes={paramBytes}, String bytes={stringBytes}, Total est={paramBytes + stringBytes} bytes\n");
                    }

                    SafeFileLogger.SafeAppendText(refreshLogName,
                        $"[{DateTime.Now}] [MEMORY_DEBUG] ===========================================\n");
                }

                // Enforce UI host-type filter on new clashes
                try
                {
                    var allowedHostTypesUI_NewZones = new HashSet<string>(
                        FilterUiStateProvider.GetSelectedHostCategories?.Invoke() ?? new List<string>(),
                        StringComparer.OrdinalIgnoreCase);
                    if (allowedHostTypesUI_NewZones.Count > 0 && newClashZones != null)
                    {
                        foreach (var cz in newClashZones)
                        {
                            // Handle plural/singular mismatch: "Walls" (UI) vs "Wall" (Revit)
                            bool isEligible = allowedHostTypesUI_NewZones.Contains(cz.StructuralElementType) ||
                                             allowedHostTypesUI_NewZones.Contains(cz.StructuralElementType + "s") ||
                                             allowedHostTypesUI_NewZones.Any(t => t.TrimEnd('s').Equals(cz.StructuralElementType, StringComparison.OrdinalIgnoreCase));
                            cz.IsEligibleByCurrentUi = isEligible;
                        }
                        var eligible = newClashZones.Count(cz => cz.IsEligibleByCurrentUi);
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CLASH_DEBUG] UI host filter marked NEW zones eligible: {eligible}/{newClashZones.Count}");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] UI host filter marked NEW zones eligible: {eligible}/{newClashZones.Count}\n");
                    }
                }
                catch (Exception uiFilterEx)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[CLASH_DEBUG] UI host filter failed: {uiFilterEx.Message}");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] UI host filter failed: {uiFilterEx.Message}\n");
                }

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] About to start Step 8: Save results");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] About to start Step 8: Save results\n");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] newClashZones count: {newClashZones?.Count ?? 0}\n");

                // Step 8: Save results
                _progressBar.Value = 70;
                _statusLabel.Text = "Saving results...";

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] Step 8 started - Progress bar set to 70");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Step 8 started - Progress bar set to 70\n");

                // Save to enabled filter
                var enabledFilter = filtersToProcess.FirstOrDefault(f => f.IsEnabled);
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Looking for enabled filter... filtersToProcess.Count={filtersToProcess.Count}\n");
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] Found {filtersToProcess.Count} filters to process");
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] Enabled filter: {(enabledFilter != null ? enabledFilter.Name : "NULL")}");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Found {filtersToProcess.Count} filters to process\n");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Enabled filter: {(enabledFilter != null ? enabledFilter.Name : "NULL")}\n");
                if (enabledFilter != null)
                {
                    if (!detectionRan)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info("[CLASH_DEBUG] Replay-only refresh: skipping persistence and leaving Filter XML unchanged.");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Replay path detected – Filter/Global XML not modified\n");
                    }
                    else
                    {
                        var targetFilter = enabledFilter;
                        // ✅ FIX: Save ALL clash zones (existing + new), not just new ones
                        // This preserves flags (IsResolved, IsClusterResolved) from previous runs
                        var allClashZones = existingClashZones?.ClashZones ?? new List<Models.ClashZone>(); // Use existing loaded clash zones
                                                                                                            // Merge existing + new zones (avoid duplicates by Id) BEFORE saving, so snapshot bags persist to XML
                        var existingById = allClashZones.ToDictionary(z => z.Id, z => z);

                        foreach (var nz in newClashZones ?? new List<Models.ClashZone>())
                        {
                            if (!existingById.ContainsKey(nz.Id))
                            {
                                existingById[nz.Id] = nz;
                                allClashZones.Add(nz); // Only one list!
                            }
                            else
                            {
                                var ez = existingById[nz.Id];
                                ez.SourceDocKey = nz.SourceDocKey;
                                ez.HostDocKey = nz.HostDocKey;
                                ez.MepParameterValues = nz.MepParameterValues;
                                ez.HostParameterValues = nz.HostParameterValues;
                                if (nz.StructuralElementThickness > 0)
                                    ez.StructuralElementThickness = nz.StructuralElementThickness;
                                if (nz.StructuralElementNormal != null)
                                    ez.StructuralElementNormal = nz.StructuralElementNormal;
                            }
                        }

                        // ---------------------------------------------------------------------
                        // 1. ENSURE ONLY ONE MASTER CLASH ZONE LIST
                        // ---------------------------------------------------------------------
                        if (existingClashZones != null && !object.ReferenceEquals(existingClashZones.ClashZones, allClashZones))
                            existingClashZones.ClashZones = allClashZones;

                        // ✅ SAFEGUARD: Ensure we only persist clash zones for categories selected in the UI
                        if (selectedMepCategories != null && selectedMepCategories.Count > 0)
                        {
                            var allowedCategories = new HashSet<string>(selectedMepCategories, StringComparer.OrdinalIgnoreCase);
                            int removedCount = allClashZones.RemoveAll(cz => cz == null || !allowedCategories.Contains(cz.MepElementCategory ?? string.Empty));

                            if (removedCount > 0)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[REFRESH] Filtered out {removedCount} clash zones not in UI-selected categories: [{string.Join(", ", allowedCategories)}]");

                                SafeFileLogger.SafeAppendText(refreshLogName,
                                    $"[{DateTime.Now}] [REFRESH] Filtered out {removedCount} clash zones not matching selected categories\n");
                            }

                            // ✅ CRITICAL: Persist the current UI selection onto the filter so category-specific saves honor it
                            try
                            {
                                targetFilter.SelectedMepCategoryNames = allowedCategories.ToList();
                            }
                            catch
                            {
                                targetFilter.SelectedMepCategoryNames = selectedMepCategories.ToList();
                            }
                        }

                        // ✅ CRITICAL FIX: Use dedicated ClashZonePersistenceService to save to both Global XML and Filter XML
                        // This ensures data consistency by using the same clash zone objects for both
                        // Called once for all categories - service internally groups by category and saves to correct files
                        // Called AFTER allClashZones is ready (line 3090) but BEFORE Filter XML file save (line 3689)
                        Models.FilterFileComboGroup CloneFilterFileCombo(Models.FilterFileComboGroup combo)
                        {
                            if (combo == null)
                                return new Models.FilterFileComboGroup
                                {
                                    LinkedFile = string.Empty,
                                    HostFile = string.Empty,
                                    ProcessedAt = DateTime.Now,
                                    ClashZones = new List<Models.ClashZone>()
                                };

                            return new Models.FilterFileComboGroup
                            {
                                LinkedFile = combo.LinkedFile,
                                HostFile = combo.HostFile,
                                ProcessedAt = combo.ProcessedAt,
                                ClashZones = combo.ClashZones?
                                    .Where(z => z != null)
                                    .ToList() ?? new List<Models.ClashZone>()
                            };
                        }

                        Models.FilterGroupForStorage CloneFilterGroup(Models.FilterGroupForStorage group)
                        {
                            if (group == null)
                                return new Models.FilterGroupForStorage
                                {
                                    Name = string.Empty,
                                    FileCombos = new List<Models.FilterFileComboGroup>()
                                };

                            return new Models.FilterGroupForStorage
                            {
                                Name = group.Name,
                                FileCombos = group.FileCombos?.Select(CloneFilterFileCombo).ToList() ?? new List<Models.FilterFileComboGroup>()
                            };
                        }

                        try
                    {
                        var baseFilterName = targetFilter?.Name ?? enabledFilter?.Name ?? filtersToProcess.FirstOrDefault(f => f.IsEnabled)?.Name ?? string.Empty;
                        var persistenceService = new ClashZonePersistenceService(_document, _guidManager, refreshLogName);
                        persistenceService.SaveClashZones(allClashZones, baseFilterName, targetFilter, allowStructuralUpdates: true);
                        SyncFilterStorageFromZones(targetFilter, allClashZones, baseFilterName);

                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[REFRESH] ✅ Saved {allClashZones.Count} clash zones using ClashZonePersistenceService");

                        // ✅ CRITICAL FIX: Mark ALL selected file combos as processed in Global XML, even if they have no clash zones
                        // This ensures that file combos selected in UI are always added to global XML, preventing them from being re-detected
                        // Example: User selects AR001 + WALL but no intersections found → Still mark combo as processed so it appears in global XML
                        if (selectedMepCategories != null && selectedMepCategories.Count > 0 &&
                            selectedReferenceFiles != null && selectedReferenceFiles.Count > 0 &&
                            selectedHostFiles != null && selectedHostFiles.Count > 0)
                        {
                            try
                            {
                                // Build all file combos from UI selections
                                var allSelectedFileCombos = new List<(string LinkedFile, string HostFile)>();
                                foreach (var refFile in selectedReferenceFiles)
                                {
                                    foreach (var hostFile in selectedHostFiles)
                                    {
                                        if (!string.IsNullOrWhiteSpace(refFile) && !string.IsNullOrWhiteSpace(hostFile))
                                        {
                                            allSelectedFileCombos.Add((refFile, hostFile));
                                        }
                                    }
                                }

                                if (allSelectedFileCombos.Count > 0)
                                {
                                    // ✅ DEBUG: Log UI selections before marking as processed
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        DebugLogger.Info($"[GLOBAL-INDEX] UI Selected file combos to mark as processed: {allSelectedFileCombos.Count} combos");
                                        foreach (var combo in allSelectedFileCombos.Take(10))
                                        {
                                            DebugLogger.Info($"[GLOBAL-INDEX]   UI Combo: Linked='{combo.LinkedFile}', Host='{combo.HostFile}'");
                                        }
                                    }
                                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] UI Selected file combos: {allSelectedFileCombos.Count}\n");
                                    foreach (var combo in allSelectedFileCombos.Take(10))
                                    {
                                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX]   UI Combo: Linked='{combo.LinkedFile}', Host='{combo.HostFile}'\n");
                                    }

                                    // Mark all file combos as processed for each category
                                    foreach (var category in selectedMepCategories)
                                    {
                                        if (string.IsNullOrWhiteSpace(category))
                                            continue;

                                        try
                                        {
                                            GlobalIndexService.MarkFileCombosAsProcessed(_document, category, allSelectedFileCombos, baseFilterName);
                                            
                                            if (!DeploymentConfiguration.DeploymentMode)
                                            {
                                                DebugLogger.Info($"[GLOBAL-INDEX] ✅ Marked {allSelectedFileCombos.Count} file combos as processed for category '{category}' (including combos with zero clash zones)");
                                            }
                                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] ✅ Marked {allSelectedFileCombos.Count} file combos as processed for '{category}'\n");
                                        }
                                        catch (Exception markEx)
                                        {
                                            if (!DeploymentConfiguration.DeploymentMode)
                                            {
                                                DebugLogger.Warning($"[GLOBAL-INDEX] Error marking file combos as processed for category '{category}': {markEx.Message}");
                                            }
                                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] ERROR marking combos for '{category}': {markEx.Message}\n");
                                        }
                                    }
                                }
                            }
                            catch (Exception comboMarkEx)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Warning($"[GLOBAL-INDEX] Error marking file combos as processed: {comboMarkEx.Message}");
                                }
                                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] ERROR marking file combos: {comboMarkEx.Message}\n");
                            }
                        }

                        var persistedFilterGroupCount = targetFilter?.ClashZoneStorage?.Filters?.Count ?? 0;
                        var persistedFileComboCount = targetFilter?.ClashZoneStorage?.Filters?
                            .Where(g => g != null && g.FileCombos != null)
                            .SelectMany(g => g.FileCombos)
                            .Count() ?? 0;
                        var persistedZoneCount = targetFilter?.ClashZoneStorage?.EnumerateAllZones()?.Count() ?? 0;

                        SafeFileLogger.SafeAppendText(refreshLogName,
                            $"[{DateTime.Now}] [REFRESH-PERSIST] After SaveClashZones → FilterGroups={persistedFilterGroupCount}, FileCombos={persistedFileComboCount}, Zones={persistedZoneCount}\n");

                        // ✅ PHASE SQLITE-2: Skip Filter XML persistence when SQLite is primary (SQLite has the data)
                        if (!DeploymentConfiguration.UseSqliteAsPrimary)
                        {
                            // Legacy mode: Persist updated filter tree to XML files (main + per-category)
                            try
                            {
                                var filterDir = ProjectPathService.GetFiltersDirectory(_document);
                                if (!Directory.Exists(filterDir))
                                    Directory.CreateDirectory(filterDir);

                                if (targetFilter != null)
                                {
                                    // ✅ PHASE 2: Only save to XML if XML creation is enabled
                                    if (!DeploymentConfiguration.DisableXmlCreation)
                                    {
                                        var mainFilePath = Path.Combine(filterDir, $"{targetFilter.Name}.xml");

                                        SafeFileLogger.SafeAppendText(refreshLogName,
                                            $"[{DateTime.Now}] [REFRESH-PERSIST] Saving main filter '{targetFilter.Name}' (Zones={targetFilter.ClashZoneStorage?.EnumerateAllZones()?.Count() ?? 0})\n");

                                        _filterManagementService.SaveFilterToXmlFile(targetFilter, mainFilePath);
                                    }
                                    else
                                    {
                                        SafeFileLogger.SafeAppendText(refreshLogName,
                                            $"[{DateTime.Now}] [REFRESH-PERSIST] ⚠️ XML creation disabled - skipping filter XML save (database only mode)\n");
                                    }

                                if (targetFilter.ClashZoneStorage?.Filters != null)
                                {
                                    foreach (var filterGroup in targetFilter.ClashZoneStorage.Filters)
                                    {
                                        if (filterGroup == null)
                                        {
                                            SafeFileLogger.SafeAppendText(refreshLogName,
                                                $"[{DateTime.Now}] [REFRESH-PERSIST] Skipping null filter group entry\n");
                                            continue;
                                        }

                                        var comboCount = filterGroup.FileCombos?.Count ?? 0;
                                        SafeFileLogger.SafeAppendText(refreshLogName,
                                            $"[{DateTime.Now}] [REFRESH-PERSIST] Processing group '{filterGroup.Name}' → FileCombos={comboCount}\n");

                                        if (filterGroup.FileCombos == null || filterGroup.FileCombos.Count == 0)
                                        {
                                            SafeFileLogger.SafeAppendText(refreshLogName,
                                                $"[{DateTime.Now}] [REFRESH-PERSIST]   Group '{filterGroup.Name}' has no combos - skipping XML write\n");
                                            continue;
                                        }

                                        foreach (var combo in filterGroup.FileCombos.Where(fc => fc != null))
                                        {
                                            var comboZoneCount = combo.ClashZones?.Count ?? 0;
                                            SafeFileLogger.SafeAppendText(refreshLogName,
                                                $"[{DateTime.Now}] [REFRESH-PERSIST]   Combo Linked='{combo.LinkedFile}' Host='{combo.HostFile}' Zones={comboZoneCount}\n");
                                        }

                                        var groupClone = CloneFilterGroup(filterGroup);
                                        var groupZones = groupClone.FileCombos?
                                            .SelectMany(fc => fc.ClashZones ?? Enumerable.Empty<Models.ClashZone>())
                                            .Where(z => z != null)
                                            .GroupBy(z => z.Id)
                                            .Select(g => g.First())
                                            .ToList() ?? new List<Models.ClashZone>();

                                        // ✅ FIX: Skip groups that have no zones after extraction (even if they have FileCombos)
                                        // This prevents saving empty XML files for groups like 'Electrical' when zones are in 'Electrical_cable_trays'
                                        if (groupZones.Count == 0)
                                        {
                                            SafeFileLogger.SafeAppendText(refreshLogName,
                                                $"[{DateTime.Now}] [REFRESH-PERSIST]   Group '{filterGroup.Name}' has no zones after extraction - skipping XML write\n");
                                            continue;
                                        }

                                        var categoryFilter = new Models.OpeningFilter
                                        {
                                            Name = filterGroup.Name,
                                            Category = targetFilter.Category,
                                            OpeningType = targetFilter.OpeningType,
                                            IsEnabled = targetFilter.IsEnabled,
                                            LastModified = DateTime.Now,
                                            SelectedMepCategoryNames = groupZones
                                                .Select(z => z.MepElementCategory)
                                                .Where(cat => !string.IsNullOrWhiteSpace(cat))
                                                .Distinct(StringComparer.OrdinalIgnoreCase)
                                                .ToList(),
                                            ClashZoneStorage = new Models.ClashZoneStorage
                                            {
                                                Filters = new List<Models.FilterGroupForStorage> { groupClone },
                                                ClashZones = new List<Models.ClashZone>(groupZones)
                                            }
                                        };

                                        var categoryFilePath = Path.Combine(filterDir, $"{filterGroup.Name}.xml");

                                        SafeFileLogger.SafeAppendText(refreshLogName,
                                            $"[{DateTime.Now}] [REFRESH-PERSIST] Saving category filter '{filterGroup.Name}' (Zones={groupZones.Count}) → {categoryFilePath}\n");

                                        _filterManagementService.SaveFilterToXmlFile(categoryFilter, categoryFilePath);
                                    }
                                }
                            }
                            }
                            catch (Exception xmlPersistEx)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Warning($"[REFRESH] Failed to persist filter XML files after unified save: {xmlPersistEx.Message}");
                                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [REFRESH] WARNING: Failed to persist filter XML files: {xmlPersistEx.Message}\n");
                                }
                            }
                        }
                        else
                        {
                            // ✅ PHASE SQLITE-2: SQLite is primary - skip Filter XML persistence
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[REFRESH] ✅ PHASE 2: Skipped Filter XML persistence - SQLite is primary store");
                            SafeFileLogger.SafeAppendText(refreshLogName,
                                $"[{DateTime.Now}] [REFRESH-PERSIST] ✅ PHASE 2: Skipped Filter XML persistence - SQLite is primary store\n");
                        }
                    }
                    catch (Exception persistenceEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Error($"[REFRESH] Error saving clash zones via ClashZonePersistenceService: {persistenceEx.Message}");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [REFRESH] ERROR saving via ClashZonePersistenceService: {persistenceEx.Message}\n");
                        }
                        // Don't throw - continue with rest of refresh
                    }

                    newClashZones = null;

                    // ✅ CRITICAL FIX: Sync flags from Global XML AFTER merging new clash zones
                    // When a new linked file is added, refresh creates NEW clash zones with IsResolved=false
                    // But if these clash zones have the SAME GUID as previously placed sleeves, flags must be synced
                    // This prevents duplicate sleeves when the same filter is run again with additional linked files
                    try
                    {
                        var clashZonesByCategoryAfterMerge = allClashZones
                            .GroupBy(cz => cz.MepElementCategory)
                            .ToList();

                        int totalSyncedAfterMerge = 0;
                        foreach (var categoryGroup in clashZonesByCategoryAfterMerge)
                        {
                            var category = categoryGroup.Key;
                            var categoryClashZones = categoryGroup.ToList();

                            // Sync flags from Global XML (source of truth)
                            _flagManager.SyncFlagsFromGlobal(categoryClashZones, category);

                            int syncedInCategory = categoryClashZones.Count(cz =>
                                (cz.IsResolved || cz.IsClusterResolved) &&
                                !string.IsNullOrEmpty(cz.MepElementCategory));

                            totalSyncedAfterMerge += syncedInCategory;

                            if (!DeploymentConfiguration.DeploymentMode && syncedInCategory > 0)
                            {
                                DebugLogger.Info($"[REFRESH-GLOBAL-SYNC-AFTER-MERGE] Synced flags for {syncedInCategory} clash zones in category '{category}' AFTER merging new zones");
                                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [REFRESH-GLOBAL-SYNC-AFTER-MERGE] Synced {syncedInCategory} clash zones from Global XML for category '{category}'\n");
                            }
                        }

                        if (totalSyncedAfterMerge > 0 && !DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[REFRESH-GLOBAL-SYNC-AFTER-MERGE] ✅ Total {totalSyncedAfterMerge} clash zones synced from Global XML AFTER merging new zones");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [REFRESH-GLOBAL-SYNC-AFTER-MERGE] ✅ Total {totalSyncedAfterMerge} clash zones synced from Global XML after merge\n");
                        }
                    }
                    catch (Exception syncAfterMergeEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Error($"[REFRESH-GLOBAL-SYNC-AFTER-MERGE] Error syncing flags after merge: {syncAfterMergeEx.Message}");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [REFRESH-GLOBAL-SYNC-AFTER-MERGE] ERROR: {syncAfterMergeEx.Message}\n");
                        }
                    }

                    // ---------------------------------------------------------------------
                    // 2. DEFER CLEARING REVIT API OBJECTS UNTIL AFTER PERSISTENCE
                    // IntersectionPointX/Y/Z and serialized bbox getters depend on non-null objects.
                    // We'll clear after XML files are saved to avoid writing zeros.

                    // ---------------------------------------------------------------------
                    // 3. INTERN STRING FIELDS IN PARAM SNAPSHOTS (saves memory)
                    var internPool = new HashSet<string>(StringComparer.Ordinal);
                    foreach (var cz in allClashZones)
                    {
                        if (cz.MepParameterValues != null)
                            foreach (var kv in cz.MepParameterValues)
                            {
                                if (!internPool.Add(kv.Key)) kv.Key = internPool.First(s => s == kv.Key);
                                if (!internPool.Add(kv.Value)) kv.Value = internPool.First(s => s == kv.Value);
                            }
                        if (cz.HostParameterValues != null)
                            foreach (var kv in cz.HostParameterValues)
                            {
                                if (!internPool.Add(kv.Key)) kv.Key = internPool.First(s => s == kv.Key);
                                if (!internPool.Add(kv.Value)) kv.Value = internPool.First(s => s == kv.Value);
                            }
                    }

                    // ---------------------------------------------------------------------
                    // 4. FORCE GC AND TAKE SNAPSHOTS
                    _memoryProfiler?.TakeSnapshot("BEFORE_GC_MERGE", allClashZones.Count);
                    GC.Collect(2, GCCollectionMode.Forced, true);
                    GC.WaitForPendingFinalizers();
                    GC.Collect(2, GCCollectionMode.Forced, true);
                    _memoryProfiler?.TakeSnapshot("AFTER_GC_MERGE", allClashZones.Count);

                    // Detailed memory diagnostics after GC
                    try
                    {
                        var managedBytes = GC.GetTotalMemory(false);
                        var uniqueIds = new HashSet<Guid>();
                        foreach (var _cz in allClashZones)
                            uniqueIds.Add(_cz.Id);
                        var uniqueCount = uniqueIds.Count;
                        double duplicationFactor = uniqueCount == 0 ? 0.0 : (double)allClashZones.Count / uniqueCount;
                        var managedMb = managedBytes / (1024.0 * 1024.0);
                        var diagLine = $"[MEMORY_DEBUG] PHASE=AFTER_GC_MERGE: managed={managedMb:F1} MB, counts: totalZones={allClashZones.Count}, uniqueZones={uniqueCount}, duplicationFactor={duplicationFactor:F2}" + Environment.NewLine;
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] {diagLine}");
                    }
                    catch { }

                    // ---------------------------------------------------------------------
                    // 5. BATCHED LOGGER (Optional: activate if using string-heavy logging)
                    batchedLogger?.Dispose();
                    batchedLogger = new BatchedLogger(msg => SafeFileLogger.SafeAppendText(refreshLogName, msg), batchSize: 100, flushIntervalSeconds: 1);

                    // ---------------------------------------------------------------------
                    // 6. MEMORY SNAPSHOT & LOG
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[PARAM_SNAPSHOT] Persisting {allClashZones.Count} zones (cleared, deduped, interned)");
                    batchedLogger.Log($"[{DateTime.Now}] [MEMORY_DEBUG] AFTER_MERGE_NO_DUP: {allClashZones.Count} zones, post-GC memory snapshot taken{Environment.NewLine}");
                    // ✅ CRITICAL FIX: ClashZonePersistenceService already saved clash zones to Filter XML
                    // targetFilter.ClashZoneStorage is already populated by the service
                    // Just ensure clashZoneStorage variable is set for compatibility
                    var clashZoneStorage = targetFilter?.ClashZoneStorage ?? new Models.ClashZoneStorage
                    {
                        ClashZones = allClashZones,
                        CreatedAt = DateTime.Now,
                        LastUpdated = DateTime.Now,
                        DocumentPath = _document.PathName,
                        DocumentHash = _document.PathName ?? "Unknown",
                        AlgorithmVersion = "1.0"
                    };

                    targetFilter.LastModified = DateTime.Now;
                    // ✅ CRITICAL FIX: UI state NOT saved to Filter XML (only placement data)
                    // UI state is saved to Global XML ProcessedFileCombos only (see line ~3741)
                    // Also persist host categories selected in UI (this is placement-related data, not UI state)
                    try { targetFilter.SelectedHostCategories = FilterUiStateProvider.GetSelectedHostCategories?.Invoke() ?? new List<string>(); } catch { }
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] ✅ UI state NOT saved to Filter XML (only placement data saved)");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ✅ UI state NOT saved to Filter XML (only placement data)\n");

                    // ✅ FLAG RESET: After processing new files, mark them as "old" files for next refresh
                    // This ensures that after processing intersection detection for new files (e.g., Model B),
                    // they are saved to Filter XML and will be considered "old" files in the next refresh
                    // Next refresh will skip intersection detection for these files and only process truly new files
                    if (hasNewFilesAdded && (newReferenceFiles != null && newReferenceFiles.Count > 0 || newHostFiles != null && newHostFiles.Count > 0))
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[FLAG-RESET] ✅ NEW FILES PROCESSED - Resetting flag to 'OLD' status:");
                            if (newReferenceFiles != null && newReferenceFiles.Count > 0)
                                DebugLogger.Info($"[FLAG-RESET]   Processed NEW reference files: {string.Join(", ", newReferenceFiles)} → Now saved as OLD files in Filter XML");
                            if (newHostFiles != null && newHostFiles.Count > 0)
                                DebugLogger.Info($"[FLAG-RESET]   Processed NEW host files: {string.Join(", ", newHostFiles)} → Now saved as OLD files in Filter XML");
                            DebugLogger.Info($"[FLAG-RESET]   Next refresh will skip intersection detection for these files (treated as already-processed)");
                        }
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [FLAG-RESET] ✅ NEW FILES PROCESSED - Resetting to 'OLD' status (ref: {newReferenceFiles?.Count ?? 0}, host: {newHostFiles?.Count ?? 0})\n");
                    }

                    // Save to profile configuration for persistence
                    if (currentProfile.Configuration == null)
                    {
                        // Create a proper UserConfiguration object
                        currentProfile.Configuration = new Models.UserConfiguration(); // TODO: Initialize with proper values
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info("[CLASH_DEBUG] Created new profile configuration during Refresh");
                    }

                    // ✅ CRITICAL FIX: Set clash zone storage in profile configuration
                    // This ensures the OK button can find clash zones after refresh
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] About to set clash zone storage to profile configuration...\n");
                    currentProfile.Configuration.ClashZoneStorage = clashZoneStorage;
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] Saved {clashZoneStorage.ClashZones.Count} clash zones to profile configuration");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Saved {clashZoneStorage.ClashZones.Count} clash zones to profile configuration\n");

                    // Save the profile to persist changes
                    try
                    {
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] About to call SaveCurrentProfile()...\n");
                        _appProfileService.SaveCurrentProfile();
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info("[CLASH_DEBUG] Profile saved successfully after setting clash zone storage");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Profile saved successfully\n");
                    }
                    catch (Exception saveEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error($"[CLASH_DEBUG] Failed to save profile: {saveEx.Message}");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ERROR saving profile: {saveEx.Message}\n");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Profile save exception: {saveEx.StackTrace}\n");
                    }
                    // TODO: Persist current opening conditions (e.g., clearance values) into profile configuration
                    // if (currentProfile.Configuration.OpeningSettings == null)
                    // {
                    //     currentProfile.Configuration.OpeningSettings = new Models.OpeningSettings();
                    // }
                    
                    // AFTER PERSISTENCE: Now safe to clear heavy Revit API objects so future runs use less memory
                    try
                    {
                        foreach (var cz in allClashZones)
                            cz.ClearRevitApiObjects();
                    }
                    catch { }
                        var savedCount = clashZoneStorage?.ClashZones?.Count ?? 0;
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] SUCCESS: Saved {savedCount} clash zones to filter '{targetFilter.Name}' and profile configuration\n");
                    }
                }
                else
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning("[CLASH_DEBUG] No enabled filter found to save clash zones");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] WARNING: No enabled filter found to save clash zones\n");
                }

                // Step 9: Update UI with results
                _progressBar.Value = 100;

                // ✅ CRITICAL FIX: Calculate final statistics from saved filter data
                int finalTotal = 0;
                int finalUnresolved = 0;
                int finalNewZones = 0;

                if (enabledFilter != null && enabledFilter.ClashZoneStorage != null && enabledFilter.ClashZoneStorage.AllZones != null)
                {
                    finalTotal = enabledFilter.ClashZoneStorage.AllZones.Count;
                    finalUnresolved = enabledFilter.ClashZoneStorage.AllZones.Count(cz => !cz.IsResolved);
                    finalNewZones = newClashZones?.Count ?? 0;
                }

                // Show results with log location
                string logDir = SafeFileLogger.GetLogDirectory();
                _statusLabel.Text = $"Clash zones: {finalTotal} total, {finalUnresolved} unresolved, {finalNewZones} new | Logs: {logDir}";

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] Clash zone refresh complete: {finalTotal} total, {finalUnresolved} unresolved, {finalNewZones} new");

                // ✅ USER FEEDBACK: Prompt user if no new clash zones found AND no zones need processing
                // Only show if: 1) No new zones detected, 2) No zones were invalidated, 3) All zones are resolved (cluster or individual)
                var invalidatedCount = existingClashZones?.ClashZones != null ?
                    (existingCount - (existingClashZones.ClashZones.Count)) : 0;
                // ✅ FLAG MANAGEMENT: Count unresolved zones (neither cluster resolved nor individual resolved)
                // Flags are managed by FlagManager.SyncFlagsFromGlobal() - we just check the flag properties here
                // ✅ CRITICAL FIX: Use finalTotal from saved filter data (not existingClashZones which might be stale)
                // ✅ CRITICAL FIX: After resetting flags for deleted sleeves, use the updated clash zones from allClashZones
                // This ensures we count unresolved zones correctly after flags have been reset
                int unresolvedCount = 0;
                int totalCount = 0;
                int clusterResolvedCount = 0;
                int individualResolvedCount = 0;
                // ✅ CRITICAL FIX: Use the updated clash zones list from enabledFilter (which was just saved with reset flags)
                // After flags are reset, enabledFilter.ClashZoneStorage is updated with the reset flags
                // This ensures we count unresolved zones correctly after flags have been reset for deleted sleeves
                List<Models.ClashZone> clashZonesToCheck = null;
                // ✅ CRITICAL: Use enabledFilter.ClashZoneStorage which contains the updated flags after reset
                // The flags were reset in existingClashZones.ClashZones, then saved to enabledFilter.ClashZoneStorage
                if (enabledFilter != null && enabledFilter.ClashZoneStorage != null && enabledFilter.ClashZoneStorage.AllZones != null)
                {
                    clashZonesToCheck = enabledFilter.ClashZoneStorage.AllZones;

                    // ✅ CRITICAL FIX: Sync flags from Global XML BEFORE counting unresolved zones
                    // This ensures zones with reset flags (IsResolved=false after sleeve deletion) are correctly counted
                    if (_flagManager != null && clashZonesToCheck.Count > 0)
                    {
                        try
                        {
                            // Group by category and sync flags for each category
                            var clashZonesByCategory = clashZonesToCheck
                                .GroupBy(cz => cz.MepElementCategory)
                                .ToList();

                            foreach (var categoryGroup in clashZonesByCategory)
                            {
                                var category = categoryGroup.Key;
                                var categoryClashZones = categoryGroup.ToList();

                                if (!string.IsNullOrWhiteSpace(category))
                                {
                                    _flagManager.SyncFlagsFromGlobal(categoryClashZones, category);

                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[CLASH_DEBUG] Synced flags from Global XML for {categoryClashZones.Count} zones in category '{category}' before unresolved count check");
                                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Synced flags from Global XML for {categoryClashZones.Count} zones in category '{category}' before unresolved count check\n");
                                }
                            }
                        }
                        catch (Exception syncEx)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[CLASH_DEBUG] Failed to sync flags from Global XML before unresolved count check: {syncEx.Message}");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Failed to sync flags from Global XML: {syncEx.Message}\n");
                        }
                    }

                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] Using enabledFilter.ClashZoneStorage for unresolved count check: {clashZonesToCheck.Count} zones");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Using enabledFilter.ClashZoneStorage for unresolved count check: {clashZonesToCheck.Count} zones\n");
                }
                else if (existingClashZones != null && existingClashZones.ClashZones != null && existingClashZones.ClashZones.Count > 0)
                {
                    // Fallback to existingClashZones if enabledFilter is not available (flags were reset here)
                    clashZonesToCheck = existingClashZones.ClashZones;

                    // ✅ CRITICAL FIX: Also sync flags from Global XML for fallback case
                    if (_flagManager != null && clashZonesToCheck.Count > 0)
                    {
                        try
                        {
                            var clashZonesByCategory = clashZonesToCheck
                                .GroupBy(cz => cz.MepElementCategory)
                                .ToList();

                            foreach (var categoryGroup in clashZonesByCategory)
                            {
                                var category = categoryGroup.Key;
                                var categoryClashZones = categoryGroup.ToList();

                                if (!string.IsNullOrWhiteSpace(category))
                                {
                                    _flagManager.SyncFlagsFromGlobal(categoryClashZones, category);
                                }
                            }
                        }
                        catch (Exception syncEx)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[CLASH_DEBUG] Failed to sync flags from Global XML (fallback): {syncEx.Message}");
                        }
                    }

                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] Fallback: Using existingClashZones for unresolved count check: {clashZonesToCheck.Count} zones");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Fallback: Using existingClashZones for unresolved count check: {clashZonesToCheck.Count} zones\n");
                }
                if (clashZonesToCheck != null)
                {
                    totalCount = clashZonesToCheck.Count;
                    unresolvedCount = clashZonesToCheck.Count(cz => cz != null && !cz.IsClusterResolved && !cz.IsResolved);
                    clusterResolvedCount = clashZonesToCheck.Count(cz => cz.IsClusterResolved);
                    individualResolvedCount = clashZonesToCheck.Count(cz => cz.IsResolved && !cz.IsClusterResolved);

                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] Flag check (Filter XML): {totalCount} total, {unresolvedCount} unresolved (after reset and sync), {clusterResolvedCount} cluster resolved, {individualResolvedCount} individual resolved");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Flag check (Filter XML): {totalCount} total, {unresolvedCount} unresolved (after reset and sync), {clusterResolvedCount} cluster resolved, {individualResolvedCount} individual resolved\n");

                    // ✅ CRITICAL FIX: ALWAYS check Global XML for unresolved zones, even if Filter XML has clash zones
                    // This handles case where unresolved zone exists in Global XML but NOT in Filter XML (deleted/filtered out)
                    // Count unresolved zones from Global XML that aren't in Filter XML
                    if (selectedMepCategories != null && selectedMepCategories.Count > 0)
                    {
                        int globalUnresolvedCount = 0;
                        var filterXmlGuids = new HashSet<string>(clashZonesToCheck.Select(cz => cz.Id.ToString()), StringComparer.OrdinalIgnoreCase);

                        foreach (var category in selectedMepCategories)
                        {
                            if (string.IsNullOrWhiteSpace(category))
                                continue;

                            try
                            {
                                var globalIndex = GlobalIndexService.LoadOrCreate(_document, category);

                                // ✅ CRITICAL FIX: Use GetAllEntries to get entries from BOTH hierarchical and flat structures
                                var allEntries = GlobalIndexService.GetAllEntries(globalIndex).ToList();

                                if (allEntries != null && allEntries.Count > 0)
                                {
                                    // Count unresolved entries from Global XML that are NOT in Filter XML
                                    int categoryUnresolved = allEntries.Count(e =>
                                        !e.IsResolved && !e.IsClusterResolved &&
                                        !filterXmlGuids.Contains(e.Id)); // Only count if NOT in Filter XML

                                    globalUnresolvedCount += categoryUnresolved;

                                    if (!DeploymentConfiguration.DeploymentMode && categoryUnresolved > 0)
                                    {
                                        DebugLogger.Info($"[CLASH_DEBUG] Global XML check for '{category}': {categoryUnresolved} unresolved zones NOT in Filter XML");
                                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Global XML '{category}': {categoryUnresolved} unresolved zones NOT in Filter XML\n");
                                    }
                                }
                            }
                            catch (Exception globalEx)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Warning($"[CLASH_DEBUG] Error checking Global XML for category '{category}': {globalEx.Message}");
                                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ERROR checking Global XML for '{category}': {globalEx.Message}\n");
                            }
                        }

                        // ✅ CRITICAL FIX: Add unresolved zones from Global XML to the count
                        // This ensures we count ALL unresolved zones, even if they're not in Filter XML
                        unresolvedCount += globalUnresolvedCount;

                        if (!DeploymentConfiguration.DeploymentMode && globalUnresolvedCount > 0)
                        {
                            DebugLogger.Info($"[CLASH_DEBUG] ✅ Added {globalUnresolvedCount} unresolved zones from Global XML (not in Filter XML). Total unresolved: {unresolvedCount}");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ✅ Added {globalUnresolvedCount} unresolved zones from Global XML. Total unresolved: {unresolvedCount}\n");
                        }
                    }
                }
                else
                {
                    // ✅ CRITICAL FIX: If Filter XML is empty, check Global XML directly for unresolved zones
                    // This handles case where user deleted sleeves and Filter XML is empty/filtered out, but Global XML still has unresolved zones
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[CLASH_DEBUG] WARNING: No clash zones in Filter XML - checking Global XML directly for unresolved zones");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] WARNING: No clash zones in Filter XML - checking Global XML directly\n");

                    // Check Global XML directly for unresolved zones
                    if (selectedMepCategories != null && selectedMepCategories.Count > 0)
                    {
                        int globalUnresolvedCount = 0;
                        int globalTotalCount = 0;

                        foreach (var category in selectedMepCategories)
                        {
                            if (string.IsNullOrWhiteSpace(category))
                                continue;

                            try
                            {
                                var globalIndex = GlobalIndexService.LoadOrCreate(_document, category);

                                // ✅ CRITICAL FIX: Use GetAllEntries to get entries from BOTH hierarchical and flat structures
                                var allEntries = GlobalIndexService.GetAllEntries(globalIndex).ToList();

                                if (allEntries != null && allEntries.Count > 0)
                                {
                                    // Count unresolved entries from Global XML (single source of truth for flags)
                                    int categoryUnresolved = allEntries.Count(e => !e.IsResolved && !e.IsClusterResolved);
                                    int categoryTotal = allEntries.Count;

                                    globalUnresolvedCount += categoryUnresolved;
                                    globalTotalCount += categoryTotal;

                                    if (!DeploymentConfiguration.DeploymentMode && categoryUnresolved > 0)
                                    {
                                        DebugLogger.Info($"[CLASH_DEBUG] Global XML check for '{category}': {categoryUnresolved} unresolved out of {categoryTotal} total entries");
                                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Global XML '{category}': {categoryUnresolved}/{categoryTotal} unresolved\n");
                                    }
                                }
                            }
                            catch (Exception globalEx)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Warning($"[CLASH_DEBUG] Error checking Global XML for category '{category}': {globalEx.Message}");
                                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ERROR checking Global XML for '{category}': {globalEx.Message}\n");
                            }
                        }

                        // Use Global XML counts if Filter XML is empty
                        if (globalTotalCount > 0)
                        {
                            unresolvedCount = globalUnresolvedCount;
                            totalCount = globalTotalCount;
                            clusterResolvedCount = 0; // Can't determine from Global XML alone
                            individualResolvedCount = 0; // Can't determine from Global XML alone

                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[CLASH_DEBUG] Global XML check: {totalCount} total entries, {unresolvedCount} unresolved (Filter XML empty, using Global XML counts)");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Global XML: {totalCount} total, {unresolvedCount} unresolved (Filter XML empty)\n");
                        }
                        else
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[CLASH_DEBUG] No clash zones in Filter XML AND no entries in Global XML");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] No clash zones in Filter XML AND no entries in Global XML\n");
                        }
                    }
                }

                // ✅ CRITICAL FIX: Show "no new zones" message if:
                // 1. No new zones detected (finalNewZones == 0)
                // 2. All zones are resolved (unresolvedCount == 0)
                // 3. There are zones in the filter (totalCount > 0)
                if (finalNewZones == 0 && unresolvedCount == 0 && totalCount > 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] ✅ All {totalCount} existing clash zones are resolved ({clusterResolvedCount} cluster, {individualResolvedCount} individual) - no zones need processing");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ✅ All {totalCount} zones resolved - no zones need processing\n");

                    // ✅ CRITICAL FIX: Always show message when all zones are resolved (even if invalidatedCount > 0)
                    // Show user-friendly message
                    System.Windows.Forms.MessageBox.Show(
                        $"Refresh complete!\n\n" +
                        $"✅ All {totalCount} clash zones are resolved.\n" +
                        $"✅ No new clash zones detected.\n" +
                        $"✅ Total: {finalTotal} clash zones (all resolved).\n\n" +
                        $"No zones need processing - all intersections are handled.",
                        "Refresh Complete - No Zones Need Processing",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                }
                else if (finalNewZones == 0 && unresolvedCount > 0)
                {
                    // No new zones but there are unresolved existing zones - don't show "no new zones" message
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] Refresh complete: {unresolvedCount} unresolved zones need processing, {finalNewZones} new zones detected");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Refresh complete: {unresolvedCount} unresolved zones need processing, {finalNewZones} new zones detected\n");
                }
                else if (finalNewZones == 0 && totalCount == 0)
                {
                    // ✅ CRITICAL FIX: Check Global XML FIRST before showing "No Zones Found"
                    // If Global XML has unresolved zones, don't show "No Zones Found" - user should process them
                    // This handles case where sleeve was deleted (flag reset in Global XML) but clash zone not in Filter XML
                    if (unresolvedCount > 0)
                    {
                        // Global XML has unresolved zones even though Filter XML is empty
                        // Don't show "no zones found" - user should process these unresolved zones
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CLASH_DEBUG] Refresh complete: {unresolvedCount} unresolved zones in Global XML need processing (Filter XML empty/filtered, but Global XML has unresolved zones)");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Refresh complete: {unresolvedCount} unresolved zones in Global XML (Filter XML empty)\n");

                        // Don't show message box - let OK button be enabled so user can place sleeves
                        // The OK button enablement logic will handle this case
                    }
                    else
                    {
                        // ✅ CRITICAL FIX: Check if all intersections were filtered out because resolved
                        if (allIntersectionsFilteredAsResolved)
                        {
                            // All intersections were filtered out because they're resolved in Global XML
                            // Show "all resolved" message
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[CLASH_DEBUG] ✅ All intersections filtered out - all resolved in Global XML ({skippedResolvedCount} intersections skipped)");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ✅ All {skippedResolvedCount} intersections resolved - no zones need processing\n");

                            System.Windows.Forms.MessageBox.Show(
                                $"Refresh complete!\n\n" +
                                $"✅ All intersections are resolved.\n" +
                                $"✅ No new clash zones detected.\n" +
                                $"✅ All {skippedResolvedCount} intersections already have sleeves placed.\n\n" +
                                $"No zones need processing - all intersections are handled.",
                                "Refresh Complete - No Zones Need Processing",
                                System.Windows.Forms.MessageBoxButtons.OK,
                                System.Windows.Forms.MessageBoxIcon.Information);
                        }
                        else
                        {
                            // No zones at all - no intersections detected AND no unresolved zones in Global XML
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[CLASH_DEBUG] Refresh complete: No clash zones found");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Refresh complete: No clash zones found\n");

                            System.Windows.Forms.MessageBox.Show(
                                $"Refresh complete!\n\n" +
                                $"No clash zones found.\n" +
                                $"No intersections detected between selected MEP and structural elements.",
                                "Refresh Complete - No Zones Found",
                                System.Windows.Forms.MessageBoxButtons.OK,
                                System.Windows.Forms.MessageBoxIcon.Information);
                        }
                    }
                }

                var __refreshEnd = DateTime.Now;
                var __refreshMs = (long)(__refreshEnd - __refreshStart).TotalMilliseconds;
                SafeFileLogger.SafeAppendText("performance.log", $"REFRESH_END {__refreshEnd:O} DURATION_MS {__refreshMs}");

                // Step 10: Update parameter dropdowns
                _progressBar.Value = 90;
                _statusLabel.Text = "Updating parameter dropdowns...";

                // This will need to be handled by the UI
                // UpdateParameterDropdowns();

                _progressBar.Visible = false;
                _refreshButton.Enabled = true;

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info("[REFRESH] Refresh process completed successfully");

                // ⚠️ CRITICAL: Force memory cleanup after refresh completes
                try
                {
                    _memoryManager?.ForceCleanup();
                    var elapsedTime = _memoryManager?.ElapsedTime ?? TimeSpan.Zero;
                    SafeFileLogger.SafeAppendText("refresh_memory.log",
                        $"Refresh completed. Final memory: {_memoryManager?.CurrentMemoryMB ?? 0}MB, Elapsed: {elapsedTime:mm\\:ss}");
                }
                catch (Exception cleanupEx)
                {
                    SafeFileLogger.SafeAppendText("refresh_memory.log", $"Error during final cleanup: {cleanupEx.Message}");
                }

                // ✅ MEMORY PROFILING: Generate final report comparing theoretical vs actual memory usage
                try
                {
                    int totalClashZones = finalTotal; // Use final count from saved data
                    _memoryProfiler?.TakeSnapshot("REFRESH_COMPLETE", totalClashZones);
                    _memoryProfiler?.GenerateFinalReport(totalClashZones);

                    string memoryLogPath = SafeFileLogger.GetLogFilePath($"refresh_memory_profiling_{timestamp}.log");
                    string refreshLogPathFinal = SafeFileLogger.GetLogFilePath(refreshLogName); // Get full path again

                    // ✅ INVESTIGATION: Use SafeFileLogger only (no direct writes)
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] ✅ Memory profiling complete!\n");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] Memory profiling report: {memoryLogPath}\n");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] Open this file to see memory usage per clash zone vs theoretical estimates.\n");

                    // Update status label with log location
                    _statusLabel.Text = $"Refresh complete. Logs: {refreshLogPathFinal}";

                    // ✅ REMOVED: Message box prompt removed - logs are now working and can be found in status label
                    // Logs are written to: Refresh_*.log and refresh_memory_profiling_*.log in Logs directory
                    // Status label already shows log location, no need for popup dialog
                }
                catch (Exception profilerEx)
                {
                    SafeFileLogger.SafeAppendText("refresh_memory.log", $"Error generating memory profiling report: {profilerEx.Message}");
                }

                // ✅ MEMORY OPTIMIZATION: Final flush of batched logs before returning
                try
                {
                    if (batchedLogger != null)
                    {
                        batchedLogger.Flush();
                        batchedLogger.Dispose();
                    }
                }
                catch (Exception flushEx)
                {
                    // Non-fatal: Log the error but don't fail the entire refresh
                    SafeFileLogger.SafeAppendText("refresh_batch_log_error.log",
                        $"[{DateTime.Now}] Error flushing batched logs: {flushEx.Message}\n");
                }

                // ✅ FINAL ROOT CLEANUP: release large references and compact LOH
                try
                {
                    // Release large in-method collections/services
                    currentIntersections = null;

                    // Compact LOH once and force Gen2 collection
                    System.Runtime.GCSettings.LargeObjectHeapCompactionMode = System.Runtime.GCLargeObjectHeapCompactionMode.CompactOnce;
                    GC.Collect(2, GCCollectionMode.Forced, true);
                    GC.WaitForPendingFinalizers();
                    GC.Collect(2, GCCollectionMode.Forced, true);

                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [MEMORY_DEBUG] FINAL_CLEANUP: LOH compacted and Gen2 GC forced\n");
                }
                catch (Exception finalGcEx)
                {
                    SafeFileLogger.SafeAppendText("refresh_memory.log", $"Final GC/cleanup error: {finalGcEx}\n");
                }

                return Autodesk.Revit.UI.Result.Succeeded;
            }
            catch (Exception ex)
            {
                // Log error and return Failed result
                SafeFileLogger.SafeAppendText("refresh_errors.log", $"Refresh operation failed: {ex.Message}\n{ex.StackTrace}\n");
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[REFRESH] Refresh operation failed: {ex.Message}");

                _statusLabel.Text = $"Refresh failed: {ex.Message}";
                _progressBar.Visible = false;
                _refreshButton.Enabled = true;

                return Autodesk.Revit.UI.Result.Failed;
            }
            finally
            {
                // ✅ FIX 2: Clear geometry cache AFTER completing (even on error or early return)
                // This prevents memory accumulation across refresh operations
                try
                {
                    MepIntersectionService.ClearGeometryCache();
                    MepIntersectionService.ClearTransformCache();

                    // Force GC to release large objects from cache
                    GC.Collect(2, GCCollectionMode.Forced, true);
                    GC.WaitForPendingFinalizers();
                    GC.Collect(2, GCCollectionMode.Forced, true);
                }
                catch (Exception cacheEx)
                {
                    SafeFileLogger.SafeAppendText("refresh_memory.log", $"Cache clearing error: {cacheEx.Message}\n");
                }
            }
        }

        #region Helper Methods

        /// <summary>
        /// ✅ FIX 5: Analyze clash zone memory usage to identify memory culprits
        /// Samples first 100 clash zones and logs parameter frequency and memory estimates
        /// </summary>
        private void AnalyzeClashZoneMemory(List<Models.ClashZone> zones, string logFile)
        {
            try
            {
                if (zones == null || zones.Count == 0) return;

                int totalParams = 0;
                int totalStringChars = 0;
                var paramNameFrequency = new Dictionary<string, int>();

                int sampleSize = Math.Min(100, zones.Count);
                foreach (var cz in zones.Take(sampleSize))
                {
                    // Count parameters
                    int czParams = (cz.MepParameterValues?.Count ?? 0) +
                                  (cz.HostParameterValues?.Count ?? 0);
                    totalParams += czParams;

                    // Track parameter names
                    if (cz.MepParameterValues != null)
                    {
                        foreach (var kv in cz.MepParameterValues)
                        {
                            if (!paramNameFrequency.ContainsKey(kv.Key))
                                paramNameFrequency[kv.Key] = 0;
                            paramNameFrequency[kv.Key] = paramNameFrequency[kv.Key] + 1;
                            totalStringChars += (kv.Key?.Length ?? 0) +
                                               (kv.Value?.Length ?? 0);
                        }
                    }

                    if (cz.HostParameterValues != null)
                    {
                        foreach (var kv in cz.HostParameterValues)
                        {
                            if (!paramNameFrequency.ContainsKey(kv.Key))
                                paramNameFrequency[kv.Key] = 0;
                            paramNameFrequency[kv.Key] = paramNameFrequency[kv.Key] + 1;
                            totalStringChars += (kv.Key?.Length ?? 0) +
                                               (kv.Value?.Length ?? 0);
                        }
                    }
                }

                double avgParams = totalParams / (double)sampleSize;
                double avgStringBytes = (totalStringChars * 2) / (double)sampleSize; // UTF-16 = 2 bytes per char

                var topParams = paramNameFrequency
                    .OrderByDescending(kv => kv.Value)
                    .Take(20)
                    .ToList();

                var log = new StringBuilder();
                log.AppendLine($"[{DateTime.Now}] [MEMORY_ANALYSIS] Sample Size: {sampleSize}");
                log.AppendLine($"  Avg Parameters per Zone: {avgParams:F1}");
                log.AppendLine($"  Avg String Storage: {avgStringBytes:F0} bytes");
                log.AppendLine($"  Estimated Total: {avgParams * 150:F0} bytes/zone");
                log.AppendLine($"\nTop 20 Most Common Parameters:");
                foreach (var kvp in topParams)
                {
                    log.AppendLine($"    {kvp.Key}: {kvp.Value} occurrences");
                }

                SafeFileLogger.SafeAppendText(logFile, log.ToString());
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText(logFile,
                    $"[{DateTime.Now}] [MEMORY_ANALYSIS] Memory analysis failed: {ex.Message}\n");
            }
        }

        /// <summary>
        /// Filters clearance settings by MEP element category
        /// </summary>
        private Dictionary<string, double> FilterClearancesByCategory(Dictionary<string, double> allClearances, string category)
        {
            var categoryClearances = new Dictionary<string, double>();

            try
            {
                if (allClearances == null || allClearances.Count == 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLEARANCE_DEBUG] No clearance settings to filter for category: {category}");
                    return categoryClearances;
                }

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLEARANCE_DEBUG] Filtering {allClearances.Count} clearance settings for category: {category}");

                foreach (var kvp in allClearances)
                {
                    string key = kvp.Key;
                    double value = kvp.Value;

                    // Check if this clearance key belongs to the specified category
                    if (IsClearanceKeyForCategory(key, category))
                    {
                        categoryClearances[key] = value;
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CLEARANCE_DEBUG] Included clearance: {key} = {value}mm (for category: {category})");
                    }
                    else
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CLEARANCE_DEBUG] Excluded clearance: {key} = {value}mm (not for category: {category})");
                    }
                }

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLEARANCE_DEBUG] Filtered {categoryClearances.Count} clearance settings for category '{category}'");
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[CLEARANCE_DEBUG] Error in FilterClearancesByCategory: {ex.Message}");
            }

            return categoryClearances;
        }

        /// <summary>
        /// Checks if a clearance key belongs to a specific category
        /// </summary>
        private bool IsClearanceKeyForCategory(string key, string category)
        {
            if (string.IsNullOrEmpty(key) || string.IsNullOrEmpty(category))
                return false;

            string keyLower = key.ToLower();
            string categoryLower = category.ToLower();

            switch (categoryLower)
            {
                case "ducts":
                    return keyLower.Contains("duct") && !keyLower.Contains("accessory");

                case "duct accessories":
                    return keyLower.Contains("ductaccessories") ||
                           keyLower.Contains("duct_accessory") ||
                           keyLower.Contains("fire_damper") ||
                           keyLower.Contains("duct accessory");

                case "pipes":
                    return keyLower.Contains("pipe");

                case "cable trays":
                    return keyLower.Contains("cabletray") || keyLower.Contains("cable_tray");

                default:
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[CLEARANCE_DEBUG] Unknown category for clearance filtering: {category}");
                    return false;
            }
        }

        /// <summary>
        /// Gets the appropriate MEP type for a category
        /// </summary>
        private string GetCategorySpecificMepType(string category)
        {
            switch (category?.ToLower())
            {
                case "ducts":
                    return "Ducts";
                case "duct accessories":
                    return "Duct Accessories";
                case "pipes":
                    // Get actual opening type selection from UI settings (Rectangular/Circular)
                    var openingType = OpeningSettingsHelper.GetOpeningTypeForCategory("Pipes");
                    return $"Pipes_{openingType}";
                case "cable trays":
                    return "Cable Trays";
                default:
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[CLEARANCE_DEBUG] Unknown category for MEP type: {category}");
                    return "Unknown";
            }
        }
        // ============================================================================
        // FIX: Ensure Cable Tray Clash Zones are Saved for BOTH X-Walls and Y-Walls
        // Location: RefreshService.cs, FilterClashZonesByCategory method
        // ============================================================================
        /// <summary>
        /// Filters clash zones by MEP element category
        /// FIXED: Robust category matching for cable trays from linked files
        /// </summary>
        private List<Models.ClashZone> FilterClashZonesByCategory(List<Models.ClashZone> clashZones, string category, Document document, string refreshLogName)
        {
            var filteredZones = new List<Models.ClashZone>();

            try
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] Filtering {clashZones.Count} clash zones for category: {category}");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Filtering {clashZones.Count} clash zones for category: {category}\n");

                // DEBUG: Show first few clash zones to understand the data
                for (int i = 0; i < Math.Min(5, clashZones.Count); i++)
                {
                    var cz = clashZones[i];
                    if (OptimizationFlags.UseDiagnosticMode)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CLASH_DEBUG] Sample zone {i}: MEP ID={cz.MepElementId}, Cat='{cz.MepElementCategory}', Structural={cz.StructuralElementType}");
                        SafeFileLogger.SafeAppendText(refreshLogName,
                            $"[{DateTime.Now}] [CLASH_DEBUG] Sample zone {i}: MEP ID={cz.MepElementId}, Cat='{cz.MepElementCategory}', Structural={cz.StructuralElementType}\n");
                    }
                }

                // Count zones by structural type for debugging
                var byStructType = clashZones.GroupBy(cz => cz.StructuralElementType).ToDictionary(g => g.Key, g => g.Count());
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] Zones by structural type: {string.Join(", ", byStructType.Select(kv => $"{kv.Key}={kv.Value}"))}");
                SafeFileLogger.SafeAppendText(refreshLogName,
                    $"[{DateTime.Now}] [CLASH_DEBUG] Zones by structural type: {string.Join(", ", byStructType.Select(kv => $"{kv.Key}={kv.Value}"))}\n");

                foreach (var clashZone in clashZones)
                {
                    try
                    {
                        // ✅ CRITICAL FIX: Use cached category from ClashZone (reliable for linked files)
                        var elementCategory = clashZone.MepElementCategory ?? "Unknown";

                        // 🔍 ENHANCED DEBUGGING: Log every zone being checked
                        var structType = clashZone.StructuralElementType ?? "Unknown";
                        var structNormal = clashZone.StructuralElementNormal;
                        var wallOrientation = "Unknown";

                        if (structNormal != null)
                        {
                            double absX = Math.Abs(structNormal.X);
                            double absY = Math.Abs(structNormal.Y);
                            wallOrientation = absX > absY ? "X-Wall" : "Y-Wall";
                        }

                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CLASH_DEBUG] Zone {clashZone.Id}: Cat='{elementCategory}', Struct='{structType}', Orient='{wallOrientation}', Checking against '{category}'");
                        SafeFileLogger.SafeAppendText(refreshLogName,
                            $"[{DateTime.Now}] [CLASH_DEBUG] Zone {clashZone.Id}: Cat='{elementCategory}', Struct='{structType}', Orient='{wallOrientation}', Checking against '{category}'\n");

                        // ✅ FIX: Enhanced category matching with fallback for linked files
                        bool isMatch = IsElementInCategoryEnhanced(elementCategory, category, clashZone, document);

                        if (isMatch)
                        {
                            filteredZones.Add(clashZone);
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[CLASH_DEBUG] ✅ Zone {clashZone.Id} MATCHED '{category}' - Cat='{elementCategory}', Orient='{wallOrientation}'");
                            SafeFileLogger.SafeAppendText(refreshLogName,
                                $"[{DateTime.Now}] [CLASH_DEBUG] ✅ Zone {clashZone.Id} MATCHED '{category}' - Cat='{elementCategory}', Orient='{wallOrientation}'\n");
                        }
                        else
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[CLASH_DEBUG] ❌ Zone {clashZone.Id} NOT matched '{category}' - Cat='{elementCategory}'");
                            SafeFileLogger.SafeAppendText(refreshLogName,
                                $"[{DateTime.Now}] [CLASH_DEBUG] ❌ Zone {clashZone.Id} NOT matched '{category}' - Cat='{elementCategory}'\n");
                        }
                    }
                    catch (Exception ex)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error($"[CLASH_DEBUG] Error filtering clash zone {clashZone.Id}: {ex.Message}");
                    }
                }

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] ✅ Filtered {filteredZones.Count}/{clashZones.Count} clash zones for category '{category}'");
                SafeFileLogger.SafeAppendText(refreshLogName,
                    $"[{DateTime.Now}] [CLASH_DEBUG] ✅ Filtered {filteredZones.Count}/{clashZones.Count} clash zones for category '{category}'\n");
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[CLASH_DEBUG] Error in FilterClashZonesByCategory: {ex.Message}");
            }

            return filteredZones;
        }
        /// <summary>
        /// Enhanced category matching with fallback for linked file elements
        /// FIXED: Handles cable trays, conduits, and all MEP categories robustly
        /// </summary>
        private bool IsElementInCategoryEnhanced(string elementCategory, string requestedCategory, Models.ClashZone clashZone, Document document)
        {
            try
            {
                // Convert to lowercase for case-insensitive comparison
                var elementCatLower = elementCategory?.ToLower() ?? "";
                var requestedCatLower = requestedCategory?.ToLower() ?? "";

                // ✅ PHASE 1: Standard category matching (fast path)
                bool standardMatch = IsElementInCategory(elementCategory, requestedCategory);

                if (standardMatch)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CATEGORY_MATCH] ✅ Standard match for '{elementCategory}' -> '{requestedCategory}'");
                    return true;
                }

                // ✅ PHASE 2: Enhanced matching for problematic categories
                if (requestedCatLower == "cable trays")
                {
                    // Cable Trays: Check multiple variations
                    bool isCableTray =
                        elementCatLower.Contains("cable") ||
                        elementCatLower.Contains("tray") ||
                        elementCatLower.Contains("conduit") ||
                        elementCatLower == "cable tray" ||
                        elementCatLower == "cable trays" ||
                        elementCatLower == "cable tray fittings" ||
                        elementCatLower == "cable tray fitting" ||
                        elementCatLower == "conduit" ||
                        elementCatLower == "conduits" ||
                        elementCatLower == "conduit fittings" ||
                        elementCatLower == "conduit fitting";

                    if (isCableTray)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CATEGORY_MATCH] ✅ Cable tray enhanced match for '{elementCategory}'");
                        return true;
                    }

                    // ✅ PHASE 3: Fallback - check actual element from Revit (for linked files)
                    try
                    {
                        var mepElement = ElementRetrievalService.GetElementFromDocumentOrLinked(document, clashZone.MepElementId, enableLogging: false);
                        if (mepElement != null)
                        {
                            var actualCategory = mepElement.Category?.Name ?? "";
                            var actualCatLower = actualCategory.ToLower();

                            bool isCableTrayActual =
                                actualCatLower.Contains("cable") ||
                                actualCatLower.Contains("tray") ||
                                actualCatLower.Contains("conduit");

                            if (isCableTrayActual)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[CATEGORY_MATCH] ✅ Cable tray FALLBACK match from actual element: '{actualCategory}'");
                                // Update cached category for future use
                                clashZone.MepElementCategory = actualCategory;
                                return true;
                            }
                            else
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[CATEGORY_MATCH] ❌ Cable tray fallback check failed: actual category = '{actualCategory}'");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[CATEGORY_MATCH] Fallback element check failed: {ex.Message}");
                    }
                }

                // Similar enhanced matching for other categories
                if (requestedCatLower == "pipes")
                {
                    bool isPipe =
                        elementCatLower.Contains("pipe") ||
                        elementCatLower == "pipes" ||
                        elementCatLower == "pipe curves" ||
                        elementCatLower == "pipe fittings" ||
                        elementCatLower == "pipe accessories";

                    if (isPipe)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CATEGORY_MATCH] ✅ Pipe enhanced match for '{elementCategory}'");
                        return true;
                    }
                }

                if (requestedCatLower == "ducts")
                {
                    bool isDuct =
                        (elementCatLower.Contains("duct") && !elementCatLower.Contains("accessory")) ||
                        elementCatLower == "ducts" ||
                        elementCatLower == "duct curves" ||
                        elementCatLower == "duct fittings";

                    if (isDuct)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CATEGORY_MATCH] ✅ Duct enhanced match for '{elementCategory}'");
                        return true;
                    }
                }

                if (requestedCatLower == "duct accessories")
                {
                    bool isDuctAccessory =
                        elementCatLower.Contains("duct") && elementCatLower.Contains("accessor") ||
                        elementCatLower == "duct accessories" ||
                        elementCatLower == "duct accessory";

                    if (isDuctAccessory)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CATEGORY_MATCH] ✅ Duct accessory enhanced match for '{elementCategory}'");
                        return true;
                    }
                }

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CATEGORY_MATCH] ❌ No match found for '{elementCategory}' -> '{requestedCategory}'");
                return false;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[CATEGORY_MATCH] Error in IsElementInCategoryEnhanced: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Gets element from document or linked documents
        /// </summary>
        // ✅ OOP REFACTORING: Removed duplicate GetElementFromDocumentOrLinked - now uses ElementRetrievalService.GetElementFromDocumentOrLinked()
        // All calls have been replaced with ElementRetrievalService.GetElementFromDocumentOrLinked(document, elementId, enableLogging: false)

        /// <summary>
        /// Gets the category name of an element
        /// </summary>
        private string GetElementCategory(Element element)
        {
            try
            {
                if (element?.Category != null)
                {
                    return element.Category.Name;
                }
                return "Unknown";
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[CLASH_DEBUG] Error getting element category: {ex.Message}");
                return "Unknown";
            }
        }

        /// <summary>
        /// Gets element category with fallback logic to handle inconsistent category detection
        /// </summary>
        private string GetElementCategoryWithFallback(Element element, ElementId elementId, Document document, string refreshLogName)
        {
            try
            {
                // First try the standard category detection
                var standardCategory = GetElementCategory(element);

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] Standard category detection for element {elementId}: '{standardCategory}'");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Standard category detection for element {elementId}: '{standardCategory}'\n");

                // If we get a meaningful category, use it
                if (!string.IsNullOrEmpty(standardCategory) &&
                    standardCategory != "Unknown" &&
                    !standardCategory.Contains("Sketch") &&
                    !standardCategory.Contains("Room Tags"))
                {
                    return standardCategory;
                }

                // Fallback: Use element type and family information to determine category
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] Triggering fallback category detection for element {elementId} with standard category '{standardCategory}'");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Triggering fallback category detection for element {elementId} with standard category '{standardCategory}'\n");

                var fallbackCategory = GetCategoryFromElementType(element);

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] Fallback category detection for element {elementId}: '{fallbackCategory}' (original: '{standardCategory}')");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Fallback category detection for element {elementId}: '{fallbackCategory}' (original: '{standardCategory}')\n");

                // If fallback also fails, try to infer from intersection detection results
                if (fallbackCategory == "Unknown")
                {
                    // Since we know these elements were detected as pipes during intersection detection,
                    // and they're being categorized as "Automatic Sketch Dimensions", they're likely pipes
                    if (standardCategory.Contains("Sketch"))
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CLASH_DEBUG] Inferring 'Pipes' category for sketch element {elementId} based on intersection detection");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Inferring 'Pipes' category for sketch element {elementId} based on intersection detection\n");
                        return "Pipes";
                    }
                }

                return fallbackCategory;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[CLASH_DEBUG] Error in GetElementCategoryWithFallback for element {elementId}: {ex.Message}");
                return "Unknown";
            }
        }

        /// <summary>
        /// Determines category from element type and family information
        /// </summary>
        private string GetCategoryFromElementType(Element element)
        {
            try
            {
                if (element == null) return "Unknown";

                // Check element type
                var elementType = element.GetType();
                var typeName = elementType.Name;

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] Element type analysis: TypeName='{typeName}'");

                // Check for specific MEP element types
                if (element is FamilyInstance familyInstance)
                {
                    var familyName = familyInstance.Symbol?.FamilyName ?? "Unknown";
                    var categoryName = familyInstance.Category?.Name ?? "Unknown";

                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] FamilyInstance analysis: FamilyName='{familyName}', CategoryName='{categoryName}'");

                    // Check family name patterns
                    if (familyName.ToLower().Contains("pipe") ||
                        familyName.ToLower().Contains("fitting") ||
                        familyName.ToLower().Contains("valve") ||
                        familyName.ToLower().Contains("pump"))
                    {
                        return "Pipes";
                    }

                    if (familyName.ToLower().Contains("duct") ||
                        familyName.ToLower().Contains("damper") ||
                        familyName.ToLower().Contains("vav"))
                    {
                        return "Duct Accessories";
                    }

                    if (familyName.ToLower().Contains("cable") ||
                        familyName.ToLower().Contains("tray"))
                    {
                        return "Cable Trays";
                    }
                }

                // Check for MEPCurve elements (pipes, ducts, cable trays)
                if (element is MEPCurve mepCurve)
                {
                    var categoryName = mepCurve.Category?.Name ?? "Unknown";

                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] MEPCurve analysis: CategoryName='{categoryName}'");

                    if (categoryName.Contains("Pipe") || categoryName.Contains("Piping"))
                    {
                        return "Pipes";
                    }

                    if (categoryName.Contains("Duct") || categoryName.Contains("Ductwork"))
                    {
                        return "Ducts";
                    }

                    if (categoryName.Contains("Cable") || categoryName.Contains("Tray"))
                    {
                        return "Cable Trays";
                    }
                }

                // Check for Conduit elements (using string comparison for accessibility)
                var elementTypeName = element.GetType().Name;
                if (elementTypeName.Contains("Conduit"))
                {
                    return "Cable Trays";
                }

                // Check for Electrical elements (using string comparison for accessibility)
                if (elementTypeName.Contains("ElectricalSystem") || elementTypeName.Contains("ElectricalEquipment"))
                {
                    return "Electrical Equipment";
                }

                // Final fallback - return the original category if it's not sketch-related
                var originalCategory = element.Category?.Name ?? "Unknown";
                if (!originalCategory.Contains("Sketch") && !originalCategory.Contains("Room"))
                {
                    return originalCategory;
                }

                return "Unknown";
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[CLASH_DEBUG] Error in GetCategoryFromElementType: {ex.Message}");
                return "Unknown";
            }
        }

        /// <summary>
        /// Checks if element category matches the requested category
        /// </summary>
        private bool IsElementInCategory(string elementCategory, string requestedCategory)
        {
            try
            {
                // Convert both to lowercase for case-insensitive comparison
                var elementCatLower = elementCategory?.ToLower() ?? "";
                var requestedCatLower = requestedCategory?.ToLower() ?? "";

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] Checking if element category '{elementCategory}' matches requested '{requestedCategory}'");

                switch (requestedCatLower)
                {
                    case "duct accessories":
                        // Match actual Revit category names for duct accessories (check this FIRST)
                        return elementCatLower == "duct accessories" ||
                               elementCatLower == "duct accessory" ||
                               elementCatLower.Contains("duct accessory");

                    case "ducts":
                        // Match actual Revit category names for ducts (but NOT duct accessories)
                        return (elementCatLower == "duct curves" ||
                               elementCatLower == "duct fittings" ||
                               elementCatLower == "duct terminals" ||
                               elementCatLower == "duct curve" ||
                               elementCatLower == "duct fitting" ||
                               elementCatLower == "duct terminal" ||
                               elementCatLower.Contains("duct")) &&
                               !elementCatLower.Contains("accessory");

                    case "pipes":
                        // Match actual Revit category names for pipes
                        return elementCatLower == "pipes" ||
                               elementCatLower == "pipe curves" ||
                               elementCatLower == "pipe fittings" ||
                               elementCatLower == "pipe terminals" ||
                               elementCatLower == "pipe curve" ||
                               elementCatLower == "pipe fitting" ||
                               elementCatLower == "pipe terminal" ||
                               elementCatLower.Contains("pipe");

                    case "cable trays":
                        // Match actual Revit category names for cable trays (enhanced matching)
                        return elementCatLower == "cable trays" ||
                               elementCatLower == "cable tray fittings" ||
                               elementCatLower == "cable tray" ||
                               elementCatLower.Contains("cable") ||
                               elementCatLower.Contains("tray");

                    default:
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[CLASH_DEBUG] Unknown requested category: '{requestedCategory}'");
                        return false;
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[CLASH_DEBUG] Error checking category match: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Loads the most recently saved OpeningFilter XML from any project Filters folder under AppData.
        /// Used to read saved host categories when filtering existing zones.
        /// </summary>
        private Models.OpeningFilter LoadLatestOpeningFilter()
        {
            try
            {
                var projectsRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Projects");
                if (!Directory.Exists(projectsRoot)) return null;

                string latestFile = null;
                DateTime latestWrite = DateTime.MinValue;

                foreach (var projectDir in Directory.GetDirectories(projectsRoot))
                {
                    var filtersDir = Path.Combine(projectDir, "Filters");
                    if (!Directory.Exists(filtersDir)) continue;

                    foreach (var xml in Directory.GetFiles(filtersDir, "*.xml"))
                    {
                        var t = File.GetLastWriteTime(xml);
                        if (t > latestWrite)
                        {
                            latestWrite = t;
                            latestFile = xml;
                        }
                    }
                }

                if (string.IsNullOrEmpty(latestFile)) return null;

                // Reuse a static serializer to avoid per-refresh type cache allocations
                var serializer = OpeningFilterSerializer;
                using (var reader = new StreamReader(latestFile))
                {
                    return (Models.OpeningFilter)serializer.Deserialize(reader);
                }
            }
            catch
            {
                return null;
            }
        }
        /// <summary>
        /// Load existing clash zones from Filter XML files (selected filter + categories)
        /// This preserves flags (IsResolved, IsClusterResolved) from previous runs
        /// 
        /// ⚠️ CRITICAL: Even when UseGlobalCategoryIndexForRefresh is true, we MUST load existing clash zones
        /// to populate the known intersections map. Otherwise, all intersections are treated as new,
        /// causing duplicate clash zones and ignoring resolved flags from Global XML.
        /// </summary>
        private Models.ClashZoneStorage LoadExistingClashZonesFromFilterXml(List<string> selectedFilterNames, List<string> selectedCategories)
        {
            // ⚠️ CRITICAL FIX: Always load existing clash zones, even when optimization flag is true
            // The known intersections map requires existing clash zones to prevent duplicates
            // The optimization flag should only affect how we query Global XML, not whether we load Filter XML
            var mergedStorage = new Models.ClashZoneStorage
            {
                ClashZones = new List<Models.ClashZone>(),
                LastUpdated = DateTime.Now
            };

            try
            {
                var filtersDirectory = ProjectPathService.GetFiltersDirectory(_document);

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[LoadExistingClashZones] Looking for XML files in: {filtersDirectory}");
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] [{DateTime.Now}] [LoadExistingClashZones] Looking for XML files in: {filtersDirectory}\n");

                if (!Directory.Exists(filtersDirectory))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[LoadExistingClashZones] Filters directory does not exist!");
                    return mergedStorage;
                }

                // 🔥 DEBUG: List ALL XML files in the directory
                var allXmlFiles = Directory.GetFiles(filtersDirectory, "*.xml");
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[LoadExistingClashZones] Found {allXmlFiles.Length} total XML files");
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] [{DateTime.Now}] [LoadExistingClashZones] Found {allXmlFiles.Length} total XML files\n");

                foreach (var xmlFile in allXmlFiles)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[LoadExistingClashZones] - {Path.GetFileName(xmlFile)}");
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] [{DateTime.Now}] [LoadExistingClashZones] - {Path.GetFileName(xmlFile)}\n");
                }

                // Load clash zones from each category-specific XML file
                foreach (var filterName in selectedFilterNames)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[LoadExistingClashZones] Processing filter: {filterName}");
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] [{DateTime.Now}] [LoadExistingClashZones] Processing filter: {filterName}\n");

                    foreach (var category in selectedCategories)
                    {
                        var pattern = $"{filterName}_{category.ToLower().Replace(" ", "_")}.xml";
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[LoadExistingClashZones] Looking for pattern: {pattern}");
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CLASH_DEBUG] [{DateTime.Now}] [LoadExistingClashZones] Looking for pattern: {pattern}\n");

                        var matchingFiles = Directory.GetFiles(filtersDirectory, pattern);

                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[LoadExistingClashZones] Found {matchingFiles.Length} matching files for pattern: {pattern}");
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CLASH_DEBUG] [{DateTime.Now}] [LoadExistingClashZones] Found {matchingFiles.Length} matching files for pattern: {pattern}\n");

                        if (matchingFiles.Length > 0)
                        {
                            var xmlFile = matchingFiles.First();
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[LoadExistingClashZones] ✅ Loading from {Path.GetFileName(xmlFile)}");

                            try
                            {
                                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(Models.OpeningFilter));
                                using (var reader = new StreamReader(xmlFile))
                                {
                                    var filter = (Models.OpeningFilter)serializer.Deserialize(reader);
                                    if (filter?.ClashZoneStorage?.AllZones != null)
                                    {
                                        mergedStorage.ClashZones.AddRange(filter.ClashZoneStorage.AllZones);
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[LoadExistingClashZones] ✅ Loaded {filter.ClashZoneStorage.AllZones.Count} clash zones from {Path.GetFileName(xmlFile)}");
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[CLASH_DEBUG] [{DateTime.Now}] [LoadExistingClashZones] ✅ Loaded {filter.ClashZoneStorage.AllZones.Count} clash zones\n");
                                    }
                                    else
                                    {
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Warning($"[LoadExistingClashZones] ❌ No clash zones in {Path.GetFileName(xmlFile)}");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Error($"[LoadExistingClashZones] ❌ Error loading {Path.GetFileName(xmlFile)}: {ex.Message}");
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[CLASH_DEBUG] [{DateTime.Now}] [LoadExistingClashZones] ❌ ERROR: {ex.Message}\n");
                            }
                        }
                    }
                }

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[LoadExistingClashZones] ✅ Total clash zones loaded: {mergedStorage.ClashZones.Count}");
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] [{DateTime.Now}] [LoadExistingClashZones] ✅ FINAL: Total clash zones loaded: {mergedStorage.ClashZones.Count}\n");
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[LoadExistingClashZones] Error: {ex.Message}");
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] [{DateTime.Now}] [LoadExistingClashZones] EXCEPTION: {ex.Message}\n");
            }

            return mergedStorage;
        }

        #endregion

        /// <summary>
        /// ✅ FRESH RUN DETECTION: Check if any XML files exist for selected filters/categories
        /// Returns true if at least one category has either:
        /// - Filter XML file: {filterName}_{categoryXmlSuffix}.xml (for any selected filter)
        /// - Global XML file: {categoryXmlSuffix}_global.xml
        /// </summary>
        /// <param name="selectedFilterNames">Selected filter names</param>
        /// <param name="selectedCategories">Selected MEP categories</param>
        /// <returns>True if any XML files exist, false if none exist (fresh run)</returns>
        private bool CheckIfAnyXmlFilesExist(List<string> selectedFilterNames, List<string> selectedCategories)
        {
            try
            {
                if (selectedCategories == null || selectedCategories.Count == 0)
                    return false; // No categories selected → no XML files to check

                var filtersDirectory = ProjectPathService.GetFiltersDirectory(_document);
                if (!Directory.Exists(filtersDirectory))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[FRESH-RUN-CHECK] Filters directory does not exist - treating as fresh run");
                    return false; // No filters directory → fresh run
                }

                // Check each selected category
                foreach (var category in selectedCategories)
                {
                    if (string.IsNullOrWhiteSpace(category))
                        continue;

                    // Get XML suffix for category (e.g., "Pipes" → "pipes", "Duct Accessories" → "duct_accessories")
                    string categoryXmlSuffix = MepCategoryConstants.GetXmlSuffix(category);

                    // Check 1: Filter XML files - {filterName}_{categoryXmlSuffix}.xml
                    if (selectedFilterNames != null && selectedFilterNames.Count > 0)
                    {
                        foreach (var filterName in selectedFilterNames)
                        {
                            if (string.IsNullOrWhiteSpace(filterName))
                                continue;

                            string filterXmlPattern = $"{filterName}_{categoryXmlSuffix}.xml";
                            string filterXmlPath = Path.Combine(filtersDirectory, filterXmlPattern);

                            if (File.Exists(filterXmlPath))
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[FRESH-RUN-CHECK] ✅ Found Filter XML: {filterXmlPattern} for category '{category}'");
                                return true; // Found at least one filter XML file
                            }
                        }
                    }

                    // Check 2: Global XML file - {categoryXmlSuffix}_global.xml
                    string globalXmlFileName = $"{categoryXmlSuffix}_global.xml";
                    string globalXmlPath = Path.Combine(filtersDirectory, globalXmlFileName);

                    if (File.Exists(globalXmlPath))
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[FRESH-RUN-CHECK] ✅ Found Global XML: {globalXmlFileName} for category '{category}'");
                        return true; // Found at least one global XML file
                    }
                }

                // No XML files found for any selected category → fresh run
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[FRESH-RUN-CHECK] ❌ No XML files found for selected categories - treating as fresh run");
                return false;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[FRESH-RUN-CHECK] Error checking XML files: {ex.Message} - treating as fresh run");
                return false; // Error → treat as fresh run to be safe
            }
        }
        /// <summary>
        /// Helper method to calculate intersection point between MEP and structural elements
        /// ⚠️ PERFORMANCE NOTE: This method uses get_Geometry() which is VERY EXPENSIVE
        /// However, geometry extraction is necessary here because:
        /// 1. We need to find the exact intersection point between MEP line and structural solid
        /// 2. Bounding box would only give approximate location, not precise intersection
        /// 3. This is called infrequently during clash zone loading, not in tight loops
        /// Consider caching geometry results if this method is called repeatedly for same elements
        /// </summary>
        private XYZ CalculateIntersectionPoint(Element mepElement, Element structuralElement)
        {
            try
            {
                // ✅ PERFORMANCE: Geometry extraction is expensive but necessary for accurate intersection calculation
                // Bounding box would not provide precise intersection point required here
                var mepGeometry = mepElement.get_Geometry(new Options());
                var structuralGeometry = structuralElement.get_Geometry(new Options());

                if (mepGeometry == null || structuralGeometry == null) return null;

                // Find intersection points using the same logic as MepIntersectionService
                var intersectionPoints = new List<XYZ>();

                // Get MEP element line (same as MepIntersectionService)
                Line? line = null;
                foreach (GeometryObject geo in mepGeometry)
                {
                    if (geo is Curve curve)
                    {
                        line = curve as Line;
                        if (line != null) break;
                    }
                }

                if (line == null) return null;

                // Get structural element solid and find intersections
                foreach (GeometryObject structGeo in structuralGeometry)
                {
                    if (structGeo is Solid structSolid)
                    {
                        // Use the same intersection logic as MepIntersectionService
                        foreach (Face face in structSolid.Faces)
                        {
                            if (face == null) continue;
                            IntersectionResultArray? ira;
                            var res = face.Intersect(line, out ira);
                            if (res == SetComparisonResult.Overlap && ira != null)
                            {
                                foreach (IntersectionResult ir in ira)
                                {
                                    intersectionPoints.Add(ir.XYZPoint);
                                }
                                if (intersectionPoints.Count > 0)
                                {
                                    // Early exit after finding first intersection (same as MepIntersectionService)
                                    return intersectionPoints[0];
                                }
                            }
                        }
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[CLASH_DEBUG] Failed to calculate intersection point: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Transform coordinates from linked document to active document
        /// </summary>
        // ✅ OOP REFACTORING: Removed duplicate TransformToActiveDocumentCoordinates - now uses CoordinateTransformService.TransformToActiveDocumentCoordinates()

        /// <summary>
        /// Extracts the base filter name by removing any existing category suffix.
        /// This prevents duplicates like "Ventilation_ducts_ducts" or "Ventilation_duct_accessories_accessories"
        /// </summary>
        private string ExtractBaseFilterName(string filterName, string normalizedCategory)
        {
            if (string.IsNullOrWhiteSpace(filterName)) return filterName;
            if (string.IsNullOrWhiteSpace(normalizedCategory)) return filterName;

            // Check if filter name ends with the category suffix
            string suffixPattern = $"_{normalizedCategory}";
            if (filterName.EndsWith(suffixPattern, StringComparison.OrdinalIgnoreCase))
            {
                return filterName.Substring(0, filterName.Length - suffixPattern.Length);
            }

            return filterName;
        }

        private void SyncFilterStorageFromZones(
            OpeningFilter? targetFilter,
            List<Models.ClashZone>? allClashZones,
            string baseFilterName)
        {
            if (targetFilter?.ClashZoneStorage == null || allClashZones == null)
                return;

            var distinctZones = allClashZones
                .Where(z => z != null)
                .GroupBy(z => z.Id)
                .Select(g => g.First())
                .ToList();

            var filterGroups = new List<Models.FilterGroupForStorage>();

            var zonesByCategory = distinctZones
                .Where(z => !string.IsNullOrWhiteSpace(z.MepElementCategory))
                .GroupBy(z => z.MepElementCategory!, StringComparer.OrdinalIgnoreCase);

            foreach (var categoryGroup in zonesByCategory)
            {
                var categoryName = categoryGroup.Key;
                var groupName = BuildCategoryFilterGroupName(baseFilterName, categoryName);

                var comboList = categoryGroup
                    .GroupBy(GetFileComboKeyForTree)
                    .Where(g => !string.IsNullOrWhiteSpace(g.Key.Linked) && !string.IsNullOrWhiteSpace(g.Key.Host))
                    .Select(g => new Models.FilterFileComboGroup
                    {
                        LinkedFile = g.Key.Linked,
                        HostFile = g.Key.Host,
                        ProcessedAt = DateTime.Now,
                        ClashZones = g.Select(z => z).ToList()
                    })
                    .ToList();

                if (comboList.Count == 0)
                    continue;

                filterGroups.Add(new Models.FilterGroupForStorage
                {
                    Name = groupName,
                    FileCombos = comboList
                });
            }

            if (filterGroups.Count > 0)
            {
                targetFilter.ClashZoneStorage.Filters = filterGroups;
            }

            targetFilter.ClashZoneStorage.ClashZones = distinctZones.ToList();
            targetFilter.ClashZoneStorage.LastUpdated = DateTime.Now;
            targetFilter.LastModified = DateTime.Now;
        }

        private static string BuildCategoryFilterGroupName(string baseFilterName, string category)
        {
            var normalizedCategory = (category ?? string.Empty).Trim().ToLowerInvariant();

            if (string.IsNullOrWhiteSpace(baseFilterName))
                return string.IsNullOrWhiteSpace(normalizedCategory) ? "Unknown" : normalizedCategory;

            if (string.IsNullOrWhiteSpace(normalizedCategory))
                return baseFilterName;

            var suffix = "_" + normalizedCategory.Replace(' ', '_');
            return baseFilterName.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)
                ? baseFilterName
                : baseFilterName + suffix;
        }

        private (string Linked, string Host) GetFileComboKeyForTree(Models.ClashZone zone)
        {
            if (zone == null)
                return ("unknown-linked", "unknown-host");

            var linked = FirstNonEmptyForTree(
                zone.SourceDocKey,
                zone.DocumentPath,
                zone.SourceDocKey != null ? System.IO.Path.GetFileName(zone.SourceDocKey) : null);

            var host = FirstNonEmptyForTree(
                zone.HostDocKey,
                zone.StructuralElementDocumentTitle,
                zone.HostDocKey != null ? System.IO.Path.GetFileName(zone.HostDocKey) : null);

            return (
                string.IsNullOrWhiteSpace(linked) ? "unknown-linked" : NormalizeFileNameForTree(linked),
                string.IsNullOrWhiteSpace(host) ? "unknown-host" : NormalizeFileNameForTree(host));
        }

        private static string FirstNonEmptyForTree(params string?[] values)
        {
            if (values == null) return string.Empty;
            foreach (var value in values)
            {
                if (!string.IsNullOrWhiteSpace(value))
                    return value!;
            }
            return string.Empty;
        }

        /// <summary>
        /// ✅ HELPER: Copies tree structure from targetFilter, filtering to only include clash zones for specified category
        /// Creates FilterGroupForStorage → FilterFileComboGroup → ClashZone tree structure
        /// </summary>
        private List<Models.FilterGroupForStorage> CopyTreeStructureForCategory(
            List<Models.FilterGroupForStorage> sourceFilters,
            string category,
            string baseFilterName,
            List<Models.ClashZone> mergedCategoryClashZones)
        {
            if (sourceFilters == null || sourceFilters.Count == 0)
            {
                // No source tree structure - create new tree structure from mergedCategoryClashZones
                return CreateTreeStructureFromClashZones(mergedCategoryClashZones, baseFilterName);
            }

            var result = new List<Models.FilterGroupForStorage>();

            // Find the filter group matching baseFilterName
            var sourceFilterGroup = sourceFilters.FirstOrDefault(f => f.Name == baseFilterName);
            if (sourceFilterGroup == null)
            {
                // Filter group not found - create new tree structure from mergedCategoryClashZones
                return CreateTreeStructureFromClashZones(mergedCategoryClashZones, baseFilterName);
            }

            // Create new filter group for this category
            var categoryFilterGroup = new Models.FilterGroupForStorage
            {
                Name = baseFilterName,
                FileCombos = new List<Models.FilterFileComboGroup>()
            };

            // Copy FileComboGroups, filtering clash zones by category
            if (sourceFilterGroup.FileCombos != null)
            {
                foreach (var sourceFileCombo in sourceFilterGroup.FileCombos)
                {
                    if (sourceFileCombo?.ClashZones == null) continue;

                    // Filter clash zones for this category
                    var categoryClashZones = sourceFileCombo.ClashZones
                        .Where(cz => string.Equals(cz.MepElementCategory, category, StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    if (categoryClashZones.Count > 0)
                    {
                        var categoryFileCombo = new Models.FilterFileComboGroup
                        {
                            LinkedFile = sourceFileCombo.LinkedFile,
                            HostFile = sourceFileCombo.HostFile,
                            ProcessedAt = sourceFileCombo.ProcessedAt,
                            ClashZones = categoryClashZones
                        };
                        categoryFilterGroup.FileCombos.Add(categoryFileCombo);
                    }
                }
            }

            // If no clash zones found in tree structure, create from mergedCategoryClashZones
            if (categoryFilterGroup.FileCombos.Count == 0 && mergedCategoryClashZones != null && mergedCategoryClashZones.Count > 0)
            {
                return CreateTreeStructureFromClashZones(mergedCategoryClashZones, baseFilterName);
            }

            result.Add(categoryFilterGroup);
            return result;
        }

        /// <summary>
        /// ✅ HELPER: Creates tree structure from flat clash zones list
        /// Groups by file combo (LinkedFile + HostFile)
        /// </summary>
        private List<Models.FilterGroupForStorage> CreateTreeStructureFromClashZones(
            List<Models.ClashZone> clashZones,
            string baseFilterName)
        {
            if (clashZones == null || clashZones.Count == 0)
            {
                return new List<Models.FilterGroupForStorage>();
            }

            var result = new List<Models.FilterGroupForStorage>();
            var filterGroup = new Models.FilterGroupForStorage
            {
                Name = baseFilterName,
                FileCombos = new List<Models.FilterFileComboGroup>()
            };

            // Group clash zones by file combo
            var clashZonesByFileCombo = clashZones
                .GroupBy(cz =>
                {
                    string linkedFile = NormalizeFileNameForTree(cz.SourceDocKey);
                    string hostFile = NormalizeFileNameForTree(cz.HostDocKey ?? cz.StructuralElementDocumentTitle);
                    return (LinkedFile: linkedFile, HostFile: hostFile);
                })
                .Where(g => !string.IsNullOrWhiteSpace(g.Key.LinkedFile) && !string.IsNullOrWhiteSpace(g.Key.HostFile))
                .ToList();

            foreach (var fileComboGroup in clashZonesByFileCombo)
            {
                var (linkedFile, hostFile) = fileComboGroup.Key;
                var comboClashZones = fileComboGroup.ToList();

                var fileCombo = new Models.FilterFileComboGroup
                {
                    LinkedFile = linkedFile,
                    HostFile = hostFile,
                    ProcessedAt = DateTime.Now,
                    ClashZones = comboClashZones
                };
                filterGroup.FileCombos.Add(fileCombo);
            }

            if (filterGroup.FileCombos.Count > 0)
            {
                result.Add(filterGroup);
            }

            return result;
        }

        /// <summary>
        /// Helper to normalize file names (matches FilterFileComboGroup.GetNormalizedKey logic)
        /// </summary>
        private string NormalizeFileNameForTree(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName)) return string.Empty;

            var trimmed = fileName.Trim();
            trimmed = System.IO.Path.GetFileNameWithoutExtension(trimmed);
            var idxParen = trimmed.IndexOf('(');
            if (idxParen >= 0) trimmed = trimmed.Substring(0, idxParen).Trim();

            return trimmed;
        }

        private static bool DetectNewUiSelections(Document? document,
            OpeningFilter? filter,
            List<string> selectedReferenceFiles,
            List<string> selectedHostFiles,
            List<string> selectedMepCategories,
            List<string> selectedHostCategories)
        {
            // ✅ NEW-COMBO CHECK: Compare current UI combos against Global XML processed combos
            if (HasNewFileCombosInGlobal(document, selectedReferenceFiles, selectedHostFiles, selectedMepCategories))
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info("[DETECT-NEW-UI] Found UI file combo not present in Global XML processed list → detection required");
                }
                return true;
            }

            if (filter == null)
                return false;

            static bool HasNew(List<string>? current, List<string>? stored)
            {
                if (current == null || current.Count == 0)
                    return false;
                if (stored == null || stored.Count == 0)
                    return current.Count > 0;

                foreach (var item in current)
                {
                    var value = item?.Trim();
                    if (string.IsNullOrEmpty(value))
                        continue;

                    bool exists = stored.Any(existing => string.Equals(existing?.Trim(), value, StringComparison.OrdinalIgnoreCase));
                    if (!exists)
                        return true;
                }

                return false;
            }

            return HasNew(selectedReferenceFiles, filter.SelectedReferenceFiles)
                   || HasNew(selectedHostFiles, filter.SelectedHostFiles)
                   || HasNew(selectedMepCategories, filter.SelectedMepCategoryNames)
                   || HasNew(selectedHostCategories, filter.SelectedHostCategories);
        }

        private static bool HasNewFileCombosInGlobal(
            Document? document,
            List<string>? currentReferenceFiles,
            List<string>? currentHostFiles,
            List<string>? currentMepCategories)
        {
            if (document == null)
                return false;

            var requestedComboKeys = BuildNormalizedComboKeys(currentReferenceFiles, currentHostFiles);
            if (requestedComboKeys.Count == 0)
                return false;

            if (currentMepCategories == null || currentMepCategories.Count == 0)
                return false;

            foreach (var category in currentMepCategories)
            {
                if (string.IsNullOrWhiteSpace(category))
                    continue;

                try
                {
                    var processedKeys = GlobalIndexService.GetProcessedFileComboKeys(document, category) ?? Enumerable.Empty<string>();
                    var processedSet = new HashSet<string>(processedKeys, StringComparer.OrdinalIgnoreCase);

                    foreach (var comboKey in requestedComboKeys)
                    {
                        if (!processedSet.Contains(comboKey))
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[DETECT-NEW-UI] Combo '{comboKey}' for category '{category}' not found in Global XML");
                            }
                            return true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[DETECT-NEW-UI] Failed to read processed combos for category '{category}': {ex.Message}");
                    }
                    // Fail-safe: treat as new to force detection when combo status is uncertain
                    return true;
                }
            }

            return false;
        }

        private static HashSet<string> BuildNormalizedComboKeys(List<string>? referenceFiles, List<string>? hostFiles)
        {
            var comboKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (referenceFiles == null || hostFiles == null || referenceFiles.Count == 0 || hostFiles.Count == 0)
                return comboKeys;

            foreach (var reference in referenceFiles)
            {
                foreach (var host in hostFiles)
                {
                    var combo = new ProcessedFileCombo
                    {
                        LinkedFile = reference ?? string.Empty,
                        HostFile = host ?? string.Empty
                    };
                    comboKeys.Add(combo.GetNormalizedKey());
                }
            }

            return comboKeys;
        }
    }
}