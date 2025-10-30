using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

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

        // UI References (passed from main dialog)
        private System.Windows.Forms.Label _statusLabel;
        private System.Windows.Forms.ProgressBar _progressBar;
        private System.Windows.Forms.Button _refreshButton;

        public RefreshService(Document document, UIDocument uiDocument, ApplicationProfileService appProfileService)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _uiDocument = uiDocument ?? throw new ArgumentNullException(nameof(uiDocument));
            _appProfileService = appProfileService ?? throw new ArgumentNullException(nameof(appProfileService));
            _filterManagementService = new FilterManagementService(msg => DebugLogger.Info(msg), msg => DebugLogger.Error(msg));
            
            // ⚠️ CRITICAL FIX: Initialize with EMPTY storage in constructor ⚠️
            // The existing clash zones will be loaded during ExecuteRefresh from the current profile
            _clashZoneService = new ClashZoneService(new Models.ClashZoneStorage(), msg => DebugLogger.Info(msg));
            _intersectionService = new IntersectionDetectionService(msg => DebugLogger.Info(msg));
            
            // Initialize crash-safe executor for timeout protection
            _crashSafeExecutor = new CrashSafeExecutor();
            
            // ⚠️ CRITICAL: Initialize memory manager for large/unclean files (5 min timeout)
            // Memory limit is auto-calculated based on system RAM (10% of total RAM, capped at 12GB)
            // For 64GB system: ~6.4GB limit (10% of 64GB, capped at 12GB)
            _memoryManager = new MemoryManager(maxMemoryMB: null, timeoutMinutes: 5);
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
                var filtersDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Projects", "Default", "Filters");
                
                if (!Directory.Exists(filtersDirectory))
                {
                    DebugLogger.Info("[RefreshService] No existing filters directory found - starting with empty clash zones");
                    return;
                }

                // Look for existing XML files with clash zone data
                var xmlFiles = Directory.GetFiles(filtersDirectory, "*.xml");
                
                if (xmlFiles.Length == 0)
                {
                    DebugLogger.Info("[RefreshService] No existing XML files found - starting with empty clash zones");
                    return;
                }

                // Load clash zones from the most recently modified XML file
                var mostRecentFile = xmlFiles
                    .OrderByDescending(f => File.GetLastWriteTime(f))
                    .First();

                DebugLogger.Info($"[RefreshService] Loading existing clash zone data from: {mostRecentFile}");

                // Reuse a static serializer to avoid per-refresh type cache allocations
                var serializer = OpeningFilterSerializer;
                using (var reader = new StreamReader(mostRecentFile))
                {
                    var filter = (Models.OpeningFilter)serializer.Deserialize(reader);
                    
                    if (filter?.ClashZoneStorage?.ClashZones != null)
                    {
                        // ✅ CRITICAL: Preserve existing clash zone data including cluster information
                        _clashZoneService = new ClashZoneService(filter.ClashZoneStorage, msg => DebugLogger.Info(msg));
                        
                        int clusterResolvedCount = filter.ClashZoneStorage.ClashZones.Count(cz => cz.IsClusterResolved);
                        int individualResolvedCount = filter.ClashZoneStorage.ClashZones.Count(cz => cz.IsResolved);
                        
                        DebugLogger.Info($"[RefreshService] ✅ Loaded existing clash zone data: {filter.ClashZoneStorage.ClashZones.Count} total, {clusterResolvedCount} cluster-resolved, {individualResolvedCount} individual-resolved");
                    }
                    else
                    {
                        DebugLogger.Info("[RefreshService] No clash zone data found in existing XML file");
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[RefreshService] Error loading existing clash zone data: {ex.Message}");
                // Continue with empty storage if loading fails
            }
        }

        /// <summary>
        /// Ensure shared parameters are loaded into the project
        /// </summary>
        private void EnsureSharedParametersLoaded(Document doc)
        {
            try
            {
                DebugLogger.Info("[SHARED_PARAMS] Ensuring shared parameters are loaded into project");

                // Path to shared parameter file
                string sharedParamFile = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Resources\Opening family shared parameter.txt";

                if (!File.Exists(sharedParamFile))
                {
                    DebugLogger.Warning($"[SHARED_PARAMS] Shared parameter file not found: {sharedParamFile}");
                    return;
                }

                // Check if shared parameter file is already loaded
                var currentSharedParams = doc.Application.SharedParametersFilename;
                if (!string.IsNullOrEmpty(currentSharedParams) && currentSharedParams.Contains("Opening family shared parameter.txt"))
                {
                    DebugLogger.Info($"[SHARED_PARAMS] Shared parameters already loaded: {currentSharedParams}");
                    return;
                }

                // Load the shared parameter file
                doc.Application.SharedParametersFilename = sharedParamFile;
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
                        DebugLogger.Info($"[SHARED_PARAMS] Created shared parameter group: {groupName}");
                    }
                }

                DebugLogger.Info("[SHARED_PARAMS] Shared parameters loaded successfully");
            }
            catch (Exception ex)
            {
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
                DebugLogger.Info("[DUCT_ACCESSORIES] Starting damper parameter validation");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DUCT_ACCESSORIES] Starting damper parameter validation\n");

                if (selectedReferenceFiles == null || selectedReferenceFiles.Count == 0)
                {
                    var result = System.Windows.Forms.MessageBox.Show(
                        "No linked mechanical files selected. Please select a linked mechanical file to validate damper parameters.",
                        "No Linked Files Selected",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    
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
                    
                    DebugLogger.Warning("[DUCT_ACCESSORIES] No linked mechanical file found with duct accessories");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DUCT_ACCESSORIES] WARNING: No linked mechanical file found with duct accessories\n");
                    return false;
                }

                // Get all duct accessories (dampers) from the linked mechanical file
                var dampers = new FilteredElementCollector(linkedMechanicalDoc)
                    .OfCategory(BuiltInCategory.OST_DuctAccessory)
                    .WhereElementIsNotElementType()
                    .ToElements();

                DebugLogger.Info($"[DUCT_ACCESSORIES] Found {dampers.Count} dampers in linked mechanical file");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DUCT_ACCESSORIES] Found {dampers.Count} dampers in linked mechanical file\n");

                if (dampers.Count == 0)
                {
                    var result = System.Windows.Forms.MessageBox.Show(
                        "No dampers found in the linked mechanical file. Please ensure the mechanical file contains duct accessories (dampers).",
                        "No Dampers Found",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    
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
                        DebugLogger.Warning($"[DUCT_ACCESSORIES] {missingStandard.Count} dampers missing 'Standard' keyword in family name");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DUCT_ACCESSORIES] WARNING: {missingStandard.Count} dampers missing 'Standard' keyword in family name\n");
                    }
                    
                    if (missingMSFD.Count > 0)
                    {
                        message += $"• {missingMSFD.Count} dampers missing 'MSFD' keyword in family name\n";
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
                        DebugLogger.Info("[DUCT_ACCESSORIES] User chose to abort due to missing damper parameters");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DUCT_ACCESSORIES] User chose to abort due to missing damper parameters\n");
                        return false;
                    }
                    else
                    {
                        DebugLogger.Info("[DUCT_ACCESSORIES] User chose to continue despite missing damper parameters");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DUCT_ACCESSORIES] User chose to continue despite missing damper parameters\n");
                    }
                }
                else
                {
                    DebugLogger.Info("[DUCT_ACCESSORIES] All dampers have required Standard and MSFD parameters");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DUCT_ACCESSORIES] All dampers have required Standard and MSFD parameters\n");
                }

                return true;
            }
            catch (Exception ex)
            {
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
            // ✅ MEMORY OPTIMIZATION: Declare batched logger early for proper scope management
            BatchedLogger batchedLogger = null;
            // ✅ CRASH-SAFE: Create timestamped refresh log file using SafeFileLogger
            // SafeFileLogger automatically creates directories and handles missing paths gracefully
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            string refreshLogName = $"Refresh_{timestamp}.log";
            var __refreshStart = DateTime.Now;

            // IMMEDIATE CONSOLE LOGGING FOR VISIBILITY
            System.Diagnostics.Debug.WriteLine($"[REFRESH] === REFRESH METHOD STARTED AT {DateTime.Now} ===");
            System.Diagnostics.Debug.WriteLine($"[REFRESH] Timestamp: {timestamp}");
            System.Diagnostics.Debug.WriteLine($"[REFRESH] Log file will be: {refreshLogName}");

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
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] === REFRESH STARTED === Timestamp: {timestamp}\n");

            DebugLogger.Info("=== REFRESH METHOD STARTED ===");
            System.Diagnostics.Debug.WriteLine("[REFRESH] DebugLogger.Info called");
            
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
                System.Diagnostics.Debug.WriteLine($"[REFRESH] ✅ Test file written to: {testFile}");
            }
            catch (Exception testEx)
            {
                System.Diagnostics.Debug.WriteLine($"[REFRESH] ⚠️ Failed to write test file to AppData Logs: {testEx.Message}");
            }
            
            // Log the actual log directory being used (for troubleshooting)
            string actualLogDir = SafeFileLogger.GetLogDirectory();
            string logDirInfo = SafeFileLogger.GetLogDirectoryInfo();
            
            // Write directory info to refresh log file (user can open this file manually)
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] ═══ LOG DIRECTORY INFO ═══\n");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] {logDirInfo}\n");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] Memory profiling log: refresh_memory_profiling_{timestamp}.log\n");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] All logs will be written to: {actualLogDir}\n");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] ════════════════════════════════════\n");
            // Extra memory snapshot to pinpoint early allocations
            try { _memoryProfiler.TakeSnapshot("AFTER_LOG_INFO", 0); } catch { }
            
            // ✅ REMOVED: Log directory prompt (user knows where logs are now)
            // Log directory info is still written to refresh log file for reference
            
            // Show user-friendly message in status label
            _statusLabel.Text = "Starting refresh...";
            
            // ✅ MEMORY PROFILING: Initialize profiler to track actual memory usage per clash zone
            _memoryProfiler = new MemoryProfiler($"refresh_memory_profiling_{timestamp}.log");
            _memoryProfiler.TakeSnapshot("REFRESH_START", 0);

            // Step 1: Use passed filter selections from UI (optional - can work without filters)
            DebugLogger.Info($"[CLASH_DEBUG] Selected filter items: {string.Join(", ", selectedFilterItems)}");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Selected filter items: {string.Join(", ", selectedFilterItems)}\n");

            // ⚠️ CRITICAL: Require filter selection - don't proceed without a filter ⚠️
            if (selectedFilterItems.Count == 0)
            {
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

            // Step 2: Process filters and detect intersections
            var filtersToProcess = new List<Models.OpeningFilter>();
            
            // Try to load actual filters first
            foreach (var filterName in selectedFilterItems)
            {
                DebugLogger.Info($"[CLASH_DEBUG] Attempting to load filter: '{filterName}'");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Attempting to load filter: '{filterName}'\n");
                
                var filter = _filterManagementService.LoadFilterAuto(filterName);
                
                if (filter != null)
                {
                    DebugLogger.Info($"[CLASH_DEBUG] ✓ Successfully loaded filter: '{filterName}'");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ✓ Successfully loaded filter: '{filterName}'\n");
                    filtersToProcess.Add(filter);
                }
                else
                {
                    DebugLogger.Warning($"[CLASH_DEBUG] ✗ Failed to load filter: '{filterName}' - file may not exist or deserialization failed");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ✗ Failed to load filter: '{filterName}' - file may not exist or deserialization failed\n");
                }
            }
            
            // ✅ FIX: If no filters found, ERROR - don't auto-create fallback files
            if (filtersToProcess.Count == 0)
            {
                string errorMessage = "No filter found. Please create a filter first before running Refresh.";
                DebugLogger.Error($"[CLASH_DEBUG] {errorMessage}");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ERROR: {errorMessage}\n");
                
                TaskDialog.Show("Filter Required", 
                    "No filter found!\n\n" +
                    "Please create a filter in the Filter Management section before running Refresh.\n\n" +
                    "Filters define which MEP categories and reference files to process.");
                
                return Autodesk.Revit.UI.Result.Failed; // Stop refresh - don't proceed without a filter
            }

            // ⚠️ CRITICAL: Ensure shared parameters are loaded into the project before any parameter operations
            DebugLogger.Info("[SHARED_PARAMS] Ensuring shared parameters are loaded into project before refresh operations");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [SHARED_PARAMS] Ensuring shared parameters are loaded into project\n");
            EnsureSharedParametersLoaded(_document);

            DebugLogger.Info($"[CLASH_DEBUG] MEP filters to process: {filtersToProcess.Count}");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] MEP filters to process: {filtersToProcess.Count}\n");

            foreach (var filter in filtersToProcess)
            {
                DebugLogger.Info($"[CLASH_DEBUG] Filter: {filter.Name} - Category: {filter.Category} - Enabled: {filter.IsEnabled}");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Filter: {filter.Name} - Category: {filter.Category} - Enabled: {filter.IsEnabled}\n");
            }

            // Step 3: Get current profile and initialize clash zone service
            var currentProfile = _appProfileService.GetCurrentProfile();
            DebugLogger.Info($"[CLASH_DEBUG] Current profile: {currentProfile?.Name ?? "NULL"}");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Profile loaded: {currentProfile?.Name ?? "NULL"}\n");

            if (currentProfile?.Configuration?.ClashZoneStorage == null)
            {
                DebugLogger.Info("[CLASH_DEBUG] No clash zone storage in profile - proceeding with direct clash detection");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] INFO: No clash zone storage - proceeding with direct clash detection\n");
            }

            // Step 4: Check for active document
            if (_document == null)
            {
                DebugLogger.Error("[CLASH_DEBUG] No active Revit document found");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ERROR: No active Revit document found\n");
                
                _statusLabel.Text = "No active document";
                _progressBar.Visible = false;
                _refreshButton.Enabled = true;
                return Autodesk.Revit.UI.Result.Failed;
            }

            DebugLogger.Info($"[CLASH_DEBUG] Document: {_document.Title}");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Document: {_document.Title}\n");

            // Step 5: Check for Duct Accessories specific validation
            if (selectedMepCategories != null && selectedMepCategories.Count == 1 && selectedMepCategories.Contains("Duct Accessories"))
            {
                _statusLabel.Text = "Validating damper parameters...";
                _progressBar.Visible = true;
                _progressBar.Value = 10;

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

            // Step 6: Perform intersection detection
            _statusLabel.Text = "Detecting intersections...";
            _progressBar.Visible = true;
            _progressBar.Value = 20;

            DebugLogger.Info("[CLASH_DEBUG] Starting MEP-structural intersection detection...");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Starting MEP-structural intersection detection...\n");

            // Build allowed host types from current UI (UI PRECEDENCE)
            var allowedHostTypesUI = new HashSet<string>(
                FilterUiStateProvider.GetSelectedHostElementTypes?.Invoke() ?? new List<string>(),
                StringComparer.OrdinalIgnoreCase);
            DebugLogger.Info($"[CLASH_DEBUG] UI Selected MEP categories: {string.Join(", ", selectedMepCategories ?? new List<string>())}");
            DebugLogger.Info($"[CLASH_DEBUG] UI Selected Host types: {string.Join(", ", allowedHostTypesUI)}");
            DebugLogger.Info($"[CLASH_DEBUG] UI Selected Host files: {string.Join(", ", selectedHostFiles ?? new List<string>())}");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] UI Selected MEP categories: {string.Join(", ", selectedMepCategories ?? new List<string>())}\n");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] UI Selected Host types: {string.Join(", ", allowedHostTypesUI)}\n");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] UI Selected Host files: {string.Join(", ", selectedHostFiles ?? new List<string>())}\n");

            // Route intersection logs into both debug and refresh log files
            var intersectionService = new IntersectionDetectionService(msg =>
            {
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

            // Use passed MEP categories, reference files, and host filters from UI (UI PRECEDENCE)
            var currentIntersections = intersectionService.FindIntersections(
                _document,
                _document.ActiveView as View3D,
                selectedMepCategories,
                selectedReferenceFiles,
                selectedHostFiles,
                allowedHostTypesUI.ToList());

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

            DebugLogger.Info($"[CLASH_DEBUG] INTERSECTION DETECTION COMPLETE: {currentIntersections.Count} intersections found");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] INTERSECTION DETECTION COMPLETE: {currentIntersections.Count} intersections found\n");

            // Log intersection details to timestamped file
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

                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] Found {currentIntersections.Count} intersections during refresh\n");
            }
            else
            {
                DebugLogger.Warning("[CLASH_DEBUG] WARNING: No intersections found - clash detection returned empty results");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] WARNING: No intersections found - clash detection returned empty results\n");
            }

            // Step 6: Process clash zones
            _progressBar.Value = 40;
            _statusLabel.Text = "Processing clash zones...";

            // ✅ FIX: Load existing clash zones from Filter XML files (NOT profile)
            // This preserves IsResolved and IsClusterResolved flags from previous placement/clustering
            // Ensure per-project filters directory exists
            try { ProjectPathService.EnsureFiltersDirectory(_document); } catch { }
            // Validate and auto-heal global flags per selected category before proceeding
            var __globalsResolved = new HashSet<Guid>();
            try
            {
                foreach (var cat in selectedMepCategories ?? new List<string>())
                {
                    var (resolvedSet, clusterResolvedSet) = GlobalIndexService.ValidateAndFixFlags(_document, cat);
                    foreach (var gid in resolvedSet) __globalsResolved.Add(gid);
                    foreach (var gid in clusterResolvedSet) __globalsResolved.Add(gid);
                }
            }
            catch { }

            var existingClashZones = LoadExistingClashZonesFromFilterXml(selectedFilterItems, selectedMepCategories);
            __logPhaseMem("AFTER_XML_LOAD", existingClashZones?.ClashZones, null, null);
            var existingCount = existingClashZones?.ClashZones?.Count ?? 0;
            
            DebugLogger.Info($"[CLASH_DEBUG] Loaded {existingCount} existing clash zones from Filter XML files");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Loaded {existingCount} existing clash zones from Filter XML files (selected filter + categories)\n");
            
            // Reinitialize with existing clash zones
            // If a 3D view with section box is active, pre-filter existing zones to the oriented box
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
                    existingClashZones.ClashZones = kept;
                    DebugLogger.Info($"[CLASH_DEBUG] Section-box filtered existing zones: {kept.Count}");
                }
                catch (Exception ex)
                {
                    DebugLogger.Warning($"[CLASH_DEBUG] Section-box prefilter failed: {ex.Message}");
                }
            }

            // ✅ FIX: Recalculate intersection points for existing clash zones to use corrected intersection logic
            if (existingClashZones?.ClashZones != null && existingClashZones.ClashZones.Count > 0)
            {
                DebugLogger.Info($"[CLASH_DEBUG] ===== RECALCULATING INTERSECTION POINTS FOR {existingClashZones.ClashZones.Count} EXISTING CLASH ZONES =====");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ===== RECALCULATING INTERSECTION POINTS FOR {existingClashZones.ClashZones.Count} EXISTING CLASH ZONES =====\n");
                
                int updatedCount = 0;
                int skippedCount = 0;
                int errorCount = 0;
                
                // Recalculate intersection points for existing clash zones
                foreach (var existingZone in existingClashZones.ClashZones)
                {
                    try
                    {
                        // Get the MEP and structural elements
                        var mepElement = GetElementFromDocumentOrLinked(_document, existingZone.MepElementId);
                        var structuralElement = GetElementFromDocumentOrLinked(_document, existingZone.StructuralElementId);
                        
                        if (mepElement == null)
                        {
                            DebugLogger.Warning($"[CLASH_DEBUG] Zone {existingZone.Id}: MEP element {existingZone.MepElementId} not found");
                            skippedCount++;
                            continue;
                        }
                        
                        if (structuralElement == null)
                        {
                            DebugLogger.Warning($"[CLASH_DEBUG] Zone {existingZone.Id}: Structural element {existingZone.StructuralElementId} not found");
                            skippedCount++;
                            continue;
                        }
                        
                        // Recalculate intersection point using corrected logic
                        var newIntersectionPoint = CalculateIntersectionPoint(mepElement, structuralElement);
                        if (newIntersectionPoint != null)
                        {
                            // Update the intersection point if it differs significantly
                            var existingPoint = new XYZ(existingZone.IntersectionPointX, existingZone.IntersectionPointY, existingZone.IntersectionPointZ);
                            var distance = existingPoint.DistanceTo(newIntersectionPoint);
                            
                            if (distance > 0.001) // If points differ by more than 1mm
                            {
                                // Update intersection point (linked coordinates for individual sleeve placement)
                                existingZone.IntersectionPointX = newIntersectionPoint.X;
                                existingZone.IntersectionPointY = newIntersectionPoint.Y;
                                existingZone.IntersectionPointZ = newIntersectionPoint.Z;
                                
                                // ✅ CRITICAL: Transform to active document coordinates for proximity calculation
                                var activeDocPoint = TransformToActiveDocumentCoordinates(newIntersectionPoint, mepElement.Document);
                                existingZone.SleevePlacementPointActiveDocumentX = activeDocPoint.X;
                                existingZone.SleevePlacementPointActiveDocumentY = activeDocPoint.Y;
                                existingZone.SleevePlacementPointActiveDocumentZ = activeDocPoint.Z;
                                
                                DebugLogger.Info($"[CLASH_DEBUG] ✅ Zone {existingZone.Id}: Updated from {existingPoint} -> {newIntersectionPoint} (Δ={distance:F3}ft)");
                                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ✅ Zone {existingZone.Id}: Updated from ({existingPoint.X:F2}, {existingPoint.Y:F2}, {existingPoint.Z:F2}) -> ({newIntersectionPoint.X:F2}, {newIntersectionPoint.Y:F2}, {newIntersectionPoint.Z:F2}) (Δ={distance:F3}ft)\n");
                                updatedCount++;
                            }
                            else
                            {
                                DebugLogger.Info($"[CLASH_DEBUG] Zone {existingZone.Id}: No update needed (distance={distance:F6}ft)");
                                skippedCount++;
                            }
                        }
                        else
                        {
                            DebugLogger.Warning($"[CLASH_DEBUG] Zone {existingZone.Id}: Could not calculate new intersection point");
                            skippedCount++;
                        }

                        // ✅ ORIENTATION: Cache minimal data (no heavy API retention) and compute orientation post-API
                        try
                        {
                            // Compute wall/framing host orientation X/Y based on host normal
                            XYZ hostNormal = null;
                            if (structuralElement is Wall wall)
                            {
                                // Wall normal from location curve
                                var lc = wall.Location as LocationCurve;
                                if (lc != null)
                                {
                                    var dir = (lc.Curve.GetEndPoint(1) - lc.Curve.GetEndPoint(0)).Normalize();
                                    // Wall normal is perpendicular in XY plane
                                    hostNormal = new XYZ(-dir.Y, dir.X, 0).Normalize();
                                }
                            }
                            else if (structuralElement.Category?.Name?.IndexOf("Fram", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // Framing: fallback to X normal in XY plane (no unsupported built-in param)
                                hostNormal = new XYZ(1, 0, 0);
                            }

                            // MEP orientation vector (projection to XY)
                            XYZ mepDir = null;
                            if (mepElement.Location is LocationCurve mepLc)
                            {
                                mepDir = (mepLc.Curve.GetEndPoint(1) - mepLc.Curve.GetEndPoint(0)).Normalize();
                            }

                            // Cache numeric components (post-API usage only)
                            if (hostNormal != null)
                            {
                                existingZone.StructuralElementNormalX = hostNormal.X;
                                existingZone.StructuralElementNormalY = hostNormal.Y;
                                existingZone.StructuralElementNormalZ = hostNormal.Z;
                            }
                            if (mepDir != null)
                            {
                                existingZone.MepElementOrientationX = mepDir.X;
                                existingZone.MepElementOrientationY = mepDir.Y;
                                existingZone.MepElementOrientationZ = mepDir.Z;
                            }

                            // Derive HostOrientation string using cached values only
                            string hostOrientation = "Unknown";
                            var nx = existingZone.StructuralElementNormalX;
                            var ny = existingZone.StructuralElementNormalY;
                            if (Math.Abs(nx) >= Math.Abs(ny)) hostOrientation = "X"; else hostOrientation = "Y";
                            existingZone.HostOrientation = hostOrientation;
                        }
                        catch { /* non-fatal */ }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Warning($"[CLASH_DEBUG] ❌ Zone {existingZone.Id}: Error recalculating - {ex.Message}");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ❌ Zone {existingZone.Id}: Error - {ex.Message}\n");
                        errorCount++;
                    }
                }
                
                DebugLogger.Info($"[CLASH_DEBUG] Recalculation summary: Updated={updatedCount}, Skipped={skippedCount}, Errors={errorCount}");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Recalculation summary: Updated={updatedCount}, Skipped={skippedCount}, Errors={errorCount}\n");
                
                // ✅ MEMORY OPTIMIZATION: Clear Revit API objects after recalculation to free memory
                foreach (var existingZone in existingClashZones.ClashZones)
                {
                    existingZone.ClearRevitApiObjects();
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
            _clashZoneService = new ClashZoneService(existingClashZones, msg => batchedLogger.Log(msg));
            
            DebugLogger.Info($"[CLASH_DEBUG] ClashZoneService reinitialized with {existingCount} existing zones from XML");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ClashZoneService reinitialized with {existingCount} existing zones from XML\n");

            // Step 7: Filter and detect new clash zones
            _progressBar.Value = 50;
            _statusLabel.Text = "Detecting new clash zones...";

            // Use passed file selections and clearance settings from UI

            DebugLogger.Info("[CLASH_DEBUG] Skipping cleanup - no cleanup needed");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Skipping cleanup - no cleanup needed\n");

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
                DebugLogger.Info($"[CLASH_DEBUG] Selected MEP categories from UI: [{string.Join(", ", selectedMepCategories ?? new List<string>())}]");
                DebugLogger.Info($"[CLASH_DEBUG] Allowed MEP categories: [{string.Join(", ", allowedMepCats)}]");
                var allowedHostTypes = new HashSet<string>(
                    FilterUiStateProvider.GetSelectedHostElementTypes?.Invoke() ?? new List<string>(),
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
                    DebugLogger.Info($"[CLASH_DEBUG] BEFORE FILTERING: {before} clash zones, IsClusterResolved=True: {clusterResolvedBefore}, IsResolved=True: {individualResolvedBefore}");
                    
                    existingClashZones.ClashZones = existingClashZones.ClashZones.Where(cz =>
                    {
                        // Filter by MEP category
                        bool categoryMatch = allowedMepCats.Count == 0 || allowedMepCats.Contains(cz.MepElementCategory);
                        
                        // Filter by host type - Handle plural/singular mismatch: "Walls" (UI) vs "Wall" (Revit)
                        bool hostTypeMatch = allowedHostTypes.Count == 0 || 
                                           allowedHostTypes.Contains(cz.StructuralElementType) ||
                                           allowedHostTypes.Contains(cz.StructuralElementType + "s") ||
                                           allowedHostTypes.Any(t => t.TrimEnd('s').Equals(cz.StructuralElementType, StringComparison.OrdinalIgnoreCase));
                        
                        // 🔥 CRITICAL DEBUG: Log individual clash zone filtering
                        if (!categoryMatch)
                        {
                            DebugLogger.Info($"[CLASH_DEBUG] FILTERED OUT: ClashZone {cz.Id} - Category mismatch: '{cz.MepElementCategory}' not in [{string.Join(", ", allowedMepCats)}]");
                            // ✅ DEPLOYMENT MODE: Skip hardcoded log writes
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                            File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", 
                                $"[{DateTime.Now}] [CLASH_DEBUG] FILTERED OUT: ClashZone {cz.Id} - Category mismatch: '{cz.MepElementCategory}' not in [{string.Join(", ", allowedMepCats)}]\n");
                            }
                        }
                        if (!hostTypeMatch)
                        {
                            DebugLogger.Info($"[CLASH_DEBUG] FILTERED OUT: ClashZone {cz.Id} - Host type mismatch: '{cz.StructuralElementType}' not in [{string.Join(", ", allowedHostTypes)}]");
                            // ✅ DEPLOYMENT MODE: Skip hardcoded log writes
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                            File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", 
                                $"[{DateTime.Now}] [CLASH_DEBUG] FILTERED OUT: ClashZone {cz.Id} - Host type mismatch: '{cz.StructuralElementType}' not in [{string.Join(", ", allowedHostTypes)}]\n");
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
                            DebugLogger.Info($"[CLASH_DEBUG] FILTERED OUT: ClashZone {cz.Id} - Reference file mismatch: '{cz.SourceDocKey}' -> filename: '{refFileName}' -> normalized: '{norm(refFileName)}' not in [{string.Join(", ", allowedRefFiles)}]");
                            File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", 
                                $"[{DateTime.Now}] [CLASH_DEBUG] FILTERED OUT: ClashZone {cz.Id} - Reference file mismatch: '{refFileName}' not in [{string.Join(", ", allowedRefFiles)}]\n");
                        }
                        if (!hostFileMatch)
                        {
                            DebugLogger.Info($"[CLASH_DEBUG] FILTERED OUT: ClashZone {cz.Id} - Host file mismatch: '{cz.StructuralElementDocumentTitle}' -> filename: '{hostFileName}' -> normalized: '{norm(hostFileName)}' not in [{string.Join(", ", allowedHostFiles)}]");
                            File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", 
                                $"[{DateTime.Now}] [CLASH_DEBUG] FILTERED OUT: ClashZone {cz.Id} - Host file mismatch: '{hostFileName}' not in [{string.Join(", ", allowedHostFiles)}]\n");
                        }
                        
                        return categoryMatch && hostTypeMatch && refFileMatch && hostFileMatch;
                    }).ToList();
                    
                    DebugLogger.Info($"[CLASH_DEBUG] Current UI filter: MEP cats={allowedMepCats.Count}, Host types={allowedHostTypes.Count}, Ref files={allowedRefFiles.Count}, Host files={allowedHostFiles.Count}");
                    DebugLogger.Info($"[CLASH_DEBUG] Filtered existing zones by current UI: {before} -> {existingClashZones.ClashZones.Count}");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Filtered existing zones by current UI: {before} -> {existingClashZones.ClashZones.Count}\n");
                    __logPhaseMem("AFTER_UI_FILTER_EXISTING", existingClashZones?.ClashZones, null, null);
                    
                    // 🔥 CRITICAL DEBUG: Log flags AFTER filtering
                    int clusterResolvedAfter = existingClashZones.ClashZones.Count(cz => cz.IsClusterResolved);
                    int individualResolvedAfter = existingClashZones.ClashZones.Count(cz => cz.IsResolved);
                    DebugLogger.Info($"[CLASH_DEBUG] AFTER FILTERING: {existingClashZones.ClashZones.Count} clash zones, IsClusterResolved=True: {clusterResolvedAfter}, IsResolved=True: {individualResolvedAfter}");
                    DebugLogger.Info($"[CLASH_DEBUG] [{DateTime.Now}]  AFTER FILTERING: {existingClashZones.ClashZones.Count} clash zones, IsClusterResolved=True: {clusterResolvedAfter}, IsResolved=True: {individualResolvedAfter}\n");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[CLASH_DEBUG] Current-selection filter failed: {ex.Message}");
            }

            DebugLogger.Info($"[CLASH_DEBUG] Calling DetectNewClashZones with {currentIntersections.Count} intersections...");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Calling DetectNewClashZones with {currentIntersections.Count} intersections...\n");

            // DEBUG: Log intersection breakdown by category
            var intersectionBreakdown = currentIntersections.GroupBy(i => GetElementCategory(i.Item1))
                .ToDictionary(g => g.Key, g => g.Count());
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
            DebugLogger.Info($"[CLASH_DEBUG] ═══ BEFORE OPTIMIZATION: {ductWallIntersections.Count} Duct-Wall intersections found ═══");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ═══ BEFORE OPTIMIZATION: {ductWallIntersections.Count} Duct-Wall intersections (Total intersections: {currentIntersections.Count}) ═══\n");

            // ✅ CRITICAL FIX: Wrap DetectNewClashZones in try-catch to handle any exceptions
            List<Models.ClashZone> newClashZones = null;
            try
            {
                // ⚠️ CRITICAL: Check memory/timeout before heavy processing
                _memoryManager.CheckLimits();
                
                DebugLogger.Info($"[CLASH_DEBUG] About to call DetectNewClashZones...");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] About to call DetectNewClashZones...\n");
                
                // ⚠️ CRITICAL: Pass memory manager to clash zone service for timeout/memory protection
                _clashZoneService.SetMemoryManager(_memoryManager);
                
                // ✅ MEMORY PROFILING: Pass profiler to track memory during clash zone detection
                _memoryProfiler?.TakeSnapshot("BEFORE_DETECT_NEW_CLASH_ZONES", 0);
                _clashZoneService.SetMemoryProfiler(_memoryProfiler);
                
                newClashZones = _clashZoneService.DetectNewClashZones(currentIntersections, _document, clearanceSettings, selectedMepCategories);
                
                // ✅ MEMORY PROFILING: Record snapshot after clash zone detection
                _memoryProfiler?.TakeSnapshot("AFTER_DETECT_NEW_CLASH_ZONES", newClashZones?.Count ?? 0);
                
                // ✅ MEMORY OPTIMIZATION: Flush batched logs after clash zone detection completes
                if (batchedLogger != null)
                {
                    batchedLogger.Flush();
                }
                
            DebugLogger.Info($"[CLASH_DEBUG] ✅ DetectNewClashZones RETURNED: {newClashZones?.Count ?? 0} clash zones");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ✅ DetectNewClashZones RETURNED: {newClashZones?.Count ?? 0} clash zones\n");
                // Skip zones flagged resolved in per-category globals
                if (__globalsResolved != null && __globalsResolved.Count > 0 && newClashZones != null)
                {
                    int before = newClashZones.Count;
                    newClashZones = newClashZones.Where(cz => !__globalsResolved.Contains(cz.Id)).ToList();
                    int filtered = before - newClashZones.Count;
                    if (filtered > 0)
                    {
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
                DebugLogger.Error($"[CLASH_DEBUG] ❌ ERROR in DetectNewClashZones: {detectEx.Message}");
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
                    var mep = GetElementFromDocumentOrLinked(_document, cz.MepElementId);
                    var host = GetElementFromDocumentOrLinked(_document, cz.StructuralElementId);
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
                    if (newClashZones.Count % 100 == 0)
                    {
                        int totalParams = (mepBag?.Count ?? 0) + (hostBag?.Count ?? 0);
                        int paramBytes = totalParams * 150; // Rough estimate: 150 bytes per parameter
                        SafeFileLogger.SafeAppendText(refreshLogName, 
                            $"[{DateTime.Now}] [MEMORY_DEBUG] ClashZone {cz.Id}: MEP params={mepBag?.Count ?? 0}, Host params={hostBag?.Count ?? 0}, Total={totalParams}, Est. size={paramBytes} bytes\n");
                    }
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
                DebugLogger.Warning($"[PARAM_SNAPSHOT] Non-fatal: {ex.Message}");
            }

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
                    FilterUiStateProvider.GetSelectedHostElementTypes?.Invoke() ?? new List<string>(),
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
                    DebugLogger.Info($"[CLASH_DEBUG] UI host filter marked NEW zones eligible: {eligible}/{newClashZones.Count}");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] UI host filter marked NEW zones eligible: {eligible}/{newClashZones.Count}\n");
                }
            }
            catch (Exception uiFilterEx)
            {
                DebugLogger.Warning($"[CLASH_DEBUG] UI host filter failed: {uiFilterEx.Message}");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] UI host filter failed: {uiFilterEx.Message}\n");
            }
            
            DebugLogger.Info($"[CLASH_DEBUG] About to start Step 8: Save results");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] About to start Step 8: Save results\n");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] newClashZones count: {newClashZones?.Count ?? 0}\n");

            // Step 8: Save results
            _progressBar.Value = 70;
            _statusLabel.Text = "Saving results...";

            DebugLogger.Info($"[CLASH_DEBUG] Step 8 started - Progress bar set to 70");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Step 8 started - Progress bar set to 70\n");

            // Save to enabled filter
            var enabledFilter = filtersToProcess.FirstOrDefault(f => f.IsEnabled);
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Looking for enabled filter... filtersToProcess.Count={filtersToProcess.Count}\n");
            DebugLogger.Info($"[CLASH_DEBUG] Found {filtersToProcess.Count} filters to process");
            DebugLogger.Info($"[CLASH_DEBUG] Enabled filter: {(enabledFilter != null ? enabledFilter.Name : "NULL")}");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Found {filtersToProcess.Count} filters to process\n");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Enabled filter: {(enabledFilter != null ? enabledFilter.Name : "NULL")}\n");
            
            if (enabledFilter != null)
            {
                var targetFilter = enabledFilter;
                // ✅ FIX: Save ALL clash zones (existing + new), not just new ones
                // This preserves flags (IsResolved, IsClusterResolved) from previous runs
                var allClashZones = existingClashZones?.ClashZones ?? new List<Models.ClashZone>(); // Use existing loaded clash zones
                // Merge existing + new zones (avoid duplicates by Id) BEFORE saving, so snapshot bags persist to XML
                var mergedZones = new List<Models.ClashZone>();
                var existingById = new Dictionary<Guid, Models.ClashZone>();
                foreach (var z in allClashZones)
                {
                    if (!existingById.ContainsKey(z.Id))
                    {
                        existingById[z.Id] = z;
                        mergedZones.Add(z);
                    }
                }
                foreach (var nz in newClashZones ?? new List<Models.ClashZone>())
                {
                    if (!existingById.ContainsKey(nz.Id))
                    {
                        existingById[nz.Id] = nz;
                        mergedZones.Add(nz);
                    }
                    else
                    {
                        // Update existing entry with any newly computed snapshot fields
                        var ez = existingById[nz.Id];
                        ez.SourceDocKey = nz.SourceDocKey;
                        ez.HostDocKey = nz.HostDocKey;
                        ez.MepParameterValues = nz.MepParameterValues;
                        ez.HostParameterValues = nz.HostParameterValues;
                        // Overwrite critical geometry/thickness fields from the latest refresh
                        if (nz.StructuralElementThickness > 0)
                            ez.StructuralElementThickness = nz.StructuralElementThickness;
                        if (nz.StructuralElementNormal != null)
                            ez.StructuralElementNormal = nz.StructuralElementNormal;
                    }
                }

                DebugLogger.Info($"[PARAM_SNAPSHOT] Persisting {mergedZones.Count} zones (new={newClashZones?.Count ?? 0}, existing={allClashZones?.Count ?? 0}) with snapshot bags where available");
                __logPhaseMem("BEFORE_PERSIST", mergedZones, newClashZones, allClashZones);

                // ✅ CRITICAL DEBUG: Calculate statistics AFTER merge is complete
                var existingCountBeforeMerge = existingClashZones?.ClashZones?.Count ?? 0;
                var newZonesCount = newClashZones?.Count ?? 0;
                var mergedTotal = mergedZones.Count;
                var resolvedCount = mergedZones.Count(cz => cz.IsResolved);
                var unresolvedCount = mergedZones.Count(cz => !cz.IsResolved);
                
                DebugLogger.Info($"[CLASH_DEBUG] Merge Statistics - Existing: {existingCountBeforeMerge}, New: {newZonesCount}, Merged Total: {mergedTotal}, Resolved: {resolvedCount}, Unresolved: {unresolvedCount}");
                __logPhaseMem("AFTER_PERSIST_STATS", mergedZones, null, null);
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Merge Statistics - Existing: {existingCountBeforeMerge}, New: {newZonesCount}, Merged Total: {mergedTotal}, Resolved: {resolvedCount}, Unresolved: {unresolvedCount}\n");

                var clashZoneStorage = new Models.ClashZoneStorage
                {
                    ClashZones = mergedZones,
                    CreatedAt = DateTime.Now,
                    LastUpdated = DateTime.Now,
                    DocumentPath = _document.PathName,
                    DocumentHash = _document.PathName ?? "Unknown", // Use PathName as hash since GetDocumentHash doesn't exist
                    AlgorithmVersion = "1.0"
                };

                // CRITICAL FIX: Save clash zones to BOTH filter AND profile configuration
                targetFilter.ClashZoneStorage = clashZoneStorage;
                targetFilter.LastModified = DateTime.Now;
                // CRITICAL FIX: Set the selected MEP categories from UI selections
                targetFilter.SelectedMepCategoryNames = selectedMepCategories;
                // Persist reference and host files from UI selections
                targetFilter.SelectedReferenceFiles = selectedReferenceFiles;
                targetFilter.SelectedHostFiles = selectedHostFiles;
                // Also persist host element types selected in UI
                try { targetFilter.SelectedHostElementTypes = FilterUiStateProvider.GetSelectedHostElementTypes?.Invoke() ?? new List<string>(); } catch { }
                DebugLogger.Info($"[CLASH_DEBUG] Set targetFilter.SelectedMepCategoryNames to: {string.Join(", ", selectedMepCategories)}");
                DebugLogger.Info($"[CLASH_DEBUG] Set targetFilter.SelectedReferenceFiles to: {string.Join(", ", selectedReferenceFiles)}");
                DebugLogger.Info($"[CLASH_DEBUG] Set targetFilter.SelectedHostFiles to: {string.Join(", ", selectedHostFiles)}");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Set targetFilter.SelectedMepCategoryNames to: {string.Join(", ", selectedMepCategories)}\n");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Set targetFilter.SelectedReferenceFiles to: {string.Join(", ", selectedReferenceFiles)}\n");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Set targetFilter.SelectedHostFiles to: {string.Join(", ", selectedHostFiles)}\n");

                // Save to profile configuration for persistence
                if (currentProfile.Configuration == null)
                {
                    // Create a proper UserConfiguration object
                    currentProfile.Configuration = new Models.UserConfiguration(); // TODO: Initialize with proper values
                    DebugLogger.Info("[CLASH_DEBUG] Created new profile configuration during Refresh");
                }

                // ✅ CRITICAL FIX: Set clash zone storage in profile configuration
                // This ensures the OK button can find clash zones after refresh
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] About to set clash zone storage to profile configuration...\n");
                currentProfile.Configuration.ClashZoneStorage = clashZoneStorage;
                DebugLogger.Info($"[CLASH_DEBUG] Saved {clashZoneStorage.ClashZones.Count} clash zones to profile configuration");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Saved {clashZoneStorage.ClashZones.Count} clash zones to profile configuration\n");
                
                // Save the profile to persist changes
                try
                {
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] About to call SaveCurrentProfile()...\n");
                    _appProfileService.SaveCurrentProfile();
                    DebugLogger.Info("[CLASH_DEBUG] Profile saved successfully after setting clash zone storage");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Profile saved successfully\n");
                }
                catch (Exception saveEx)
                {
                    DebugLogger.Error($"[CLASH_DEBUG] Failed to save profile: {saveEx.Message}");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ERROR saving profile: {saveEx.Message}\n");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Profile save exception: {saveEx.StackTrace}\n");
                }

                // TODO: Persist current opening conditions (e.g., clearance values) into profile configuration
                // if (currentProfile.Configuration.OpeningSettings == null)
                // {
                //     currentProfile.Configuration.OpeningSettings = new Models.OpeningSettings();
                // }

                // Save filter to XML files
                try
                {
            var filterDir = ProjectPathService.GetFiltersDirectory(_document);
                    if (!Directory.Exists(filterDir))
                    {
                        Directory.CreateDirectory(filterDir);
                    }

                    // Save main filter file (for UI compatibility)
                    var mainFilePath = Path.Combine(filterDir, $"{targetFilter.Name}.xml");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] About to save main filter to: {mainFilePath}\n");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Filter has {targetFilter.ClashZoneStorage?.ClashZones?.Count ?? 0} clash zones\n");
                    _filterManagementService.SaveFilterToXmlFile(targetFilter, mainFilePath);
                    DebugLogger.Info($"[CLASH_DEBUG] Persisted main filter '{targetFilter.Name}' to XML file: {mainFilePath}");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Persisted main filter '{targetFilter.Name}' to XML file: {mainFilePath}\n");

                    // CRITICAL FIX: Save category-specific XML files RIGHT HERE after main filter is saved
                    DebugLogger.Info($"[CLASH_DEBUG] ===== SAVING CATEGORY-SPECIFIC XML FILES =====");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ===== SAVING CATEGORY-SPECIFIC XML FILES =====\n");

                    if (targetFilter.ClashZoneStorage != null && targetFilter.ClashZoneStorage.ClashZones != null && targetFilter.ClashZoneStorage.ClashZones.Count > 0)
                    {
                        var selectedCategories = targetFilter.SelectedMepCategoryNames ?? new List<string>();
                        DebugLogger.Info($"[CLASH_DEBUG] Saving category-specific XML files for: {string.Join(", ", selectedCategories)}");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Saving category-specific XML files for: {string.Join(", ", selectedCategories)}\n");

                        foreach (var category in selectedCategories)
                        {
                            DebugLogger.Info($"[CLASH_DEBUG] Processing category: '{category}' - About to call FilterClashZonesByCategory");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Processing category: '{category}' - About to call FilterClashZonesByCategory\n");
                            
                             var categoryClashZones = FilterClashZonesByCategory(targetFilter.ClashZoneStorage.ClashZones, category, _document, refreshLogName);
                            
                            DebugLogger.Info($"[CLASH_DEBUG] FilterClashZonesByCategory returned {categoryClashZones.Count} clash zones for category: '{category}'");
                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] FilterClashZonesByCategory returned {categoryClashZones.Count} clash zones for category: '{category}'\n");

                            // ✅ CRITICAL FIX: Always create category-specific XML files, even if empty
                            // This ensures consistency and helps with debugging
                                // Filter clearance settings for this specific category
                                var categoryClearances = FilterClearancesByCategory(targetFilter.OpeningSettings?.ClearanceSettings, category);
                                
                                var categoryFilter = new Models.OpeningFilter
                                {
                                    Name = $"{targetFilter.Name}_{category.ToLower().Replace(" ", "_")}",
                                    Category = targetFilter.Category,
                                    OpeningType = targetFilter.OpeningType,
                                    IsEnabled = true,
                                    LastModified = DateTime.Now,
                                    SelectedMepCategoryNames = new List<string> { category },
                                    SelectedReferenceFiles = targetFilter.SelectedReferenceFiles,
                                    SelectedHostFiles = targetFilter.SelectedHostFiles,
                                    OpeningSettings = new Models.OpeningSettings
                                    {
                                        ClearanceSettings = categoryClearances,
                                        SelectedMepType = GetCategorySpecificMepType(category)
                                    },
                                    ClashZoneStorage = new Models.ClashZoneStorage
                                    {
                                    ClashZones = categoryClashZones, // Can be empty
                                    CreatedAt = targetFilter.ClashZoneStorage?.CreatedAt ?? DateTime.Now,
                                        LastUpdated = DateTime.Now,
                                    DocumentPath = targetFilter.ClashZoneStorage?.DocumentPath ?? _document.PathName,
                                    DocumentHash = targetFilter.ClashZoneStorage?.DocumentHash ?? (_document.PathName ?? "Unknown"),
                                    AlgorithmVersion = targetFilter.ClashZoneStorage?.AlgorithmVersion ?? "1.0"
                                    }
                                };

                                // Save category-specific XML file
                                var categoryFilePath = Path.Combine(ProjectPathService.GetFiltersDirectory(_document), $"{categoryFilter.Name}.xml");
                                
                            if (categoryClashZones.Count > 0)
                            {
                                // 🔥 CRITICAL DEBUG: Log flags BEFORE saving to XML
                                int clusterResolvedBeforeSave = categoryClashZones.Count(cz => cz.IsClusterResolved);
                                int individualResolvedBeforeSave = categoryClashZones.Count(cz => cz.IsResolved);
                                DebugLogger.Info($"[CLASH_DEBUG] BEFORE SAVING TO XML: {categoryClashZones.Count} clash zones, IsClusterResolved=True: {clusterResolvedBeforeSave}, IsResolved=True: {individualResolvedBeforeSave}");
                                DebugLogger.Info($"[CLASH_DEBUG] [{DateTime.Now}]  BEFORE SAVING TO XML: {categoryClashZones.Count} clash zones, IsClusterResolved=True: {clusterResolvedBeforeSave}, IsResolved=True: {individualResolvedBeforeSave}\n");
                            }
                                
                                _filterManagementService.SaveFilterToXmlFile(categoryFilter, categoryFilePath);

                            if (categoryClashZones.Count > 0)
                            {
                                DebugLogger.Info($"[CLASH_DEBUG] Saved {categoryClashZones.Count} clash zones for category '{category}' to: {categoryFilePath}");
                                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Saved {categoryClashZones.Count} clash zones for category '{category}'\n");
                                if (OptimizationFlags.UseGlobalCategoryIndexForRefresh)
                                {
                                    // Write minimal per-category index for refresh-only dedupe/flags
                                    GlobalIndexService.EnsureEntries(_document, category, categoryClashZones.Select(cz => cz.Id));
                                }
                            }
                            else
                            {
                                DebugLogger.Info($"[CLASH_DEBUG] Created empty XML file for category '{category}' (no clash zones detected): {categoryFilePath}");
                                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Created empty XML file for category '{category}' (no clash zones detected)\n");
                            }
                        }
                    }
                    else
                    {
                        DebugLogger.Warning($"[CLASH_DEBUG] Cannot save category-specific XML - ClashZoneStorage is null or empty");
                    }
                }
                catch (Exception xmlSaveEx)
                {
                    DebugLogger.Error($"[CLASH_DEBUG] Failed to save filter to XML: {xmlSaveEx.Message}");
                    DebugLogger.Error($"[CLASH_DEBUG] XML Save Exception Stack Trace: {xmlSaveEx.StackTrace}");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] ERROR saving filter to XML: {xmlSaveEx.Message}\n");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] XML Save Exception Stack Trace: {xmlSaveEx.StackTrace}\n");
                }

                DebugLogger.Info($"[CLASH_DEBUG] Saved clash zones to filter '{targetFilter.Name}'");
                var savedCount = clashZoneStorage?.ClashZones?.Count ?? 0;
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] SUCCESS: Saved {savedCount} clash zones to filter '{targetFilter.Name}' and profile configuration\n");
            }
            else
            {
                DebugLogger.Warning("[CLASH_DEBUG] No enabled filter found to save clash zones");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] WARNING: No enabled filter found to save clash zones\n");
            }

            // Step 9: Update UI with results
            _progressBar.Value = 100;
            
            // ✅ CRITICAL FIX: Calculate final statistics from saved filter data
            int finalTotal = 0;
            int finalUnresolved = 0;
            int finalNewZones = 0;
            
            if (enabledFilter != null && enabledFilter.ClashZoneStorage != null && enabledFilter.ClashZoneStorage.ClashZones != null)
            {
                finalTotal = enabledFilter.ClashZoneStorage.ClashZones.Count;
                finalUnresolved = enabledFilter.ClashZoneStorage.ClashZones.Count(cz => !cz.IsResolved);
                finalNewZones = newClashZones?.Count ?? 0;
            }
            
            // Show results with log location
            string logDir = SafeFileLogger.GetLogDirectory();
            _statusLabel.Text = $"Clash zones: {finalTotal} total, {finalUnresolved} unresolved, {finalNewZones} new | Logs: {logDir}";

            DebugLogger.Info($"[CLASH_DEBUG] Clash zone refresh complete: {finalTotal} total, {finalUnresolved} unresolved, {finalNewZones} new");
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

            DebugLogger.Info("[REFRESH] Refresh process completed successfully");
            
            // ⚠️ CRITICAL: Force memory cleanup after refresh completes
            try
            {
                _memoryManager?.ForceCleanup();
                SafeFileLogger.SafeAppendText("refresh_memory.log", 
                    $"Refresh completed. Final memory: {_memoryManager?.CurrentMemoryMB ?? 0}MB, Elapsed: {_memoryManager?.ElapsedTime:mm\\:ss ?? TimeSpan.Zero:mm\\:ss}");
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
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] ✅ Memory profiling complete!\n");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] Memory profiling report: {memoryLogPath}\n");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] Open this file to see memory usage per clash zone vs theoretical estimates.\n");
                
                // Update status label with log location
                _statusLabel.Text = $"Refresh complete. Memory report: {memoryLogPath}";
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
            
            return Autodesk.Revit.UI.Result.Succeeded;
        }

        #region Helper Methods

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
                    DebugLogger.Info($"[CLEARANCE_DEBUG] No clearance settings to filter for category: {category}");
                    return categoryClearances;
                }

                DebugLogger.Info($"[CLEARANCE_DEBUG] Filtering {allClearances.Count} clearance settings for category: {category}");

                foreach (var kvp in allClearances)
                {
                    string key = kvp.Key;
                    double value = kvp.Value;
                    
                    // Check if this clearance key belongs to the specified category
                    if (IsClearanceKeyForCategory(key, category))
                    {
                        categoryClearances[key] = value;
                        DebugLogger.Info($"[CLEARANCE_DEBUG] Included clearance: {key} = {value}mm (for category: {category})");
                    }
                    else
                    {
                        DebugLogger.Info($"[CLEARANCE_DEBUG] Excluded clearance: {key} = {value}mm (not for category: {category})");
                    }
                }

                DebugLogger.Info($"[CLEARANCE_DEBUG] Filtered {categoryClearances.Count} clearance settings for category '{category}'");
            }
            catch (Exception ex)
            {
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
                DebugLogger.Info($"[CLASH_DEBUG] Filtering {clashZones.Count} clash zones for category: {category}");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Filtering {clashZones.Count} clash zones for category: {category}\n");
                
                // DEBUG: Show first few clash zones to understand the data
                for (int i = 0; i < Math.Min(5, clashZones.Count); i++)
                {
                    var cz = clashZones[i];
                    if (OptimizationFlags.UseDiagnosticMode)
                    {
                        DebugLogger.Info($"[CLASH_DEBUG] Sample zone {i}: MEP ID={cz.MepElementId}, Cat='{cz.MepElementCategory}', Structural={cz.StructuralElementType}");
                        SafeFileLogger.SafeAppendText(refreshLogName,
                            $"[{DateTime.Now}] [CLASH_DEBUG] Sample zone {i}: MEP ID={cz.MepElementId}, Cat='{cz.MepElementCategory}', Structural={cz.StructuralElementType}\n");
                    }
                }

                // Count zones by structural type for debugging
                var byStructType = clashZones.GroupBy(cz => cz.StructuralElementType).ToDictionary(g => g.Key, g => g.Count());
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

                        DebugLogger.Info($"[CLASH_DEBUG] Zone {clashZone.Id}: Cat='{elementCategory}', Struct='{structType}', Orient='{wallOrientation}', Checking against '{category}'");
                        SafeFileLogger.SafeAppendText(refreshLogName, 
                            $"[{DateTime.Now}] [CLASH_DEBUG] Zone {clashZone.Id}: Cat='{elementCategory}', Struct='{structType}', Orient='{wallOrientation}', Checking against '{category}'\n");

                        // ✅ FIX: Enhanced category matching with fallback for linked files
                        bool isMatch = IsElementInCategoryEnhanced(elementCategory, category, clashZone, document);

                        if (isMatch)
                        {
                            filteredZones.Add(clashZone);
                            DebugLogger.Info($"[CLASH_DEBUG] ✅ Zone {clashZone.Id} MATCHED '{category}' - Cat='{elementCategory}', Orient='{wallOrientation}'");
                            SafeFileLogger.SafeAppendText(refreshLogName, 
                                $"[{DateTime.Now}] [CLASH_DEBUG] ✅ Zone {clashZone.Id} MATCHED '{category}' - Cat='{elementCategory}', Orient='{wallOrientation}'\n");
                        }
                        else
                        {
                            DebugLogger.Info($"[CLASH_DEBUG] ❌ Zone {clashZone.Id} NOT matched '{category}' - Cat='{elementCategory}'");
                            SafeFileLogger.SafeAppendText(refreshLogName, 
                                $"[{DateTime.Now}] [CLASH_DEBUG] ❌ Zone {clashZone.Id} NOT matched '{category}' - Cat='{elementCategory}'\n");
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Error($"[CLASH_DEBUG] Error filtering clash zone {clashZone.Id}: {ex.Message}");
                    }
                }

                DebugLogger.Info($"[CLASH_DEBUG] ✅ Filtered {filteredZones.Count}/{clashZones.Count} clash zones for category '{category}'");
                SafeFileLogger.SafeAppendText(refreshLogName, 
                    $"[{DateTime.Now}] [CLASH_DEBUG] ✅ Filtered {filteredZones.Count}/{clashZones.Count} clash zones for category '{category}'\n");
            }
            catch (Exception ex)
            {
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
                        DebugLogger.Info($"[CATEGORY_MATCH] ✅ Cable tray enhanced match for '{elementCategory}'");
                        return true;
                    }

                    // ✅ PHASE 3: Fallback - check actual element from Revit (for linked files)
                    try
                    {
                        var mepElement = GetElementFromDocumentOrLinked(document, clashZone.MepElementId);
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
                                DebugLogger.Info($"[CATEGORY_MATCH] ✅ Cable tray FALLBACK match from actual element: '{actualCategory}'");
                                // Update cached category for future use
                                clashZone.MepElementCategory = actualCategory;
                                return true;
                            }
                            else
                            {
                                DebugLogger.Info($"[CATEGORY_MATCH] ❌ Cable tray fallback check failed: actual category = '{actualCategory}'");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
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
                        DebugLogger.Info($"[CATEGORY_MATCH] ✅ Duct accessory enhanced match for '{elementCategory}'");
                        return true;
                    }
                }

                DebugLogger.Info($"[CATEGORY_MATCH] ❌ No match found for '{elementCategory}' -> '{requestedCategory}'");
                return false;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[CATEGORY_MATCH] Error in IsElementInCategoryEnhanced: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Gets element from document or linked documents
        /// </summary>
        private Element GetElementFromDocumentOrLinked(Document document, ElementId elementId)
        {
            try
            {
                DebugLogger.Info($"[CLASH_DEBUG] Looking for element {elementId} in document '{document.Title}'");
                
                // First try host document
                var element = document.GetElement(elementId);
                if (element != null)
                {
                    DebugLogger.Info($"[CLASH_DEBUG] Found element {elementId} in host document - Category: {element.Category?.Name ?? "Unknown"}");
                    return element;
                }

                DebugLogger.Info($"[CLASH_DEBUG] Element {elementId} not found in host document, searching linked documents...");

                // If not found, search linked documents
                var linkInstances = new FilteredElementCollector(document)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>();

                DebugLogger.Info($"[CLASH_DEBUG] Found {linkInstances.Count()} linked documents to search");

                foreach (var link in linkInstances)
                {
                    var linkDoc = link.GetLinkDocument();
                    if (linkDoc != null)
                    {
                        DebugLogger.Info($"[CLASH_DEBUG] Searching in linked document: {linkDoc.Title}");
                        element = linkDoc.GetElement(elementId);
                        if (element != null)
                        {
                            DebugLogger.Info($"[CLASH_DEBUG] Found element {elementId} in linked document '{linkDoc.Title}' - Category: {element.Category?.Name ?? "Unknown"}");
                            return element;
                        }
                    }
                }

                DebugLogger.Warning($"[CLASH_DEBUG] Element {elementId} not found in host document or any linked documents");
                return null;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[CLASH_DEBUG] Error getting element {elementId}: {ex.Message}");
                return null;
            }
        }

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
                DebugLogger.Info($"[CLASH_DEBUG] Triggering fallback category detection for element {elementId} with standard category '{standardCategory}'");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Triggering fallback category detection for element {elementId} with standard category '{standardCategory}'\n");
                
                var fallbackCategory = GetCategoryFromElementType(element);
                
                DebugLogger.Info($"[CLASH_DEBUG] Fallback category detection for element {elementId}: '{fallbackCategory}' (original: '{standardCategory}')");
                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Fallback category detection for element {elementId}: '{fallbackCategory}' (original: '{standardCategory}')\n");

                // If fallback also fails, try to infer from intersection detection results
                if (fallbackCategory == "Unknown")
                {
                    // Since we know these elements were detected as pipes during intersection detection,
                    // and they're being categorized as "Automatic Sketch Dimensions", they're likely pipes
                    if (standardCategory.Contains("Sketch"))
                    {
                        DebugLogger.Info($"[CLASH_DEBUG] Inferring 'Pipes' category for sketch element {elementId} based on intersection detection");
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [CLASH_DEBUG] Inferring 'Pipes' category for sketch element {elementId} based on intersection detection\n");
                        return "Pipes";
                    }
                }

                return fallbackCategory;
            }
            catch (Exception ex)
            {
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

                DebugLogger.Info($"[CLASH_DEBUG] Element type analysis: TypeName='{typeName}'");

                // Check for specific MEP element types
                if (element is FamilyInstance familyInstance)
                {
                    var familyName = familyInstance.Symbol?.FamilyName ?? "Unknown";
                    var categoryName = familyInstance.Category?.Name ?? "Unknown";
                    
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
                        DebugLogger.Warning($"[CLASH_DEBUG] Unknown requested category: '{requestedCategory}'");
                        return false;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[CLASH_DEBUG] Error checking category match: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Loads the most recently saved OpeningFilter XML from any project Filters folder under AppData.
        /// Used to read saved host element types when filtering existing zones.
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
        /// </summary>
        private Models.ClashZoneStorage LoadExistingClashZonesFromFilterXml(List<string> selectedFilterNames, List<string> selectedCategories)
        {
            if (OptimizationFlags.UseGlobalCategoryIndexForRefresh)
            {
                DebugLogger.Info("[CLASH_DEBUG] Using per-category global index; skipping full XML load");
                return new Models.ClashZoneStorage { ClashZones = new List<Models.ClashZone>(), LastUpdated = DateTime.Now };
            }
            var mergedStorage = new Models.ClashZoneStorage
            {
                ClashZones = new List<Models.ClashZone>(),
                LastUpdated = DateTime.Now
            };

            try
            {
            var filtersDirectory = ProjectPathService.GetFiltersDirectory(_document);
                
                DebugLogger.Info($"[LoadExistingClashZones] Looking for XML files in: {filtersDirectory}");
                DebugLogger.Info($"[CLASH_DEBUG] [{DateTime.Now}] [LoadExistingClashZones] Looking for XML files in: {filtersDirectory}\n");
                
                if (!Directory.Exists(filtersDirectory))
                {
                    DebugLogger.Warning($"[LoadExistingClashZones] Filters directory does not exist!");
                    return mergedStorage;
                }

                // 🔥 DEBUG: List ALL XML files in the directory
                var allXmlFiles = Directory.GetFiles(filtersDirectory, "*.xml");
                DebugLogger.Info($"[LoadExistingClashZones] Found {allXmlFiles.Length} total XML files");
                DebugLogger.Info($"[CLASH_DEBUG] [{DateTime.Now}] [LoadExistingClashZones] Found {allXmlFiles.Length} total XML files\n");
                
                foreach (var xmlFile in allXmlFiles)
                {
                    DebugLogger.Info($"[LoadExistingClashZones] - {Path.GetFileName(xmlFile)}");
                    DebugLogger.Info($"[CLASH_DEBUG] [{DateTime.Now}] [LoadExistingClashZones] - {Path.GetFileName(xmlFile)}\n");
                }

                // Load clash zones from each category-specific XML file
                foreach (var filterName in selectedFilterNames)
                {
                    DebugLogger.Info($"[LoadExistingClashZones] Processing filter: {filterName}");
                    DebugLogger.Info($"[CLASH_DEBUG] [{DateTime.Now}] [LoadExistingClashZones] Processing filter: {filterName}\n");
                    
                    foreach (var category in selectedCategories)
                    {
                        var pattern = $"{filterName}_{category.ToLower().Replace(" ", "_")}.xml";
                        DebugLogger.Info($"[LoadExistingClashZones] Looking for pattern: {pattern}");
                        DebugLogger.Info($"[CLASH_DEBUG] [{DateTime.Now}] [LoadExistingClashZones] Looking for pattern: {pattern}\n");
                        
                        var matchingFiles = Directory.GetFiles(filtersDirectory, pattern);
                        
                        DebugLogger.Info($"[LoadExistingClashZones] Found {matchingFiles.Length} matching files for pattern: {pattern}");
                        DebugLogger.Info($"[CLASH_DEBUG] [{DateTime.Now}] [LoadExistingClashZones] Found {matchingFiles.Length} matching files for pattern: {pattern}\n");
                        
                        if (matchingFiles.Length > 0)
                        {
                            var xmlFile = matchingFiles.First();
                            DebugLogger.Info($"[LoadExistingClashZones] ✅ Loading from {Path.GetFileName(xmlFile)}");
                            
                            try
                            {
                                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(Models.OpeningFilter));
                                using (var reader = new StreamReader(xmlFile))
                                {
                                    var filter = (Models.OpeningFilter)serializer.Deserialize(reader);
                                    if (filter?.ClashZoneStorage?.ClashZones != null)
                                    {
                                        mergedStorage.ClashZones.AddRange(filter.ClashZoneStorage.ClashZones);
                                        DebugLogger.Info($"[LoadExistingClashZones] ✅ Loaded {filter.ClashZoneStorage.ClashZones.Count} clash zones from {Path.GetFileName(xmlFile)}");
                                        DebugLogger.Info($"[CLASH_DEBUG] [{DateTime.Now}] [LoadExistingClashZones] ✅ Loaded {filter.ClashZoneStorage.ClashZones.Count} clash zones\n");
                                    }
                                    else
                                    {
                                        DebugLogger.Warning($"[LoadExistingClashZones] ❌ No clash zones in {Path.GetFileName(xmlFile)}");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                DebugLogger.Error($"[LoadExistingClashZones] ❌ Error loading {Path.GetFileName(xmlFile)}: {ex.Message}");
                                DebugLogger.Info($"[CLASH_DEBUG] [{DateTime.Now}] [LoadExistingClashZones] ❌ ERROR: {ex.Message}\n");
                            }
                        }
                    }
                }
                
                DebugLogger.Info($"[LoadExistingClashZones] ✅ Total clash zones loaded: {mergedStorage.ClashZones.Count}");
                DebugLogger.Info($"[CLASH_DEBUG] [{DateTime.Now}] [LoadExistingClashZones] ✅ FINAL: Total clash zones loaded: {mergedStorage.ClashZones.Count}\n");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[LoadExistingClashZones] Error: {ex.Message}");
                DebugLogger.Info($"[CLASH_DEBUG] [{DateTime.Now}] [LoadExistingClashZones] EXCEPTION: {ex.Message}\n");
            }

            return mergedStorage;
        }

        #endregion

        /// <summary>
        /// Helper method to calculate intersection point between MEP and structural elements
        /// </summary>
        private XYZ CalculateIntersectionPoint(Element mepElement, Element structuralElement)
        {
            try
            {
                // Get geometry from both elements
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
                DebugLogger.Warning($"[CLASH_DEBUG] Failed to calculate intersection point: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// Transform coordinates from linked document to active document
        /// </summary>
        private XYZ TransformToActiveDocumentCoordinates(XYZ point, Document linkedDocument)
        {
            try
            {
                // For now, return the point as-is since coordinate transformation is complex
                // TODO: Implement proper coordinate transformation if needed
                return point;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[RefreshService] Error transforming coordinates: {ex.Message}");
                return point; // Return original point as fallback
            }
        }
    }
}

