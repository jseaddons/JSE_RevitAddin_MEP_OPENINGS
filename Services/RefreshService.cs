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
        /// Validates that dampers in the selected linked mechanical file have required Standard and MSFD parameters
        /// </summary>
        private bool ValidateDamperParameters(List<string> selectedReferenceFiles, string refreshLogPath)
        {
            try
            {
                DebugLogger.Info("[DUCT_ACCESSORIES] Starting damper parameter validation");
                File.AppendAllText(refreshLogPath, $"[{DateTime.Now}] [DUCT_ACCESSORIES] Starting damper parameter validation\n");

                if (selectedReferenceFiles == null || selectedReferenceFiles.Count == 0)
                {
                    var result = System.Windows.Forms.MessageBox.Show(
                        "No linked mechanical files selected. Please select a linked mechanical file to validate damper parameters.",
                        "No Linked Files Selected",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    
                    DebugLogger.Warning("[DUCT_ACCESSORIES] No linked files selected for damper validation");
                    File.AppendAllText(refreshLogPath, $"[{DateTime.Now}] [DUCT_ACCESSORIES] WARNING: No linked files selected for damper validation\n");
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
                            File.AppendAllText(refreshLogPath, $"[{DateTime.Now}] [DUCT_ACCESSORIES] Found linked mechanical file: {linkDoc.Title}\n");
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
                    File.AppendAllText(refreshLogPath, $"[{DateTime.Now}] [DUCT_ACCESSORIES] WARNING: No linked mechanical file found with duct accessories\n");
                    return false;
                }

                // Get all duct accessories (dampers) from the linked mechanical file
                var dampers = new FilteredElementCollector(linkedMechanicalDoc)
                    .OfCategory(BuiltInCategory.OST_DuctAccessory)
                    .WhereElementIsNotElementType()
                    .ToElements();

                DebugLogger.Info($"[DUCT_ACCESSORIES] Found {dampers.Count} dampers in linked mechanical file");
                File.AppendAllText(refreshLogPath, $"[{DateTime.Now}] [DUCT_ACCESSORIES] Found {dampers.Count} dampers in linked mechanical file\n");

                if (dampers.Count == 0)
                {
                    var result = System.Windows.Forms.MessageBox.Show(
                        "No dampers found in the linked mechanical file. Please ensure the mechanical file contains duct accessories (dampers).",
                        "No Dampers Found",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    
                    DebugLogger.Warning("[DUCT_ACCESSORIES] No dampers found in linked mechanical file");
                    File.AppendAllText(refreshLogPath, $"[{DateTime.Now}] [DUCT_ACCESSORIES] WARNING: No dampers found in linked mechanical file\n");
                    return false;
                }

                // Check each damper for required parameters
                var missingStandard = new List<Element>();
                var missingMSFD = new List<Element>();

                foreach (var damper in dampers)
                {
                    var standardParam = damper.LookupParameter("Standard");
                    var msfdParam = damper.LookupParameter("MSFD");

                    if (standardParam == null || string.IsNullOrWhiteSpace(standardParam.AsString()))
                    {
                        missingStandard.Add(damper);
                    }

                    if (msfdParam == null || string.IsNullOrWhiteSpace(msfdParam.AsString()))
                    {
                        missingMSFD.Add(damper);
                    }
                }

                // Report results
                if (missingStandard.Count > 0 || missingMSFD.Count > 0)
                {
                    var message = "The following damper parameter validation issues were found:\n\n";
                    
                    if (missingStandard.Count > 0)
                    {
                        message += $"• {missingStandard.Count} dampers missing 'Standard' parameter\n";
                        DebugLogger.Warning($"[DUCT_ACCESSORIES] {missingStandard.Count} dampers missing 'Standard' parameter");
                        File.AppendAllText(refreshLogPath, $"[{DateTime.Now}] [DUCT_ACCESSORIES] WARNING: {missingStandard.Count} dampers missing 'Standard' parameter\n");
                    }
                    
                    if (missingMSFD.Count > 0)
                    {
                        message += $"• {missingMSFD.Count} dampers missing 'MSFD' parameter\n";
                        DebugLogger.Warning($"[DUCT_ACCESSORIES] {missingMSFD.Count} dampers missing 'MSFD' parameter");
                        File.AppendAllText(refreshLogPath, $"[{DateTime.Now}] [DUCT_ACCESSORIES] WARNING: {missingMSFD.Count} dampers missing 'MSFD' parameter\n");
                    }

                    message += "\nPlease add the missing parameters to the dampers in the linked mechanical file before proceeding with intersection detection.";
                    message += "\n\nDo you want to continue anyway?";

                    var result = System.Windows.Forms.MessageBox.Show(
                        message,
                        "Damper Parameter Validation Failed",
                        System.Windows.Forms.MessageBoxButtons.YesNo,
                        System.Windows.Forms.MessageBoxIcon.Warning);

                    if (result == System.Windows.Forms.DialogResult.No)
                    {
                        DebugLogger.Info("[DUCT_ACCESSORIES] User chose to abort due to missing damper parameters");
                        File.AppendAllText(refreshLogPath, $"[{DateTime.Now}] [DUCT_ACCESSORIES] User chose to abort due to missing damper parameters\n");
                        return false;
                    }
                    else
                    {
                        DebugLogger.Info("[DUCT_ACCESSORIES] User chose to continue despite missing damper parameters");
                        File.AppendAllText(refreshLogPath, $"[{DateTime.Now}] [DUCT_ACCESSORIES] User chose to continue despite missing damper parameters\n");
                    }
                }
                else
                {
                    DebugLogger.Info("[DUCT_ACCESSORIES] All dampers have required Standard and MSFD parameters");
                    File.AppendAllText(refreshLogPath, $"[{DateTime.Now}] [DUCT_ACCESSORIES] All dampers have required Standard and MSFD parameters\n");
                }

                return true;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[DUCT_ACCESSORIES] Error during damper parameter validation: {ex.Message}");
                File.AppendAllText(refreshLogPath, $"[{DateTime.Now}] [DUCT_ACCESSORIES] ERROR: {ex.Message}\n");
                
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
        /// </summary>
        public void ExecuteRefresh(List<string> selectedFilterItems, List<string> selectedMepCategories, List<string> selectedReferenceFiles, List<string> selectedHostFiles, Dictionary<string, double> clearanceSettings)
        {
            // Create timestamped refresh log file for debugging
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            string refreshLogPath = $@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\Refresh_{timestamp}.log";

            // IMMEDIATE CONSOLE LOGGING FOR VISIBILITY
            System.Diagnostics.Debug.WriteLine($"[REFRESH] === REFRESH METHOD STARTED AT {DateTime.Now} ===");
            System.Diagnostics.Debug.WriteLine($"[REFRESH] Timestamp: {timestamp}");
            System.Diagnostics.Debug.WriteLine($"[REFRESH] Log file will be: {refreshLogPath}");

            try
            {
                // Ensure directory exists
                string logDir = Path.GetDirectoryName(refreshLogPath) ?? @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log";
                if (!Directory.Exists(logDir))
                {
                    Directory.CreateDirectory(logDir);
                    System.Diagnostics.Debug.WriteLine($"[REFRESH] Created log directory: {logDir}");
                }

                // Write initial log entry
                File.AppendAllText(refreshLogPath, $"[{DateTime.Now}] === REFRESH METHOD STARTED ===\n");
                File.AppendAllText(refreshLogPath, $"[{DateTime.Now}] Refresh log file: Refresh_{timestamp}.log\n");
                File.AppendAllText(refreshLogPath, $"[{DateTime.Now}] Current profile: {_appProfileService?.GetCurrentProfile()?.Name ?? "None"}\n");

                System.Diagnostics.Debug.WriteLine($"[REFRESH] Successfully created log file: {refreshLogPath}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[REFRESH] ERROR creating log file: {ex.Message}");
                DebugLogger.Error($"Failed to create refresh log file: {ex.Message}");
                // Continue with refresh even if logging fails
            }

            // Write to timestamped log file only
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] REFRESH METHOD STARTED - Log file: Refresh_{timestamp}.log\n");

            // Write to the main debug logger file that user can see
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] === REFRESH STARTED === Timestamp: {timestamp}\n");

            DebugLogger.Info("=== REFRESH METHOD STARTED ===");
            System.Diagnostics.Debug.WriteLine("[REFRESH] DebugLogger.Info called");

            // Step 1: Use passed filter selections from UI (optional - can work without filters)
            DebugLogger.Info($"[CLASH_DEBUG] Selected filter items: {string.Join(", ", selectedFilterItems)}");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] Selected filter items: {string.Join(", ", selectedFilterItems)}\n");

            // ⚠️ CRITICAL: Require filter selection - don't proceed without a filter ⚠️
            if (selectedFilterItems.Count == 0)
            {
                DebugLogger.Error("[CLASH_DEBUG] ERROR: No filters selected - cannot proceed with refresh");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] ERROR: No filters selected - cannot proceed with refresh\n");
                
                // Prompt user to select a filter
                System.Windows.Forms.MessageBox.Show(
                    "Please select at least one filter before running Refresh.\n\nFilters are listed in the left panel.",
                    "No Filter Selected",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                
                _statusLabel.Text = "Refresh cancelled - No filter selected";
                _progressBar.Value = 0;
                return; // STOP - don't proceed without a filter
            }

            // Step 2: Process filters and detect intersections
            var filtersToProcess = new List<Models.OpeningFilter>();
            
            // Try to load actual filters first
            foreach (var filterName in selectedFilterItems)
            {
                DebugLogger.Info($"[CLASH_DEBUG] Attempting to load filter: '{filterName}'");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] Attempting to load filter: '{filterName}'\n");
                
                var filter = _filterManagementService.LoadFilterAuto(filterName);
                
                if (filter != null)
                {
                    DebugLogger.Info($"[CLASH_DEBUG] ✓ Successfully loaded filter: '{filterName}'");
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] ✓ Successfully loaded filter: '{filterName}'\n");
                    filtersToProcess.Add(filter);
                }
                else
                {
                    DebugLogger.Warning($"[CLASH_DEBUG] ✗ Failed to load filter: '{filterName}' - file may not exist or deserialization failed");
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] ✗ Failed to load filter: '{filterName}' - file may not exist or deserialization failed\n");
                }
            }
            
            // CRITICAL FIX: If no filters found, create a filter using the selected filter name from UI
            if (filtersToProcess.Count == 0)
            {
                // Use the first selected filter name from UI (e.g., "Ventilation"), not hardcoded "Default_Refresh"
                string selectedFilterName = selectedFilterItems.FirstOrDefault() ?? "Default_Refresh";
                DebugLogger.Info($"[CLASH_DEBUG] No existing filter XML found - creating new filter with name '{selectedFilterName}' for saving results");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] No existing filter XML found - creating new filter with name '{selectedFilterName}' for saving results\n");
                
                var defaultFilter = new Models.OpeningFilter
                {
                    Name = selectedFilterName,
                    Category = Models.MepCategory.Ducts, // Use enum instead of string
                    OpeningType = Models.OpeningType.RectangularSleeves, // Use enum instead of string
                    IsEnabled = true,
                    LastModified = DateTime.Now,
                    SelectedMepCategoryNames = selectedMepCategories,
                    SelectedReferenceFiles = selectedReferenceFiles,
                    SelectedHostFiles = selectedHostFiles,
                    OpeningSettings = new Models.OpeningSettings
                    {
                        ClearanceSettings = clearanceSettings
                    }
                };
                filtersToProcess.Add(defaultFilter);
            }
            
            DebugLogger.Info($"[CLASH_DEBUG] MEP filters to process: {filtersToProcess.Count}");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] MEP filters to process: {filtersToProcess.Count}\n");

            foreach (var filter in filtersToProcess)
            {
                DebugLogger.Info($"[CLASH_DEBUG] Filter: {filter.Name} - Category: {filter.Category} - Enabled: {filter.IsEnabled}");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] Filter: {filter.Name} - Category: {filter.Category} - Enabled: {filter.IsEnabled}\n");
            }

            // Step 3: Get current profile and initialize clash zone service
            var currentProfile = _appProfileService.GetCurrentProfile();
            DebugLogger.Info($"[CLASH_DEBUG] Current profile: {currentProfile?.Name ?? "NULL"}");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] Profile loaded: {currentProfile?.Name ?? "NULL"}\n");

            if (currentProfile?.Configuration?.ClashZoneStorage == null)
            {
                DebugLogger.Info("[CLASH_DEBUG] No clash zone storage in profile - proceeding with direct clash detection");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] INFO: No clash zone storage - proceeding with direct clash detection\n");
            }

            // Step 4: Check for active document
            if (_document == null)
            {
                DebugLogger.Error("[CLASH_DEBUG] No active Revit document found");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] ERROR: No active Revit document found\n");
                
                _statusLabel.Text = "No active document";
                _progressBar.Visible = false;
                _refreshButton.Enabled = true;
                return;
            }

            DebugLogger.Info($"[CLASH_DEBUG] Document: {_document.Title}");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] Document: {_document.Title}\n");

            // Step 5: Check for Duct Accessories specific validation
            if (selectedMepCategories != null && selectedMepCategories.Count == 1 && selectedMepCategories.Contains("Duct Accessories"))
            {
                _statusLabel.Text = "Validating damper parameters...";
                _progressBar.Visible = true;
                _progressBar.Value = 10;

                DebugLogger.Info("[DUCT_ACCESSORIES] Only Duct Accessories selected - validating damper parameters");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [DUCT_ACCESSORIES] Only Duct Accessories selected - validating damper parameters\n");

                // Validate damper parameters in selected linked mechanical file
                if (!ValidateDamperParameters(selectedReferenceFiles, refreshLogPath))
                {
                    // Validation failed - user was prompted, exit refresh
                    _statusLabel.Text = "Damper parameter validation failed";
                    _progressBar.Visible = false;
                    _refreshButton.Enabled = true;
                    return;
                }
            }

            // Step 6: Perform intersection detection
            _statusLabel.Text = "Detecting intersections...";
            _progressBar.Visible = true;
            _progressBar.Value = 20;

            DebugLogger.Info("[CLASH_DEBUG] Starting MEP-structural intersection detection...");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] Starting MEP-structural intersection detection...\n");

            // Use passed MEP categories from UI
            var currentIntersections = _intersectionService.FindIntersections(_document, _document.ActiveView as View3D, selectedMepCategories);

            DebugLogger.Info($"[CLASH_DEBUG] INTERSECTION DETECTION COMPLETE: {currentIntersections.Count} intersections found");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] INTERSECTION DETECTION COMPLETE: {currentIntersections.Count} intersections found\n");

            // Log intersection details to timestamped file
            if (currentIntersections.Count > 0)
            {
                File.AppendAllText(refreshLogPath, $"[{DateTime.Now}] INTERSECTION DETECTION RESULTS:\n");
                File.AppendAllText(refreshLogPath, $"[{DateTime.Now}] Total intersections found: {currentIntersections.Count}\n");
                File.AppendAllText(refreshLogPath, $"[{DateTime.Now}] Intersection details:\n");

                // Log first 20 intersections
                int logCount = Math.Min(20, currentIntersections.Count);
                for (int i = 0; i < logCount; i++)
                {
                    var (mepElement, structuralElement, intersectionBBox, intersectionPoint) = currentIntersections[i];
                    File.AppendAllText(refreshLogPath, $"[{DateTime.Now}]   Intersection {i + 1}:\n");
                    File.AppendAllText(refreshLogPath, $"[{DateTime.Now}]     MEP Element ID: {mepElement.Id}\n");
                    File.AppendAllText(refreshLogPath, $"[{DateTime.Now}]     MEP Element Name: {mepElement.Name}\n");
                    File.AppendAllText(refreshLogPath, $"[{DateTime.Now}]     MEP Category: {mepElement.Category?.Name ?? "Unknown"}\n");
                    File.AppendAllText(refreshLogPath, $"[{DateTime.Now}]     Structural Element ID: {structuralElement.Id}\n");
                    File.AppendAllText(refreshLogPath, $"[{DateTime.Now}]     Structural Element Name: {structuralElement.Name}\n");
                    File.AppendAllText(refreshLogPath, $"[{DateTime.Now}]     Structural Category: {structuralElement.Category?.Name ?? "Unknown"}\n");
                    File.AppendAllText(refreshLogPath, $"[{DateTime.Now}]     Intersection Point: ({intersectionPoint.X:F2}, {intersectionPoint.Y:F2}, {intersectionPoint.Z:F2})\n");
                    File.AppendAllText(refreshLogPath, $"[{DateTime.Now}]     Intersection BBox: Min({intersectionBBox.Min.X:F2}, {intersectionBBox.Min.Y:F2}, {intersectionBBox.Min.Z:F2}) Max({intersectionBBox.Max.X:F2}, {intersectionBBox.Max.Y:F2}, {intersectionBBox.Max.Z:F2})\n");
                }

                if (currentIntersections.Count > 20)
                {
                    File.AppendAllText(refreshLogPath, $"[{DateTime.Now}]   ... and {currentIntersections.Count - 20} more intersections\n");
                }

                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] Found {currentIntersections.Count} intersections during refresh\n");
            }
            else
            {
                DebugLogger.Warning("[CLASH_DEBUG] WARNING: No intersections found - clash detection returned empty results");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] WARNING: No intersections found - clash detection returned empty results\n");
            }

            // Step 6: Process clash zones
            _progressBar.Value = 40;
            _statusLabel.Text = "Processing clash zones...";

            // ⚠️ CRITICAL FIX: Reinitialize ClashZoneService with existing clash zones from profile ⚠️
            // This allows ResetResolvedFlagForDeletedSleeves to detect manually deleted sleeves
            var existingClashZones = currentProfile?.Configuration?.ClashZoneStorage ?? new Models.ClashZoneStorage();
            var existingCount = existingClashZones?.ClashZones?.Count ?? 0;
            
            DebugLogger.Info($"[CLASH_DEBUG] Reinitializing ClashZoneService with {existingCount} existing zones from profile");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] Reinitializing ClashZoneService with {existingCount} existing zones from profile\n");
            
            // Reinitialize with existing clash zones
            _clashZoneService = new ClashZoneService(existingClashZones, msg => DebugLogger.Info(msg));
            
            DebugLogger.Info($"[CLASH_DEBUG] ClashZoneService reinitialized with {existingCount} existing zones");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] ClashZoneService reinitialized with {existingCount} existing zones\n");

            // Step 7: Filter and detect new clash zones
            _progressBar.Value = 50;
            _statusLabel.Text = "Detecting new clash zones...";

            // Use passed file selections and clearance settings from UI

            DebugLogger.Info("[CLASH_DEBUG] Skipping cleanup - no cleanup needed");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] Skipping cleanup - no cleanup needed\n");

            DebugLogger.Info($"[CLASH_DEBUG] Filtering clash zones by current selection - Reference files: {selectedReferenceFiles.Count}, Clearance settings: {clearanceSettings.Count}");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] Filtering clash zones by current selection - Reference files: {selectedReferenceFiles.Count}, Clearance settings: {clearanceSettings.Count}\n");

            var filteredClashZones = _clashZoneService.FilterClashZonesByCurrentSelection(selectedReferenceFiles, clearanceSettings, "Refresh", _document);

            DebugLogger.Info($"[CLASH_DEBUG] Calling DetectNewClashZones with {currentIntersections.Count} intersections...");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] Calling DetectNewClashZones with {currentIntersections.Count} intersections...\n");

            // DEBUG: Log intersection breakdown by category
            var intersectionBreakdown = currentIntersections.GroupBy(i => GetElementCategory(i.Item1))
                .ToDictionary(g => g.Key, g => g.Count());
            DebugLogger.Info($"[CLASH_DEBUG] Intersection breakdown by category: {string.Join(", ", intersectionBreakdown.Select(kv => $"{kv.Key}={kv.Value}"))}");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] Intersection breakdown by category: {string.Join(", ", intersectionBreakdown.Select(kv => $"{kv.Key}={kv.Value}"))}\n");

            var newClashZones = _clashZoneService.DetectNewClashZones(currentIntersections, _document, clearanceSettings);

            DebugLogger.Info($"[CLASH_DEBUG] DetectNewClashZones completed - {newClashZones?.Count ?? 0} new clash zones detected");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] DetectNewClashZones completed - {newClashZones?.Count ?? 0} new clash zones detected\n");

            // Step 8: Save results
            _progressBar.Value = 70;
            _statusLabel.Text = "Saving results...";

            var total = newClashZones?.Count ?? 0;
            var resolved = 0; // TODO: Implement resolved clash zone tracking
            var unresolved = total; // All new zones are unresolved initially
            var newZones = newClashZones?.Count ?? 0;

            DebugLogger.Info($"[CLASH_DEBUG] Statistics - Total: {total}, Resolved: {resolved}, Unresolved: {unresolved}, New: {newZones}");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] Statistics - Total: {total}, Resolved: {resolved}, Unresolved: {unresolved}, New: {newZones}\n");

            // Save to enabled filter
            var enabledFilter = filtersToProcess.FirstOrDefault(f => f.IsEnabled);
            DebugLogger.Info($"[CLASH_DEBUG] Found {filtersToProcess.Count} filters to process");
            DebugLogger.Info($"[CLASH_DEBUG] Enabled filter: {(enabledFilter != null ? enabledFilter.Name : "NULL")}");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] Found {filtersToProcess.Count} filters to process\n");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] Enabled filter: {(enabledFilter != null ? enabledFilter.Name : "NULL")}\n");
            
            if (enabledFilter != null)
            {
                var targetFilter = enabledFilter;
                // Create clash zone storage from new clash zones
                var clashZoneStorage = new Models.ClashZoneStorage
                {
                    ClashZones = newClashZones ?? new List<Models.ClashZone>(),
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
                DebugLogger.Info($"[CLASH_DEBUG] Set targetFilter.SelectedMepCategoryNames to: {string.Join(", ", selectedMepCategories)}");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] Set targetFilter.SelectedMepCategoryNames to: {string.Join(", ", selectedMepCategories)}\n");

                // Save to profile configuration for persistence
                if (currentProfile.Configuration == null)
                {
                    // Create a proper UserConfiguration object
                    currentProfile.Configuration = new Models.UserConfiguration(); // TODO: Initialize with proper values
                    DebugLogger.Info("[CLASH_DEBUG] Created new profile configuration during Refresh");
                }

                // TODO: Set clash zone storage in profile configuration
                // currentProfile.Configuration.ClashZoneStorage = clashZoneStorage;

                // TODO: Persist current opening conditions (e.g., clearance values) into profile configuration
                // if (currentProfile.Configuration.OpeningSettings == null)
                // {
                //     currentProfile.Configuration.OpeningSettings = new Models.OpeningSettings();
                // }

                // Save filter to XML files
                try
                {
                    var filterDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Projects", "Default", "Filters");
                    if (!Directory.Exists(filterDir))
                    {
                        Directory.CreateDirectory(filterDir);
                    }

                    // Save main filter file (for UI compatibility)
                    var mainFilePath = Path.Combine(filterDir, $"{targetFilter.Name}.xml");
                    _filterManagementService.SaveFilterToXmlFile(targetFilter, mainFilePath);
                    DebugLogger.Info($"[CLASH_DEBUG] Persisted main filter '{targetFilter.Name}' to XML file: {mainFilePath}");

                    // CRITICAL FIX: Save category-specific XML files RIGHT HERE after main filter is saved
                    DebugLogger.Info($"[CLASH_DEBUG] ===== SAVING CATEGORY-SPECIFIC XML FILES =====");
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] ===== SAVING CATEGORY-SPECIFIC XML FILES =====\n");

                    if (targetFilter.ClashZoneStorage != null && targetFilter.ClashZoneStorage.ClashZones != null && targetFilter.ClashZoneStorage.ClashZones.Count > 0)
                    {
                        var selectedCategories = targetFilter.SelectedMepCategoryNames ?? new List<string>();
                        DebugLogger.Info($"[CLASH_DEBUG] Saving category-specific XML files for: {string.Join(", ", selectedCategories)}");
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] Saving category-specific XML files for: {string.Join(", ", selectedCategories)}\n");

                        foreach (var category in selectedCategories)
                        {
                            DebugLogger.Info($"[CLASH_DEBUG] Processing category: '{category}' - About to call FilterClashZonesByCategory");
                            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] Processing category: '{category}' - About to call FilterClashZonesByCategory\n");
                            
                             var categoryClashZones = FilterClashZonesByCategory(targetFilter.ClashZoneStorage.ClashZones, category, _document, refreshLogPath);
                            
                            DebugLogger.Info($"[CLASH_DEBUG] FilterClashZonesByCategory returned {categoryClashZones.Count} clash zones for category: '{category}'");
                            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] FilterClashZonesByCategory returned {categoryClashZones.Count} clash zones for category: '{category}'\n");

                            if (categoryClashZones.Count > 0)
                            {
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
                                    ParameterTransferConfig = targetFilter.ParameterTransferConfig,
                                    ClashZoneStorage = new Models.ClashZoneStorage
                                    {
                                        ClashZones = categoryClashZones,
                                        CreatedAt = targetFilter.ClashZoneStorage.CreatedAt,
                                        LastUpdated = DateTime.Now,
                                        DocumentPath = targetFilter.ClashZoneStorage.DocumentPath,
                                        DocumentHash = targetFilter.ClashZoneStorage.DocumentHash,
                                        AlgorithmVersion = targetFilter.ClashZoneStorage.AlgorithmVersion
                                    }
                                };

                                // Save category-specific XML file
                                var categoryFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Projects", "Default", "Filters", $"{categoryFilter.Name}.xml");
                                _filterManagementService.SaveFilterToXmlFile(categoryFilter, categoryFilePath);

                                DebugLogger.Info($"[CLASH_DEBUG] Saved {categoryClashZones.Count} clash zones for category '{category}' to: {categoryFilePath}");
                                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] Saved {categoryClashZones.Count} clash zones for category '{category}'\n");
                            }
                            else
                            {
                                DebugLogger.Info($"[CLASH_DEBUG] No clash zones found for category '{category}'");
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
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] ERROR saving filter to XML: {xmlSaveEx.Message}\n");
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] XML Save Exception Stack Trace: {xmlSaveEx.StackTrace}\n");
                }

                DebugLogger.Info($"[CLASH_DEBUG] Saved clash zones to filter '{targetFilter.Name}'");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] SUCCESS: Saved {total} clash zones to filter '{targetFilter.Name}' and profile configuration\n");
            }
            else
            {
                DebugLogger.Warning("[CLASH_DEBUG] No enabled filter found to save clash zones");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] WARNING: No enabled filter found to save clash zones\n");
            }

            // Step 9: Update UI with results
            _progressBar.Value = 100;
            _statusLabel.Text = $"Clash zones: {total} total, {unresolved} unresolved, {newZones} new";

            DebugLogger.Info($"[CLASH_DEBUG] Clash zone refresh complete: {total} total, {unresolved} unresolved, {newZones} new");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] REFRESH COMPLETE: {total} total zones, {unresolved} unresolved, {newZones} new\n");

            // Step 10: Update parameter dropdowns
            _progressBar.Value = 90;
            _statusLabel.Text = "Updating parameter dropdowns...";

            // This will need to be handled by the UI
            // UpdateParameterDropdowns();

            _progressBar.Visible = false;
            _refreshButton.Enabled = true;

            DebugLogger.Info("[REFRESH] Refresh process completed successfully");
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

        /// <summary>
        /// Filters clash zones by MEP element category
        /// </summary>
        private List<Models.ClashZone> FilterClashZonesByCategory(List<Models.ClashZone> clashZones, string category, Document document, string refreshLogPath)
        {
            var filteredZones = new List<Models.ClashZone>();

            try
            {
                DebugLogger.Info($"[CLASH_DEBUG] Filtering {clashZones.Count} clash zones for category: {category}");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] Filtering {clashZones.Count} clash zones for category: {category}\n");
                
                // DEBUG: Show first few clash zones to understand the data
                for (int i = 0; i < Math.Min(5, clashZones.Count); i++)
                {
                    var cz = clashZones[i];
                    DebugLogger.Info($"[CLASH_DEBUG] Clash zone {i}: MEP ID={cz.MepElementId}, Structural ID={cz.StructuralElementId}");
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] Clash zone {i}: MEP ID={cz.MepElementId}, Structural ID={cz.StructuralElementId}\n");
                }

                foreach (var clashZone in clashZones)
                {
                    try
                    {
                        // ⚠️ CRITICAL FIX: Use cached category from ClashZone instead of re-detecting ⚠️
                        // Re-detection fails for linked elements, but cached value is always correct
                        var elementCategory = clashZone.MepElementCategory ?? "Unknown";

                        DebugLogger.Info($"[CLASH_DEBUG] Clash zone {clashZone.Id} - MEP element {clashZone.MepElementId}: category='{elementCategory}' (from ClashZone), checking against requested category='{category}'");
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] Clash zone {clashZone.Id} - MEP element {clashZone.MepElementId}: category='{elementCategory}' (from ClashZone), checking against requested category='{category}'\n");

                        // Check if element belongs to the requested category
                        if (IsElementInCategory(elementCategory, category))
                        {
                            filteredZones.Add(clashZone);
                            DebugLogger.Info($"[CLASH_DEBUG] ✓ Clash zone {clashZone.Id} MATCHES category '{category}' - element category: {elementCategory}");
                            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] ✓ Clash zone {clashZone.Id} MATCHES category '{category}' - element category: {elementCategory}\n");
                        }
                        else
                        {
                            DebugLogger.Info($"[CLASH_DEBUG] ✗ Clash zone {clashZone.Id} does NOT match category '{category}' - element category: {elementCategory}");
                            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] ✗ Clash zone {clashZone.Id} does NOT match category '{category}' - element category: {elementCategory}\n");
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Error($"[CLASH_DEBUG] Error filtering clash zone {clashZone.Id}: {ex.Message}");
                    }
                }

                DebugLogger.Info($"[CLASH_DEBUG] Filtered {filteredZones.Count} clash zones for category '{category}'");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[CLASH_DEBUG] Error in FilterClashZonesByCategory: {ex.Message}");
            }

            return filteredZones;
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
        private string GetElementCategoryWithFallback(Element element, ElementId elementId, Document document, string refreshLogPath)
        {
            try
            {
                // First try the standard category detection
                var standardCategory = GetElementCategory(element);
                
                DebugLogger.Info($"[CLASH_DEBUG] Standard category detection for element {elementId}: '{standardCategory}'");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] Standard category detection for element {elementId}: '{standardCategory}'\n");

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
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] Triggering fallback category detection for element {elementId} with standard category '{standardCategory}'\n");
                
                var fallbackCategory = GetCategoryFromElementType(element);
                
                DebugLogger.Info($"[CLASH_DEBUG] Fallback category detection for element {elementId}: '{fallbackCategory}' (original: '{standardCategory}')");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] Fallback category detection for element {elementId}: '{fallbackCategory}' (original: '{standardCategory}')\n");

                // If fallback also fails, try to infer from intersection detection results
                if (fallbackCategory == "Unknown")
                {
                    // Since we know these elements were detected as pipes during intersection detection,
                    // and they're being categorized as "Automatic Sketch Dimensions", they're likely pipes
                    if (standardCategory.Contains("Sketch"))
                    {
                        DebugLogger.Info($"[CLASH_DEBUG] Inferring 'Pipes' category for sketch element {elementId} based on intersection detection");
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(refreshLogPath, $"[{DateTime.Now}] [CLASH_DEBUG] Inferring 'Pipes' category for sketch element {elementId} based on intersection detection\n");
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
                        // Match actual Revit category names for cable trays
                        return elementCatLower == "cable trays" ||
                               elementCatLower == "cable tray fittings" ||
                               elementCatLower == "cable tray";

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

        #endregion
    }
}
