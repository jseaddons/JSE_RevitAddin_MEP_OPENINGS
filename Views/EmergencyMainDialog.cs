using System;
using System.Drawing;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using WinForms = System.Windows.Forms;

namespace JSE_RevitAddin_MEP_OPENINGS.Views
{
    /// <summary>
    /// Emergency WinForms MainDialog - NO WPF - CRASH SAFE for Revit
    /// 4-section layout with toolbar and status
    /// </summary>
    public partial class EmergencyMainDialog : WinForms.Form
    {
        private readonly ApplicationProfileService _appProfileService;
        private readonly FilterManagementService _filterManagementService;
        private string _lastLoadedFilterName = string.Empty;
        private readonly Document? _document;
        private readonly UIDocument? _uiDocument;
        
        // External Event for proper UI-to-Revit API communication
        private SleevePlacementExternalEvent _sleevePlacementHandler;
        private ExternalEvent _sleevePlacementEvent;
        
        // Store all collected parameters globally to persist across refreshes
        private Dictionary<string, List<Models.ParameterInfo>> _allCollectedParameters = new Dictionary<string, List<Models.ParameterInfo>>();
        
        // Flag to prevent infinite loops during ComboBox population
        private bool _isUpdatingComboBoxes = false;
        
        // Flag to track if user has made manual changes to UI selections
        private bool _userHasMadeManualChanges = false;
        
        // Store filters from last refresh to reuse during opening creation
        private bool _isInitializing = true; // Flag to prevent refresh during initialization
        
        
        // Main panels - 4-section layout
        private WinForms.Panel _leftPanel = null!;
        private WinForms.Panel _rightPanel = null!;
        private WinForms.Panel _mainSplitter = null!;  // Using Panel instead of Splitter to avoid docking requirement
        
        // NEW: Filters panel (left column)
        private WinForms.Panel _filtersPanel = null!;
        
        // Main content panel (center) - displays intersection results
        private WinForms.Panel _mainContentPanel = null!;
        private WinForms.DataGridView _intersectionDataGrid = null!;
        private WinForms.Splitter _filtersSplitter = null!;
        
        // 4 sections within left panel (2x2 grid)
        private WinForms.Panel _topLeftPanel = null!;      // Reference Elements (linked files)
        private WinForms.Panel _topRightPanel = null!;     // Reference Categories (MEP categories)
        private WinForms.Panel _bottomLeftPanel = null!;   // Host Elements (linked files)
        private WinForms.Panel _bottomRightPanel = null!;  // Host Categories (Revit categories)
        private WinForms.Splitter _horizontalSplitter = null!;
        private WinForms.Splitter _verticalSplitter = null!;
        
        // Header and toolbar
        private WinForms.Panel _headerPanel = null!;
        private WinForms.Label _titleLabel = null!;
        private WinForms.Label _profileLabel = null!;
        private WinForms.Panel _statusBannerPanel = null!;
        private WinForms.Label _statusBannerLabel = null!;
        
        // Bottom status bar
        private WinForms.Panel _statusPanel = null!;
        private WinForms.Label _statusLabel = null!;
        private WinForms.ProgressBar _progressBar = null!;
        
        // Toolbar buttons
        private WinForms.Button _okButton = null!;
        private WinForms.Button _cancelButton = null!;
        private WinForms.Button _saveButton = null!;
        private WinForms.Button _closeButton = null!;
        private WinForms.Button _parameterTransferButton = null!;
        
        // Bottom control bar buttons (scaffolding only)
        private WinForms.Button _refreshButton = null!;
        // ✅ REMOVED: Configure button - Settings integrated into right panel
        
        
        // Dynamic UI controls (only what we actually use)
        private WinForms.ComboBox _mepTypeCombo = null!;
        private WinForms.Panel _clearancePanel = null!;
        private WinForms.Panel _cableTrayPanel = null!;
        private WinForms.Panel _damperPanel = null!;
        private WinForms.Panel _pipePanel = null!;
        
        // ✅ INTEGRATED SETTINGS: Settings controls in right panel (below Clearance)
        private WinForms.CheckBox _enableThreePointValidationCheckBox = null!;
        private WinForms.CheckBox _forceDetectionModeCheckBox = null!;
        private WinForms.TextBox _ignoreOpeningsSmallerThanTextBox = null!;
        private WinForms.TextBox _roundOpeningsRectangularTextBox = null!;
        private WinForms.TextBox _joinOpeningsDistanceTextBox = null!;
        private WinForms.TextBox _roundingValueTextBox = null!;
        private WinForms.CheckBox _roundAlwaysUpCheckBox = null!;
        private WinForms.TextBox _minWallThicknessTextBox = null!;
        private WinForms.CheckBox _ignoreArchitecturalFloorsCheckBox = null!;
        private WinForms.Panel _settingsPanel = null!; // Scrollable panel for settings
        
        // ✅ UI STATE PERSISTENCE: Store references to host category listboxes for proper save/load
        private WinForms.CheckedListBox _horizontalCategoriesListBox = null!;
        private WinForms.CheckedListBox _verticalCategoriesListBox = null!;
        private bool _hostCategoriesInitialized = false;
        
        // ═══════════════════════════════════════════════════════════════
        // ✨ NEW: Mark Prefix Panel Controls (MEPMARK Implementation)
        // Added: 2025-10-08 for custom discipline-specific mark generation
        // ═══════════════════════════════════════════════════════════════
        private WinForms.Panel _markPrefixPanel = null!;
        private WinForms.TextBox _projectPrefixTextBox = null!;
        private WinForms.TextBox _disciplinePrefixTextBox = null!;
        private WinForms.CheckBox _remarkAllCheckBox = null!;
        
        // In-memory storage for category-specific discipline prefixes
        // Synced when user switches MEP Type dropdown
        private Dictionary<string, string> _categoryPrefixes = new Dictionary<string, string>
        {
            { MepCategoryConstants.DUCTS, "DCT" },
            { MepCategoryConstants.PIPES, "PLU" },
            { MepCategoryConstants.CABLE_TRAYS, "ELE" },
            { MepCategoryConstants.DUCT_ACCESSORIES, "DMP" }
        };
        // ═══════════════════════════════════════════════════════════════
        
        // Store clearance values per category to persist user changes when switching categories
        // Format: Dictionary<category, Dictionary<textboxTag, value>>
        private Dictionary<string, Dictionary<string, string>> _categoryClearanceValues = new Dictionary<string, Dictionary<string, string>>();
        private string _currentCategory = string.Empty; // Track current category for persistence
        
        private LinkedFileService? _linkedFileService;
        private List<LinkedFileInfo> _linkedFiles = new List<LinkedFileInfo>();
        private Document? _activeDocument;
        
        // Parameter service moved to separate dialog
        private const int InnerRightWidth = 320;  // choose 300–360

        public EmergencyMainDialog(ApplicationProfileService appProfileService, Document? document = null, UIDocument? uiDocument = null)
        {
            try
            {
                // 🔍 DIAGNOSTIC: Log constructor start IMMEDIATELY to file (bypass DebugLogger)
                DebugLogger.Info($"[{DateTime.Now}] 🔍 Constructor: EmergencyMainDialog constructor STARTED\n");
                
                // 🔍 DIAGNOSTIC: Log parameters
                // ✅ DEPLOYMENT MODE: Skip hardcoded log writes if deployment mode is enabled
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string constructorDebugLogPath = SafeFileLogger.GetLogFilePath("constructor_debug.log");
                    System.IO.File.AppendAllText(constructorDebugLogPath, $"[{DateTime.Now}] 🔍 Constructor: appProfileService = {(appProfileService != null ? "NOT NULL" : "NULL")}\n");
                    System.IO.File.AppendAllText(constructorDebugLogPath, $"[{DateTime.Now}] 🔍 Constructor: document = {(document != null ? "NOT NULL" : "NULL")}\n");
                    System.IO.File.AppendAllText(constructorDebugLogPath, $"[{DateTime.Now}] 🔍 Constructor: uiDocument = {(uiDocument != null ? "NOT NULL" : "NULL")}\n");
                }
            }
            catch (Exception ex)
            {
                try
                {
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 Constructor: ERROR in initial logging: {ex.Message}\n");
                }
                catch { }
            }
            
            _appProfileService = appProfileService;
            _document = document;
            _uiDocument = uiDocument;
            
            // ✅ CRITICAL: Check section box BEFORE initializing UI - NO FALLBACK
            var docToCheck = document ?? uiDocument?.Document;
            if (docToCheck != null)
            {
                var activeView = uiDocument?.ActiveView ?? docToCheck.ActiveView;
                if (!(activeView is View3D view3D) || !view3D.IsSectionBoxActive)
                {
                    // ✅ NO FALLBACK: Section box is REQUIRED - don't initialize UI
                    throw new InvalidOperationException(
                        "A 3D view with an active section box is REQUIRED.\n\n" +
                        "Please:\n" +
                        "1. Activate a 3D view\n" +
                        "2. Enable section box in the view properties\n" +
                        "3. Adjust section box to your desired zone\n" +
                        "4. Try again\n\n" +
                        "Section box is required to limit clash detection to specific zones.");
                }
                
                // ✅ DOUBLE CHECK: Verify section box is not null
                var sectionBox = view3D.GetSectionBox();
                if (sectionBox == null || sectionBox.Min == null || sectionBox.Max == null)
                {
                    throw new InvalidOperationException(
                        "Section box is active but invalid.\n\n" +
                        "Please:\n" +
                        "1. Deactivate section box\n" +
                        "2. Reactivate section box\n" +
                        "3. Adjust section box bounds\n" +
                        "4. Try again");
                }
            }
            
            // Ensure per-project Filters directory exists as soon as the dialog opens
            try
            {
                if (_document != null)
                {
                    // Provide project title to services that rely on environment for default paths
                    try { Environment.SetEnvironmentVariable("JSE_ACTIVE_DOC_TITLE", _document.Title ?? string.Empty); } catch { }

                    ProjectPathService.EnsureFiltersDirectory(_document);
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Info($"EmergencyMainDialog: EnsureFiltersDirectory failed: {ex.Message}");
            }
            
            // Close all log files to free file handles
            try
            {
                DebugLogger.Info($"[{DateTime.Now}] 🔍 Constructor: About to close log files\n");
            }
            catch { }
            
            DebugLogger.CloseAllLogFiles();
            DebugLogger.Info("🔍 Constructor: Closed all log files");
            
            try
            {
                DebugLogger.Info($"[{DateTime.Now}] 🔍 Constructor: Closed all log files\n");
            }
            catch { }
            
            // Set logging context for OK button debugging
            DebugLogger.SetServiceContext("OKButton");
            DebugLogger.Info("🔍 Constructor: Set service context to OKButton");
            
            // STEP 1: IMMEDIATE LOG - Use rolling log system for main UI
            // Clean up old logs first to prevent accumulation
            DebugLogger.CleanupOldLogs();
            DebugLogger.Info("🔍 Constructor: Cleaned up old logs");

            // Set logging context for main UI
            DebugLogger.SetServiceContext("MainUI");
            DebugLogger.Info("🔍 Constructor: Set service context to MainUI");
            
            // ✅ REMOVED: Main UI log initialization per user request
            // DebugLogger.InitCustomLogFileOverwrite("MainUi");
            DebugLogger.Info($"Emergency Main Dialog Constructor Started");
            DebugLogger.Info($"ApplicationProfileService: {appProfileService != null}");
            DebugLogger.Info($"Document: {document?.Title ?? "null"}");
            DebugLogger.Info("🔍 Constructor: DebugLogger initialized successfully");
            // DebugLogger handles its own error handling - no need for catch block

            // STEP 2: Continue with normal initialization
            DebugLogger.Info("🔍 Constructor: About to initialize services");
            _appProfileService = appProfileService ?? throw new ArgumentNullException(nameof(appProfileService));
            DebugLogger.Info("🔍 Constructor: ApplicationProfileService initialized");
            
            _filterManagementService = new FilterManagementService(
                document,
                msg => DebugLogger.Info(msg),
                msg => _statusLabel.Text = msg
            );
            DebugLogger.Info("🔍 Constructor: FilterManagementService initialized");
            
            _linkedFileService = new LinkedFileService();
            DebugLogger.Info("🔍 Constructor: LinkedFileService initialized");
            
            _activeDocument = document;
            DebugLogger.Info("🔍 Constructor: ActiveDocument set");

            // Log successful initialization
            DebugLogger.Info("Services initialized successfully");
            DebugLogger.Info("🔍 Constructor: All services initialized successfully");

            // DebugLogger already initialized above - no need to reinitialize

            try
            {
                // 🔍 DIAGNOSTIC: Count sleeves BEFORE InitializeComponent
                if (_document != null)
                {
                    var sleevesBeforeInit = new FilteredElementCollector(_document)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                        .Count();
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 Constructor: Sleeves BEFORE InitializeComponent: {sleevesBeforeInit}\n");
                }
                
                DebugLogger.Info("About to call InitializeComponent()");
                InitializeComponent();
                DebugLogger.Info("InitializeComponent() completed successfully");
                
                // ✅ AUTO-SAVE: Add FormClosing event handler to save UI state when form closes
                this.FormClosing += (sender, e) => {
                    try
                    {
                        DebugLogger.Info("[UI_STATE] FormClosing event triggered - saving UI state");
                        SaveUIStateDirectly();
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Error($"[UI_STATE] Failed to save UI state on form close: {ex.Message}");
                    }
                };
                
                // 🔍 DIAGNOSTIC: Count sleeves AFTER InitializeComponent
                if (_document != null)
                {
                    var sleevesAfterInit = new FilteredElementCollector(_document)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                        .Count();
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 Constructor: Sleeves AFTER InitializeComponent: {sleevesAfterInit}\n");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"InitializeComponent() failed: {ex.Message}");
                DebugLogger.Error($"Stack trace: {ex.StackTrace}");
                throw; // Re-throw to see the error
            }

            try
            {
                // 🔍 DIAGNOSTIC: Count sleeves BEFORE LoadProfileInfo
                if (_document != null)
                {
                    var sleevesBeforeProfile = new FilteredElementCollector(_document)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                        .Count();
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 Constructor: Sleeves BEFORE LoadProfileInfo: {sleevesBeforeProfile}\n");
                }
                
                DebugLogger.Info("About to call LoadProfileInfo()");
                LoadProfileInfo();
                DebugLogger.Info("LoadProfileInfo() completed successfully");
                
                // 🔍 DIAGNOSTIC: Count sleeves AFTER LoadProfileInfo
                if (_document != null)
                {
                    var sleevesAfterProfile = new FilteredElementCollector(_document)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                        .Count();
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 Constructor: Sleeves AFTER LoadProfileInfo: {sleevesAfterProfile}\n");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"LoadProfileInfo() failed: {ex.Message}");
                DebugLogger.Error($"Stack trace: {ex.StackTrace}");
                throw; // Re-throw to see the error
            }

            // STEP 3: Validate required families are loaded - DISABLED
            // Family validation is now handled only for Duct Accessories in linked mechanical files
            // during refresh when Duct Accessories category is selected
            try
            {
                DebugLogger.Info("Family validation disabled - only checking damper parameters in linked mechanical files for Duct Accessories");
                // Family validation removed - using universal opening families only
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Family validation failed: {ex.Message}");
                DebugLogger.Error($"Stack trace: {ex.StackTrace}");
                // Don't throw - just log the error and continue
            }

            // STEP 4: Initialize External Event for proper sleeve placement
            try
            {
                // 🔍 DIAGNOSTIC: Count sleeves BEFORE External Event initialization
                if (_document != null)
                {
                    var sleevesBeforeExternal = new FilteredElementCollector(_document)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                        .Count();
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 Constructor: Sleeves BEFORE External Event init: {sleevesBeforeExternal}\n");
                }
                
                _sleevePlacementHandler = new SleevePlacementExternalEvent();
                
                // 🔥 CRITICAL DEBUG: Verify DLL is being loaded with current build
                DebugLogger.Info($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 🔥 DLL LOADED - SleevePlacementExternalEvent constructed\n");
                
                _sleevePlacementHandler.PlacementCompleted += () =>
                {
                    try
                    {
                        if (this.IsHandleCreated)
                        {
                            this.BeginInvoke(new Action(() =>
                            {
                                CompletePlacementProgress("Sleeve placement complete");
                                _okButton.Enabled = true;
                                _refreshButton.Enabled = true;
                                this.Show();
                                this.Activate();
                            }));
                        }
                    }
                    catch { }
                };
                _sleevePlacementEvent = ExternalEvent.Create(_sleevePlacementHandler);
                DebugLogger.Info("External Event initialized successfully for sleeve placement");
                
                // 🔍 DIAGNOSTIC: Count sleeves AFTER External Event initialization
                if (_document != null)
                {
                    var sleevesAfterExternal = new FilteredElementCollector(_document)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                        .Count();
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 Constructor: Sleeves AFTER External Event init: {sleevesAfterExternal}\n");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to initialize External Event: {ex.Message}");
            }

            // STEP 4: Log completion
            DebugLogger.Info("EmergencyMainDialog initialization COMPLETED");
            
            // 🔍 DIAGNOSTIC: Add logging to track constructor progress
            DebugLogger.Info("🔍 Constructor: About to set up Shown event handler");

            // Load real linked files if document is provided (after UI is initialized)
            if (document != null)
            {
                this.Shown += (s, e) => {
                    try
                    {
                        DebugLogger.Info("=== DIALOG SHOWN EVENT TRIGGERED ===");
                        
                        // ✅ FAMILY VALIDATION: Check for required opening families
                        var familyService = new FamilyLoadingService(document);
                        var missingFamilies = familyService.GetMissingFamilies();
                        
                        if (missingFamilies.Count > 0)
                        {
                            var missingList = string.Join("\n  • ", missingFamilies);
                            var resourcesPath = familyService.ResourcesPath;
                            
                            var dialog = new TaskDialog("Missing Required Families")
                            {
                                MainIcon = TaskDialogIcon.TaskDialogIconWarning,
                                MainInstruction = "Required Opening Families Not Loaded",
                                MainContent = $"The following opening families are required but not loaded in this project:\n\n  • {missingList}\n\n" +
                                             $"Please load these families using:\n" +
                                             $"  Insert → Load Family",
                                CommonButtons = TaskDialogCommonButtons.Ok,
                                Title = "Family Loading Required"
                            };
                            
                            DebugLogger.Warning($"[EmergencyMainDialog] Missing families detected: {missingList}");
                            DebugLogger.Info($"[EmergencyMainDialog] Resources path: {resourcesPath}");
                            
                            // ✅ CRITICAL FIX: Hide main dialog, show TaskDialog, then close UI so user can load families
                            this.Hide();
                            try
                            {
                                // Show TaskDialog (will appear on top since main dialog is hidden)
                                dialog.Show();
                                
                                // ✅ AUTO-LOAD: Attempt to automatically load families if Resources folder is deployed
                                // Resources folder is copied by installer to: [Addins Folder]\Resources\
                                if (familyService.IsResourcesFolderValid())
                                {
                                    DebugLogger.Info("[EmergencyMainDialog] Resources folder found - attempting to auto-load missing families...");
                                    bool loaded = familyService.LoadAllRequiredFamilies();
                                    
                                    if (loaded)
                                    {
                                        // Re-check after loading
                                        var stillMissing = familyService.GetMissingFamilies();
                                        if (stillMissing.Count == 0)
                                        {
                                            // All families loaded successfully - show success and restore UI
                                            var successDialog = new TaskDialog("Families Loaded")
                                            {
                                                MainIcon = TaskDialogIcon.TaskDialogIconInformation,
                                                MainInstruction = "Opening Families Loaded Successfully",
                                                MainContent = "All required opening families have been automatically loaded into your project.",
                                                CommonButtons = TaskDialogCommonButtons.Ok,
                                                Title = "Success"
                                            };
                                            successDialog.Show();
                                            this.Show();
                                            this.BringToFront();
                                            this.Activate();
                                            return; // Exit early - UI restored, no need to close
                                        }
                                        else
                                        {
                                            // Auto-load partially succeeded but some families still missing
                                            DebugLogger.Warning($"[EmergencyMainDialog] Auto-load incomplete: {stillMissing.Count} families still missing - closing UI for manual loading");
                                        }
                                    }
                                    else
                                    {
                                        DebugLogger.Warning("[EmergencyMainDialog] Auto-load failed - closing UI for manual loading");
                                    }
                                }
                                else
                                {
                                    // Resources folder not found - user must load families manually
                                    DebugLogger.Info($"[EmergencyMainDialog] Resources folder not found at: {familyService.ResourcesPath} - user must load families manually");
                                }
                            }
                            catch (Exception ex)
                            {
                                DebugLogger.Error($"[EmergencyMainDialog] Error during family loading: {ex.Message}");
                            }
                            
                            // Close UI so user can load families manually (either auto-load didn't work or user needs to load manually)
                            this.DialogResult = WinForms.DialogResult.Cancel;
                            this.Close();
                        }
                        
                        // 🔍 DIAGNOSTIC: Count sleeves BEFORE LoadRealLinkedFiles
                        var sleevesBeforeCount = new FilteredElementCollector(document)
                            .OfClass(typeof(FamilyInstance))
                            .Cast<FamilyInstance>()
                            .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                            .Count();
                        DebugLogger.Info($"🔍 Sleeves BEFORE LoadRealLinkedFiles: {sleevesBeforeCount}");
                        
                        LoadRealLinkedFiles(document);
                        
                        // 🔍 DIAGNOSTIC: Count sleeves AFTER LoadRealLinkedFiles
                        var sleevesAfterCount = new FilteredElementCollector(document)
                            .OfClass(typeof(FamilyInstance))
                            .Cast<FamilyInstance>()
                            .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                            .Count();
                        DebugLogger.Info($"🔍 Sleeves AFTER LoadRealLinkedFiles: {sleevesAfterCount}");
                        DebugLogger.Info($"🔍 Sleeves DELETED: {sleevesBeforeCount - sleevesAfterCount}");
                        
                        if (sleevesBeforeCount != sleevesAfterCount)
                        {
                            DebugLogger.Error($"⚠️ CRITICAL BUG DETECTED: {sleevesBeforeCount - sleevesAfterCount} sleeves were deleted when dialog opened!");
                        }
                        
                        // ⚠️ CRITICAL FIX: Mark initialization as complete to allow live parameter extraction
                        _isInitializing = false;
                        DebugLogger.Info("🔍 Initialization complete - _isInitializing set to false");
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Error($"Error in Shown event diagnostic logging: {ex.Message}");
                        LoadRealLinkedFiles(document);
                        
                        // ⚠️ CRITICAL FIX: Mark initialization as complete even if error occurs
                        _isInitializing = false;
                        DebugLogger.Info("🔍 Initialization complete (after error) - _isInitializing set to false");
                    }
                };
            }

            // 🔍 DIAGNOSTIC: Log after Shown event setup
            DebugLogger.Info("🔍 Constructor: Shown event handler setup completed");
            
            // Register FilterUiStateProvider delegates for Refresh/Filter services
            try
            {
                Services.FilterUiStateProvider.GetSelectedFilterItems = () =>
                {
                    return GetSelectedFilterItems();
                };

                Services.FilterUiStateProvider.GetSelectedMepCategoryNames = () =>
                {
                    return GetSelectedMepCategories();
                };

                Services.FilterUiStateProvider.GetSelectedHostCategories = () =>
                {
                    return GetSelectedHostCategories();
                };

                Services.FilterUiStateProvider.GetSelectedReferenceFiles = () =>
                {
                    return GetSelectedReferenceFiles();
                };

                Services.FilterUiStateProvider.GetSelectedHostFiles = () =>
                {
                    return GetSelectedHostFiles();
                };

                Services.FilterUiStateProvider.GetClearanceSettings = (category) =>
                {
                    return GetClearanceSettings(category);
                };

                Services.FilterUiStateProvider.ApplyFilterToUi = (filter) =>
                {
                    try
                    {
                        // ✅ CRITICAL FIX: Apply BOTH MEP categories (top-right) AND host categories (bottom-right)
                        
                        // Apply MEP categories to top-right lists
                        var mepCategories = filter?.SelectedMepCategoryNames ?? new List<string>();
                        DebugLogger.Info($"[FILTER_UI] Applying {mepCategories.Count} MEP categories to UI: {string.Join(", ", mepCategories)}");
                        var mepCategoryLists = _topRightPanel?.Controls?.OfType<System.Windows.Forms.CheckedListBox>()?.ToList();
                        if (mepCategoryLists != null)
                        {
                            foreach (var lb in mepCategoryLists)
                            {
                                for (int i = 0; i < lb.Items.Count; i++)
                                {
                                    var name = lb.Items[i]?.ToString() ?? string.Empty;
                                    bool shouldCheck = mepCategories.Contains(name, StringComparer.OrdinalIgnoreCase);
                                    lb.SetItemChecked(i, shouldCheck);
                                    if (shouldCheck)
                                    {
                                        DebugLogger.Info($"[FILTER_UI] ✅ Checked MEP category: {name}");
                                    }
                                }
                            }
                        }
                        
                        // Apply host categories to bottom-right lists
                        var hostCategories = filter?.SelectedHostCategories ?? new List<string>();
                        DebugLogger.Info($"[FILTER_UI] Applying {hostCategories.Count} host categories to UI: {string.Join(", ", hostCategories)}");
                        var hostCategoryLists = _bottomRightPanel?.Controls?.OfType<System.Windows.Forms.CheckedListBox>()?.ToList();
                        if (hostCategoryLists != null)
                        {
                            foreach (var lb in hostCategoryLists)
                            {
                                for (int i = 0; i < lb.Items.Count; i++)
                                {
                                    var name = lb.Items[i]?.ToString() ?? string.Empty;
                                    bool shouldCheck = hostCategories.Contains(name, StringComparer.OrdinalIgnoreCase);
                                    lb.SetItemChecked(i, shouldCheck);
                                    if (shouldCheck)
                                    {
                                        DebugLogger.Info($"[FILTER_UI] ✅ Checked host category: {name}");
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Error($"[FILTER_UI] Error applying filter to UI: {ex.Message}");
                    }
                };
            }
            catch { }

        }
        
        private void StartPlacementProgress(string status)
        {
            try
            {
                if (_progressBar != null)
                {
                    _progressBar.Style = WinForms.ProgressBarStyle.Marquee;
                    _progressBar.MarqueeAnimationSpeed = 30;
                    _progressBar.Value = Math.Max(_progressBar.Minimum, Math.Min(_progressBar.Maximum, _progressBar.Value));
                    _progressBar.Visible = true;
                }
                
                if (_statusLabel != null)
                {
                    _statusLabel.Text = status;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[EmergencyMainDialog] Failed to start placement progress: {ex.Message}");
            }
        }
        
        private void CompletePlacementProgress(string status)
        {
            try
            {
                if (_progressBar != null)
                {
                    _progressBar.MarqueeAnimationSpeed = 0;
                    _progressBar.Style = WinForms.ProgressBarStyle.Continuous;
                    _progressBar.Value = _progressBar.Minimum;
                    _progressBar.Visible = false;
                }
                
                if (_statusLabel != null)
                {
                    _statusLabel.Text = status;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[EmergencyMainDialog] Failed to complete placement progress: {ex.Message}");
            }
        }
        private void InitializeComponent()
        {
            // 🔍 DIAGNOSTIC: Count sleeves at START of InitializeComponent
            try
            {
                if (_document != null)
                {
                    var sleevesAtStart = new FilteredElementCollector(_document)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                        .Count();
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 InitializeComponent START: {sleevesAtStart} sleeves\n");
                }
            }
            catch { }
            
            this.SuspendLayout();

            // Form properties
            this.Text = "JSE MEP Openings - Main Interface (Emergency WinForms Mode)";
            this.Size = new System.Drawing.Size(1100, 800);   // smaller default
            this.StartPosition = WinForms.FormStartPosition.CenterScreen;
            this.MinimumSize = new System.Drawing.Size(900, 600);  // was 1000,600
            this.TopMost = true;

            // Header Panel
            _headerPanel = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.Top,
                Height = 80,
                BackColor = System.Drawing.Color.FromArgb(240, 240, 240),
                BorderStyle = WinForms.BorderStyle.FixedSingle
            };
            this.Controls.Add(_headerPanel);

            // Title Label
            _titleLabel = new WinForms.Label
            {
                Text = "MEP Openings Management",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51),
                Location = new System.Drawing.Point(20, 15),
                Size = new System.Drawing.Size(400, 30),
                AutoSize = false
            };
            _headerPanel.Controls.Add(_titleLabel);

            // Profile Label
            _profileLabel = new WinForms.Label
            {
                Text = "Profile: Loading...",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular),
                ForeColor = System.Drawing.Color.FromArgb(102, 102, 102),
                Location = new System.Drawing.Point(20, 50),
                Size = new System.Drawing.Size(400, 20),
                AutoSize = false
            };
            _headerPanel.Controls.Add(_profileLabel);

            // Status Banner Panel (shown when no linked files)
            _statusBannerPanel = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.Top,
                Height = 40,
                BackColor = System.Drawing.Color.FromArgb(255, 235, 59), // Warning yellow
                BorderStyle = WinForms.BorderStyle.FixedSingle,
                Visible = false // Hidden by default
            };
            this.Controls.Add(_statusBannerPanel);

            _statusBannerLabel = new WinForms.Label
            {
                Text = "WARNING: No linked files found. Please link MEP, architectural, and structural files to use this feature.",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51),
                Location = new System.Drawing.Point(10, 8),
                Size = new System.Drawing.Size(_statusBannerPanel.Width - 20, 25),
                AutoSize = false,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };
            _statusBannerPanel.Controls.Add(_statusBannerLabel);

            // Toolbar buttons in header - positioned on the right side
            int buttonY = 15;
            int buttonHeight = 50; // Increased height for 2-line text
            int buttonWidth = 130; // Width for 2-line text buttons
            int buttonSpacing = 15; // Spacing between buttons
            int rightMargin = 15; // Margin from right edge

            // Button layout (right to left): [Close] [Place Sleeves] [Process Clash Zones]
            int closeButtonWidth = 70;
            
            // ✅ Process Clash Zones Button (2-line text, enlarged) - leftmost
            _refreshButton = new WinForms.Button
            {
                Text = "Process" + Environment.NewLine + "Clash Zones",
                Size = new System.Drawing.Size(buttonWidth, buttonHeight),
                BackColor = System.Drawing.Color.FromArgb(108, 117, 125),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = WinForms.FlatStyle.Flat,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Right
            };
            int refreshButtonX = _headerPanel.Width - rightMargin - closeButtonWidth - buttonSpacing - buttonWidth - buttonSpacing - buttonWidth;
            _refreshButton.Location = new System.Drawing.Point(refreshButtonX, buttonY);
            _refreshButton.Click += OnRefreshClick;
            _headerPanel.Controls.Add(_refreshButton);

            // ✅ Place Sleeves Button (2-line text, enlarged) - middle
            _okButton = new WinForms.Button
            {
                Text = "Place" + Environment.NewLine + "Sleeves",
                Size = new System.Drawing.Size(buttonWidth, buttonHeight),
                BackColor = System.Drawing.Color.FromArgb(0, 122, 204),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = WinForms.FlatStyle.Flat,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Enabled = false, // Enable only after Refresh completes
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Right
            };
            int okButtonX = _headerPanel.Width - rightMargin - closeButtonWidth - buttonSpacing - buttonWidth;
            _okButton.Location = new System.Drawing.Point(okButtonX, buttonY);
            _okButton.Click += OnOkClick;
            _headerPanel.Controls.Add(_okButton);
            
            // ✅ Save and Cancel buttons removed - not being used

            // ✅ Close Button - rightmost
            _closeButton = new WinForms.Button
            {
                Text = "Close",
                Size = new System.Drawing.Size(closeButtonWidth, buttonHeight),
                BackColor = System.Drawing.Color.FromArgb(220, 53, 69),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = WinForms.FlatStyle.Flat,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold),
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Right
            };
            _closeButton.Location = new System.Drawing.Point(_headerPanel.Width - rightMargin - closeButtonWidth, buttonY);
            _closeButton.Click += OnCloseClick;
            _headerPanel.Controls.Add(_closeButton);

            // Parameter Transfer Button removed - functionality moved to Transfer All button in parameter service

            // Status Panel (bottom)
            _statusPanel = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.Bottom,
                Height = 30,
                BackColor = System.Drawing.Color.FromArgb(248, 249, 250),
                BorderStyle = WinForms.BorderStyle.FixedSingle
            };
            this.Controls.Add(_statusPanel);

            // Calculate left section right edge position (left panel width)
            int leftSectionRightEdge = (this.Width - InnerRightWidth) - 20; // 20px margin from right edge of left section
            
            // Status Label (left side, wider to show full error messages)
            _statusLabel = new WinForms.Label
            {
                Text = "Ready",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular),
                ForeColor = System.Drawing.Color.FromArgb(102, 102, 102),
                Location = new System.Drawing.Point(10, 5),
                Size = new System.Drawing.Size(200, 20), // Increased width to show full error messages
                AutoSize = false
            };
            
            // Add tooltip for status label to show full error messages
            var statusTooltip = new WinForms.ToolTip();
            statusTooltip.SetToolTip(_statusLabel, "Status information - hover to see full message");
            _statusPanel.Controls.Add(_statusLabel);

            int statusButtonSpacing = 5; // Space between status bar buttons
            int buttonStartX = 220; // Start position for buttons (after status label) - moved further right
            
            // ✅ NOTE: Refresh button moved to header panel (left of OK button)
            // Config button remains in status panel
            
            // ✅ REMOVED: Configure button - Settings integrated into right panel below Clearance
            // ✅ REMOVED: Progress bar - Will use progress dialog during refresh instead
            // Note: Parameter Service is now available as a separate command
            // This keeps the main UI focused on clash detection and sleeve placement
            
            // ✅ CREATE: Progress bar displayed in status panel for refresh feedback
            _progressBar = new WinForms.ProgressBar
            {
                Location = new System.Drawing.Point(buttonStartX, 5),
                Size = new System.Drawing.Size(_statusPanel.Width - buttonStartX - 10, 20),
                Style = WinForms.ProgressBarStyle.Continuous,
                Minimum = 0,
                Maximum = 100,
                Value = 0,
                Visible = false,
                Anchor = WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right | WinForms.AnchorStyles.Top
            };
            _statusPanel.Controls.Add(_progressBar);

            // Left Panel (expanded to fill most space) - start below header
            _leftPanel = new WinForms.Panel
            {
                BackColor = System.Drawing.Color.White,
                BorderStyle = WinForms.BorderStyle.FixedSingle
            };
            this.Controls.Add(_leftPanel);

            // Main Splitter - FIXED: Use Panel instead of Splitter to avoid docking requirement
            try
            {
                _mainSplitter = new WinForms.Panel  // Use Panel instead of Splitter
                {
                    Width = 3,
                    BackColor = System.Drawing.Color.Gray,
                    Dock = WinForms.DockStyle.None  // Panels can have Dock=None
                };
                this.Controls.Add(_mainSplitter);
                // DebugLogger.Info($"_mainSplitter (Panel) created: Width={_mainSplitter.Width}, Dock={_mainSplitter.Dock}");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to create _mainSplitter: {ex.Message}");
                throw;
            }

            // Right Panel - keep original width, make scrollable for Settings section
            _rightPanel = new WinForms.Panel
            {
                BackColor = System.Drawing.Color.White,
                BorderStyle = WinForms.BorderStyle.FixedSingle,
                Width = 460,  // Restored to original width
                AutoScroll = true // ✅ Enable scrolling for Settings section below Clearance
            };
            this.Controls.Add(_rightPanel);
            // DebugLogger.Info($"_rightPanel created: Width={_rightPanel.Width}, Location={_rightPanel.Location}, Dock={_rightPanel.Dock}");

            // Add content to panels
            
            // 🔍 DIAGNOSTIC: Count sleeves BEFORE InitializePanelContent
            try
            {
                if (_document != null)
                {
                    var sleevesBeforePanelContent = new FilteredElementCollector(_document)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                        .Count();
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 BEFORE InitializePanelContent: {sleevesBeforePanelContent} sleeves\n");
                }
            }
            catch { }
            
            InitializePanelContent();
            
            // 🔍 DIAGNOSTIC: Count sleeves AFTER InitializePanelContent
            try
            {
                if (_document != null)
                {
                    var sleevesAfterPanelContent = new FilteredElementCollector(_document)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                        .Count();
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 AFTER InitializePanelContent: {sleevesAfterPanelContent} sleeves\n");
                }
            }
            catch { }
            
            PositionPanels();
            BalanceLeftLayout();
            this.Shown += (_, __) => _bottomLeftPanel.Height = (_leftPanel.Height - _horizontalSplitter.Height) / 2;
            this.Resize += (_, __) => BalanceLeftLayout();
            _leftPanel.Resize += (_, __) => BalanceLeftLayout();

            // 🔍 DIAGNOSTIC: Count sleeves BEFORE ResumeLayout (end of InitializeComponent)
            try
            {
                if (_document != null)
                {
                    var sleevesBeforeResume = new FilteredElementCollector(_document)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                        .Count();
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 InitializeComponent END (before ResumeLayout): {sleevesBeforeResume} sleeves\n");
                }
            }
            catch { }

            this.ResumeLayout(false);
            
            // 🔍 DIAGNOSTIC: Log constructor completion
            DebugLogger.Info("🔍 Constructor: EmergencyMainDialog constructor COMPLETED successfully");
            
            try
            {
                DebugLogger.Info($"[{DateTime.Now}] 🔍 Constructor: EmergencyMainDialog constructor COMPLETED successfully\n");
            }
            catch (Exception ex)
            {
                try
                {
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 Constructor: CRITICAL ERROR in constructor: {ex.Message}\n");
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 Constructor: Stack trace: {ex.StackTrace}\n");
                }
                catch { }
                throw; // Re-throw to see the error
            }
        }

        private void InitializePanelContent()
        {
            // 🔍 DIAGNOSTIC: Count sleeves at START of InitializePanelContent
            try
            {
                if (_document != null)
                {
                    var sleevesAtStart = new FilteredElementCollector(_document)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                        .Count();
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 InitializePanelContent START: {sleevesAtStart} sleeves\n");
                }
            }
            catch { }
            
            // NEW: Create filters panel (left column)
            CreateFiltersPanel();
            
            // 🔍 DIAGNOSTIC: Count sleeves AFTER CreateFiltersPanel
            try
            {
                if (_document != null)
                {
                    var sleevesAfterFilters = new FilteredElementCollector(_document)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                        .Count();
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 AFTER CreateFiltersPanel: {sleevesAfterFilters} sleeves\n");
                }
            }
            catch { }
            
            // Create main content panel (center) - displays intersection results
            CreateMainContentPanel();
            
            // 🔍 DIAGNOSTIC: Count sleeves AFTER CreateMainContentPanel
            try
            {
                if (_document != null)
                {
                    var sleevesAfterMainContent = new FilteredElementCollector(_document)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                        .Count();
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 AFTER CreateMainContentPanel: {sleevesAfterMainContent} sleeves\n");
                }
            }
            catch { }
            
            // Create 4-section layout within left panel
            CreateFourSectionLayout();
            
            // 🔍 DIAGNOSTIC: Count sleeves AFTER CreateFourSectionLayout
            try
            {
                if (_document != null)
                {
                    var sleevesAfterFourSection = new FilteredElementCollector(_document)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                        .Count();
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 AFTER CreateFourSectionLayout: {sleevesAfterFourSection} sleeves\n");
                }
            }
            catch { }
            
            // Right Panel: Opening Configuration
            InitializeRightPanel();
            
            // 🔍 DIAGNOSTIC: Count sleeves AFTER InitializeRightPanel
            try
            {
                if (_document != null)
                {
                    var sleevesAfterRightPanel = new FilteredElementCollector(_document)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                        .Count();
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 AFTER InitializeRightPanel: {sleevesAfterRightPanel} sleeves\n");
                }
            }
            catch { }
        }

        private void CreateFiltersPanel()
        {
            // DebugLogger.Info("=== STARTING CreateFiltersPanel ===");
            
            // Create filters panel
            _filtersPanel = new WinForms.Panel
            {
                BackColor = System.Drawing.Color.FromArgb(248, 249, 250),
                BorderStyle = WinForms.BorderStyle.FixedSingle,
                Visible = true  // Ensure visibility for debugging
            };
            _filtersPanel.BringToFront();  // Bring to front to ensure it's visible
            this.Controls.Add(_filtersPanel);
            // DebugLogger.Info("_filtersPanel created and added to form");
            
            // Create filters splitter
            _filtersSplitter = new WinForms.Splitter
            {
                BackColor = System.Drawing.Color.Gray,
                Width = 3
            };
            this.Controls.Add(_filtersSplitter);
            DebugLogger.Info("_filtersSplitter created and added to form");
            
            // Populate filters content
            PopulateFiltersPanel();
            
            // DebugLogger.Info("=== CreateFiltersPanel COMPLETED ===");
        }

        private void PopulateFiltersPanel()
        {
            // DebugLogger.Info("=== STARTING PopulateFiltersPanel ===");
            
            // Title
            var title = new WinForms.Label
            {
                Text = "Filters",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51),
                Location = new System.Drawing.Point(10, 10),
                Size = new System.Drawing.Size(180, 25),
                AutoSize = false
            };
            _filtersPanel.Controls.Add(title);
                // DebugLogger.Info("Filters title label created");
            
            // ========================================================================
            // ⚠️  CRITICAL FILTER UI - DO NOT MODIFY WITHOUT USER CONSENT  ⚠️
            // ========================================================================
            // Filter List (like conVoid's filter list) - Enlarged height to -140
            // DO NOT CHANGE: Size and Anchor settings are critical for proper layout
            // ========================================================================
            var filterListBox = new WinForms.ListBox
            {
                Location = new System.Drawing.Point(10, 40),
                Size = new System.Drawing.Size(180, _filtersPanel.Height - 140), // Enlarged to -140
                BackColor = System.Drawing.Color.White,
                BorderStyle = WinForms.BorderStyle.FixedSingle,
                SelectionMode = WinForms.SelectionMode.MultiExtended,
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right
            };
            _filtersPanel.Controls.Add(filterListBox);
                // DebugLogger.Info("Filter list box created");
            
            // Add sample filters (like conVoid) via service so internal list is tracked
            
            // 🔍 DIAGNOSTIC: Count sleeves BEFORE SeedDefaultFilters
            try
            {
                if (_document != null)
                {
                    var sleevesBeforeSeed = new FilteredElementCollector(_document)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                        .Count();
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 BEFORE SeedDefaultFilters: {sleevesBeforeSeed} sleeves\n");
                }
            }
            catch { }
            
            _filterManagementService.SeedDefaultFilters(
                filterListBox,
                new System.Collections.Generic.List<string> { "Electrical", "Plumbing", "Ventilation" }
            );
            
            // ✅ PERSISTENCE: Load all saved filters from directory (shows filters saved in previous sessions)
            _filterManagementService.LoadAllSavedFilters(filterListBox);
            
            // 🔍 DIAGNOSTIC: Count sleeves AFTER SeedDefaultFilters
            try
            {
                if (_document != null)
                {
                    var sleevesAfterSeed = new FilteredElementCollector(_document)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                        .Count();
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 AFTER SeedDefaultFilters: {sleevesAfterSeed} sleeves\n");
                }
            }
            catch { }
            
            // DebugLogger.Info("Sample filters seeded via FilterManagementService");
            
            // Button panel positioned right below the filter list box
            var buttonPanel = new WinForms.Panel
            {
                Location = new System.Drawing.Point(10, _filtersPanel.Height - 70), // Positioned at bottom with 70px height
                Size = new System.Drawing.Size(180, 70), // Increased height for better spacing
                Anchor = WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right
            };
            _filtersPanel.Controls.Add(buttonPanel);
            
            // Row 1: New, Copy, Rename buttons (3 columns) - Simple text symbols, wider buttons
            var newFilterButton = new WinForms.Button
            {
                Text = "+", // Simple plus for New/Add
                Location = new System.Drawing.Point(5, 5),
                Size = new System.Drawing.Size(50, 30), // Wider to fill space better
                Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold),
                FlatStyle = WinForms.FlatStyle.Flat,
                BackColor = System.Drawing.Color.FromArgb(240, 240, 240),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51)
            };
            newFilterButton.FlatAppearance.BorderSize = 1;
            newFilterButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(200, 200, 200);
            buttonPanel.Controls.Add(newFilterButton);
                // DebugLogger.Info("New filter button created");
            
            var copyFilterButton = new WinForms.Button
            {
                Text = "⧉", // Copy/Duplicate symbol
                Location = new System.Drawing.Point(60, 5),
                Size = new System.Drawing.Size(50, 30), // Wider to fill space better
                Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold),
                FlatStyle = WinForms.FlatStyle.Flat,
                BackColor = System.Drawing.Color.FromArgb(240, 240, 240),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51)
            };
            copyFilterButton.FlatAppearance.BorderSize = 1;
            copyFilterButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(200, 200, 200);
            buttonPanel.Controls.Add(copyFilterButton);
                // DebugLogger.Info("Copy filter button created");
            
            var renameFilterButton = new WinForms.Button
            {
                Text = "✏", // Edit/Rename symbol
                Location = new System.Drawing.Point(115, 5),
                Size = new System.Drawing.Size(50, 30), // Wider to fill space better
                Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold),
                FlatStyle = WinForms.FlatStyle.Flat,
                BackColor = System.Drawing.Color.FromArgb(240, 240, 240),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51)
            };
            renameFilterButton.FlatAppearance.BorderSize = 1;
            renameFilterButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(200, 200, 200);
            buttonPanel.Controls.Add(renameFilterButton);
                // DebugLogger.Info("Rename filter button created");
            
            // Row 2: Delete, Save, Load buttons (3 columns) - Simple text symbols, wider buttons
            var deleteFilterButton = new WinForms.Button
            {
                Text = "×", // Delete symbol
                Location = new System.Drawing.Point(5, 40),
                Size = new System.Drawing.Size(50, 30), // Wider to fill space better
                Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold),
                FlatStyle = WinForms.FlatStyle.Flat,
                BackColor = System.Drawing.Color.FromArgb(240, 240, 240),
                ForeColor = System.Drawing.Color.FromArgb(200, 50, 50)
            };
            deleteFilterButton.FlatAppearance.BorderSize = 1;
            deleteFilterButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(200, 200, 200);
            buttonPanel.Controls.Add(deleteFilterButton);
                // DebugLogger.Info("Delete filter button created");
            
            var saveFilterButton = new WinForms.Button
            {
                Text = "↓", // Downwards arrow for Save (download/save action)
                Location = new System.Drawing.Point(60, 40),
                Size = new System.Drawing.Size(50, 30), // Wider to fill space better
                Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold),
                FlatStyle = WinForms.FlatStyle.Flat,
                BackColor = System.Drawing.Color.FromArgb(240, 240, 240),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51)
            };
            saveFilterButton.FlatAppearance.BorderSize = 1;
            saveFilterButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(200, 200, 200);
            buttonPanel.Controls.Add(saveFilterButton);
                // DebugLogger.Info("Save filter button created");
            
            var loadFilterButton = new WinForms.Button
            {
                Text = "↑", // Upwards arrow for Load (upload/load action)
                Location = new System.Drawing.Point(115, 40),
                Size = new System.Drawing.Size(50, 30), // Wider to fill space better
                Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold),
                FlatStyle = WinForms.FlatStyle.Flat,
                BackColor = System.Drawing.Color.FromArgb(240, 240, 240),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51)
            };
            loadFilterButton.FlatAppearance.BorderSize = 1;
            loadFilterButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(200, 200, 200);
            buttonPanel.Controls.Add(loadFilterButton);
                // DebugLogger.Info("Load filter button created");
            
            // Add tooltips for better UX
            var toolTip = new WinForms.ToolTip();
            toolTip.SetToolTip(newFilterButton, "New Filter");
            toolTip.SetToolTip(copyFilterButton, "Copy Filter");
            toolTip.SetToolTip(renameFilterButton, "Rename Filter");
            toolTip.SetToolTip(deleteFilterButton, "Delete Filter");
            toolTip.SetToolTip(saveFilterButton, "Save Filter");
            toolTip.SetToolTip(loadFilterButton, "Load Filter");
            
            // Add click event handlers for filter management buttons
            newFilterButton.Click += (s, e) => _filterManagementService.CreateNewFilter(filterListBox);
            copyFilterButton.Click += (s, e) => _filterManagementService.CopyFilter(filterListBox);
            renameFilterButton.Click += (s, e) => _filterManagementService.RenameFilter(filterListBox);
            deleteFilterButton.Click += (s, e) => _filterManagementService.DeleteFilter(filterListBox);
            saveFilterButton.Click += (s, e) => SaveFilterWithUIState(filterListBox);
            loadFilterButton.Click += (s, e) => _filterManagementService.LoadFilter(filterListBox);
            
            // Add event handler for filter selection to restore UI state
            filterListBox.SelectedIndexChanged += (s, e) => {
                DebugLogger.Info($"[FILTER_SELECTION] SelectedIndexChanged event fired - SelectedIndex={filterListBox.SelectedIndex}");
                if (filterListBox.SelectedItem != null)
                {
                    var selectedFilterName = filterListBox.SelectedItem.ToString();
                    DebugLogger.Info($"[FILTER_SELECTION] Filter selected: '{selectedFilterName}'");
                    
                    // Reset manual changes flag when user explicitly selects a filter
                    _userHasMadeManualChanges = false;
                    DebugLogger.Info("[FILTER_UI] User explicitly selected filter - resetting manual changes flag");
                    
                    // Auto-load the selected filter instead of showing file dialog
                    DebugLogger.Info($"[FILTER_SELECTION] Calling LoadFilterAuto for '{selectedFilterName}'");
                    var loadedFilter = _filterManagementService.LoadFilterAuto(selectedFilterName);
                    DebugLogger.Info($"[FILTER_SELECTION] LoadFilterAuto returned - filter is null: {loadedFilter == null}");
                    
                    if (loadedFilter != null)
                    {
                        DebugLogger.Info($"[FILTER_UI] Loaded filter '{selectedFilterName}' with {loadedFilter.SelectedMepCategoryNames?.Count ?? 0} MEP categories: {string.Join(", ", loadedFilter.SelectedMepCategoryNames ?? new List<string>())}");
                        // Pass the loaded filter directly to avoid loading it again
                        DebugLogger.Info($"[FILTER_SELECTION] Calling RestoreUIStateFromFilter");
                        RestoreUIStateFromFilter(loadedFilter);
                        DebugLogger.Info($"[FILTER_UI] Auto-loaded filter: {selectedFilterName}");
                    }
                    else
                    {
                        DebugLogger.Warning($"[FILTER_SELECTION] ⚠️ LoadFilterAuto returned NULL for '{selectedFilterName}'");
                    }
                }
                else
                {
                    DebugLogger.Info($"[FILTER_SELECTION] SelectedItem is NULL");
                }
            };
            
            // DebugLogger.Info("=== PopulateFiltersPanel COMPLETED ===");
        }
        private void CreateMainContentPanel()
        {
            // DebugLogger.Info("=== STARTING CreateMainContentPanel ===");
            
            // Create main content panel (center area)
            _mainContentPanel = new WinForms.Panel
            {
                BackColor = System.Drawing.Color.White,
                BorderStyle = WinForms.BorderStyle.FixedSingle,
                Dock = WinForms.DockStyle.None
            };
            this.Controls.Add(_mainContentPanel);
            
            // Create title for main content area
            var mainTitle = new WinForms.Label
            {
                Text = "Intersection Results",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51),
                Location = new System.Drawing.Point(10, 10),
                Size = new System.Drawing.Size(200, 25),
                AutoSize = false
            };
            _mainContentPanel.Controls.Add(mainTitle);
            
            // Create data grid for intersection results
            _intersectionDataGrid = new WinForms.DataGridView
            {
                Location = new System.Drawing.Point(10, 40),
                Size = new System.Drawing.Size(_mainContentPanel.Width - 20, _mainContentPanel.Height - 50),
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = WinForms.DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeColumnsMode = WinForms.DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = System.Drawing.Color.White,
                BorderStyle = WinForms.BorderStyle.Fixed3D,
                GridColor = System.Drawing.Color.LightGray,
                RowHeadersVisible = false,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F)
            };
            
            // Add columns to the data grid
            _intersectionDataGrid.Columns.Add("MepElement", "MEP Element");
            _intersectionDataGrid.Columns.Add("StructuralElement", "Structural Element");
            _intersectionDataGrid.Columns.Add("IntersectionPoint", "Intersection Point");
            _intersectionDataGrid.Columns.Add("Clearance", "Clearance (mm)");
            _intersectionDataGrid.Columns.Add("Status", "Status");
            
            _mainContentPanel.Controls.Add(_intersectionDataGrid);
            
            // DebugLogger.Info("=== CreateMainContentPanel COMPLETED ===");
        }

        // ========================================================================
        // ⚠️  CRITICAL UI LAYOUT - DO NOT MODIFY WITHOUT USER CONSENT  ⚠️
        // ========================================================================
        // This method creates the working 2x2 grid layout for the left section
        // Any changes to this method can break the entire UI layout
        // ========================================================================
        private void CreateFourSectionLayout()
        {
            // DebugLogger.Info("=== STARTING CreateFourSectionLayout ===");
            DebugLogger.Info($"_leftPanel size: {_leftPanel.Width}x{_leftPanel.Height}");

            // Top-Left Panel: Reference Elements (linked files)
            _topLeftPanel = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.Fill,
                BackColor = System.Drawing.Color.FromArgb(245, 245, 245),
                BorderStyle = WinForms.BorderStyle.FixedSingle,
                MinimumSize = new System.Drawing.Size(150, 0)  // ✅ Minimum width for file name section
            };
            _leftPanel.Controls.Add(_topLeftPanel);
                // DebugLogger.Info("_topLeftPanel created and added");

            // Horizontal Splitter (between top and bottom)
            _horizontalSplitter = new WinForms.Splitter
            {
                Dock = WinForms.DockStyle.Bottom,
                Height = 3,
                BackColor = System.Drawing.Color.Gray
            };
            _leftPanel.Controls.Add(_horizontalSplitter);
                // DebugLogger.Info("_horizontalSplitter created and added");

            // Bottom-Left Panel: Host Elements (linked files)
            _bottomLeftPanel = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.Bottom,
                Height = _leftPanel.Height / 2,
                BackColor = System.Drawing.Color.FromArgb(245, 245, 245),
                BorderStyle = WinForms.BorderStyle.FixedSingle,
                MinimumSize = new System.Drawing.Size(150, 0)  // ✅ Minimum width for file name section
            };
            _leftPanel.Controls.Add(_bottomLeftPanel);
                // DebugLogger.Info("_bottomLeftPanel created and added");

            // Now create the right side panels within the top and bottom panels
            CreateTopRightPanel();
            CreateBottomRightPanel();

            // Populate each section with content
            PopulateTopLeftSection();      // Reference Elements (linked files)
            PopulateTopRightSection();     // Reference Categories (MEP categories)
            PopulateBottomLeftSection();   // Host Elements (linked files)
            PopulateBottomRightSection();  // Host Categories (Revit categories)

            // DebugLogger.Info("=== CreateFourSectionLayout COMPLETED ===");
        }
        // ========================================================================
        // ⚠️  CRITICAL UI LAYOUT - DO NOT MODIFY WITHOUT USER CONSENT  ⚠️
        // ========================================================================
        // This method creates the top-right panel for MEP Categories
        // Uses DOCK to RIGHT for proper 2x2 grid positioning
        // ========================================================================
        private void CreateTopRightPanel()
        {
            // Top-Right Panel: Reference Categories (MEP categories) - DOCK to RIGHT
            _topRightPanel = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.Right,
                Width = 135,  // ✅ Reduced from 200 to 135 pixels (65px total reduction) to make room for file names
                MinimumSize = new System.Drawing.Size(110, 0),  // ✅ Minimum width for category section
                BackColor = System.Drawing.Color.FromArgb(250, 250, 250),
                BorderStyle = WinForms.BorderStyle.FixedSingle
            };
            _topLeftPanel.Controls.Add(_topRightPanel);

            // Vertical Splitter (between left and right in top section) - DOCK to RIGHT
            // ✅ CRITICAL: Splitter must be added AFTER the right panel for proper docking order
            _verticalSplitter = new WinForms.Splitter
            {
                Dock = WinForms.DockStyle.Right,
                Width = 5,  // ✅ Increased width for easier grabbing
                BackColor = System.Drawing.Color.Gray,
                Cursor = WinForms.Cursors.VSplit  // ✅ Show resize cursor
            };
            _topLeftPanel.Controls.Add(_verticalSplitter);
        }
        // ========================================================================
        // ⚠️  CRITICAL UI LAYOUT - DO NOT MODIFY WITHOUT USER CONSENT  ⚠️
        // ========================================================================
        // This method creates the bottom-right panel for Host Categories
        // Uses DOCK to RIGHT for proper 2x2 grid positioning
        // ========================================================================
        private void CreateBottomRightPanel()
        {
            // Bottom-Right Panel: Host Categories (Revit categories) - DOCK to RIGHT
            _bottomRightPanel = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.Right,
                Width = 135,  // ✅ Reduced from 200 to 135 pixels (65px total reduction) to make room for file names
                MinimumSize = new System.Drawing.Size(110, 0),  // ✅ Minimum width for category section
                BackColor = System.Drawing.Color.FromArgb(250, 250, 250),
                BorderStyle = WinForms.BorderStyle.FixedSingle
            };
            _bottomLeftPanel.Controls.Add(_bottomRightPanel);

            // Vertical Splitter (between left and right in bottom section) - DOCK to RIGHT
            // ✅ CRITICAL: Splitter must be added AFTER the right panel for proper docking order
            var bottomVerticalSplitter = new WinForms.Splitter
            {
                Dock = WinForms.DockStyle.Right,
                Width = 5,  // ✅ Increased width for easier grabbing
                BackColor = System.Drawing.Color.Gray,
                Cursor = WinForms.Cursors.VSplit  // ✅ Show resize cursor
            };
            _bottomLeftPanel.Controls.Add(bottomVerticalSplitter);
        }

        private void PopulateTopLeftSection()
        {
            // Title: "Reference Elements" 
            var title = new WinForms.Label
            {
                Text = "Reference Elements",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51),
                Location = new System.Drawing.Point(5, 2),
                Size = new System.Drawing.Size(150, 30),
                AutoSize = false
            };
            _topLeftPanel.Controls.Add(title);

            // Main list for MEP files (Reference Elements)
            var referenceFilesListBox = new WinForms.CheckedListBox
            {
                Location = new System.Drawing.Point(5, 40),
                Size = new System.Drawing.Size(_topLeftPanel.Width - 10, (_topLeftPanel.Height - 25) / 2 - 5),
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right,
                CheckOnClick = true,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular)
            };
            _topLeftPanel.Controls.Add(referenceFilesListBox);

            // "Other Files" section for Architecture/Structural files
            var otherFilesLabel = new WinForms.Label
            {
                Text = "Other Files:",
                Location = new System.Drawing.Point(5, referenceFilesListBox.Bottom + 5),
                Size = new System.Drawing.Size(200, 15),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.Gray
            };
            _topLeftPanel.Controls.Add(otherFilesLabel);

            var otherReferenceFilesListBox = new WinForms.CheckedListBox
            {
                Location = new System.Drawing.Point(5, otherFilesLabel.Bottom + 2),
                Size = new System.Drawing.Size(_topLeftPanel.Width - 10, (_topLeftPanel.Height - 25) / 2 - 5),
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right,
                CheckOnClick = true,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular)
            };
            _topLeftPanel.Controls.Add(otherReferenceFilesListBox);

            // ✅ AUTO-SAVE: Add ItemCheck handlers to save UI state automatically when reference files are checked/unchecked
            referenceFilesListBox.ItemCheck += (sender, e) => {
                System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
                timer.Interval = 100; // Small delay to avoid issues during check operation
                timer.Tick += (s, args) => {
                    timer.Stop();
                    timer.Dispose();
                    SaveUIStateDirectly();
                };
                timer.Start();
            };

            otherReferenceFilesListBox.ItemCheck += (sender, e) => {
                System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
                timer.Interval = 100; // Small delay to avoid issues during check operation
                timer.Tick += (s, args) => {
                    timer.Stop();
                    timer.Dispose();
                    SaveUIStateDirectly();
                };
                timer.Start();
            };

            // Always add the active document as a reference element (available by default)
            if (_activeDocument != null)
            {
                string activeDocName = _activeDocument.Title;
                referenceFilesListBox.Items.Add($"{activeDocName} (Active Document)", false);
                System.Diagnostics.Debug.WriteLine($"Added active document '{activeDocName}' to reference elements");
            }
            
            // Get the other files list box
            var listBoxes = _topLeftPanel.Controls.OfType<WinForms.CheckedListBox>().ToList();
            var otherRefFilesListBox = listBoxes.Skip(1).FirstOrDefault();

            // Add files to appropriate sections
            if (_linkedFiles.Count > 0 && _linkedFileService != null)
            {
                foreach (var file in _linkedFiles)
                {
                    // ✅ FIX: Display only filename - no .rvt extension, element count, or location info
                    string displayText = System.IO.Path.GetFileNameWithoutExtension(file.FileName);
                    if (!file.IsLoaded)
                    {
                        displayText += " [NOT LOADED]";
                    }
                    
                    // ✅ Show MEP files in main list
                    bool isAvailableForReference = IsFileAvailableForReference(file.FileType);
                    if (isAvailableForReference)
                    {
                        referenceFilesListBox.Items.Add(displayText, false);
                    }
                    else
                    {
                        // ✅ "Other Files" section: Only show files that are NOT in Reference Elements main list AND NOT in Host Elements main list
                        bool isAvailableForHost = IsFileAvailableForHost(file.FileType);
                        if (!isAvailableForHost)
                        {
                            // File is not in either main section → add to "Other Files"
                            otherRefFilesListBox?.Items.Add(displayText, false);
                        }
                        // If file is available for Host, skip it (it will be shown in Host Elements main list)
                    }
                }
                System.Diagnostics.Debug.WriteLine($"Added files to reference elements sections");
            }
        }

        private void PopulateTopRightSection()
        {
            // Title: "MEP Categories" - smaller and positioned at top
            var title = new WinForms.Label
            {
                Text = "MEP Categories",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51),
                Location = new System.Drawing.Point(5, 2),
                Size = new System.Drawing.Size(100, 25),
                AutoSize = false
            };
            _topRightPanel.Controls.Add(title);

            // ListBox for MEP categories with checkboxes - move down more
            var referenceCategoriesListBox = new WinForms.CheckedListBox
            {
                Location = new System.Drawing.Point(5, 35),  // Change from 30 to 35 (more space)
                Size = new System.Drawing.Size(_topRightPanel.Width - 10, _topRightPanel.Height - 40),  // Change from 35 to 40
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right,
                CheckOnClick = true,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular)
            };
            _topRightPanel.Controls.Add(referenceCategoriesListBox);

            // Add event handler for category selection changes
            referenceCategoriesListBox.ItemCheck += (sender, e) => {
                // Mark that user has made manual changes
                _userHasMadeManualChanges = true;
                DebugLogger.Info("[FILTER_UI] User manually changed MEP category selection - marking as manual change");
                
                // Use a timer to delay the update to avoid issues during the check operation
                System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
                timer.Interval = 10;
                timer.Tick += (s, args) => {
                    timer.Stop();
                    timer.Dispose();
                    UpdateMepTypeBasedOnSelection();
                    // ✅ AUTO-SAVE: Save UI state automatically when MEP category selection changes
                    SaveUIStateDirectly();
                };
                timer.Start();
            };

            // Always show all 4 MEP categories regardless of linked files
            var allMepCategories = new List<string> { "Ducts", "Duct Accessories", "Cable Trays", "Pipes" };
            
            foreach (var category in allMepCategories)
            {
                referenceCategoriesListBox.Items.Add(category, false);
            }
            
            System.Diagnostics.Debug.WriteLine($"Populated MEP categories with {allMepCategories.Count} standard categories");
        }





        private void UpdateMepTypeBasedOnSelection()
        {
            try
            {
                // Get selected categories from the MEP Categories listbox
                var selectedCategories = new List<string>();
                var mepCategoriesListBox = _topRightPanel.Controls.OfType<WinForms.CheckedListBox>().FirstOrDefault();
                
                if (mepCategoriesListBox != null)
                {
                    for (int i = 0; i < mepCategoriesListBox.Items.Count; i++)
                    {
                        if (mepCategoriesListBox.GetItemChecked(i))
                        {
                            selectedCategories.Add(mepCategoriesListBox.Items[i].ToString());
                        }
                    }
                }

                if (selectedCategories.Count == 0)
                {
                    // No categories selected - grey out and clear
                    _mepTypeCombo.Enabled = false;
                    _mepTypeCombo.BackColor = System.Drawing.Color.LightGray;
                    _mepTypeCombo.Items.Clear();
                    _mepTypeCombo.Items.Add("<Select>");
                    _mepTypeCombo.SelectedIndex = 0;
                }
                else if (selectedCategories.Count == 1)
                {
                    // Single category selected - grey out and show the category
                    _mepTypeCombo.Enabled = false;
                    _mepTypeCombo.BackColor = System.Drawing.Color.LightGray;
                    _mepTypeCombo.Items.Clear();
                    _mepTypeCombo.Items.Add(selectedCategories[0]);
                    _mepTypeCombo.SelectedIndex = 0;
                    
                    // Update clearance visibility for single selection
                    UpdateClearanceVisibilityForCategory(selectedCategories[0]);
            }
            else
            {
                    // Multiple categories selected - enable and show only selected categories
                    _mepTypeCombo.Enabled = true;
                    _mepTypeCombo.BackColor = System.Drawing.Color.White;
                    _mepTypeCombo.Items.Clear();
                    _mepTypeCombo.Items.AddRange(selectedCategories.ToArray());
                    _mepTypeCombo.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error updating MEP Type based on selection: {ex.Message}");
            }
        }

        private void PopulateBottomLeftSection()
        {
            // Title: "Host Elements"
            var title = new WinForms.Label
            {
                Text = "Host Elements",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51),
                Location = new System.Drawing.Point(5, 2),
                Size = new System.Drawing.Size(150, 30),
                AutoSize = false
            };
            _bottomLeftPanel.Controls.Add(title);

            // Main list for Architecture/Structural files (Host Elements)
            var hostFilesListBox = new WinForms.CheckedListBox
            {
                Location = new System.Drawing.Point(5, 30),
                Size = new System.Drawing.Size(_bottomLeftPanel.Width - 10, (_bottomLeftPanel.Height - 25) / 2 - 5),
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right,
                CheckOnClick = true,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular)
            };
            _bottomLeftPanel.Controls.Add(hostFilesListBox);

            // "Other Files" section for MEP files
            var otherHostFilesLabel = new WinForms.Label
            {
                Text = "Other Files:",
                Location = new System.Drawing.Point(5, hostFilesListBox.Bottom + 5),
                Size = new System.Drawing.Size(200, 15),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.Gray
            };
            _bottomLeftPanel.Controls.Add(otherHostFilesLabel);

            var otherHostFilesListBox = new WinForms.CheckedListBox
            {
                Location = new System.Drawing.Point(5, otherHostFilesLabel.Bottom + 2),
                Size = new System.Drawing.Size(_bottomLeftPanel.Width - 10, (_bottomLeftPanel.Height - 25) / 2 - 5),
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right,
                CheckOnClick = true,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular)
            };
            _bottomLeftPanel.Controls.Add(otherHostFilesListBox);

            // ✅ AUTO-SAVE: Add ItemCheck handlers to save UI state automatically when host files are checked/unchecked
            hostFilesListBox.ItemCheck += (sender, e) => {
                System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
                timer.Interval = 100; // Small delay to avoid issues during check operation
                timer.Tick += (s, args) => {
                    timer.Stop();
                    timer.Dispose();
                    SaveUIStateDirectly();
                };
                timer.Start();
            };

            otherHostFilesListBox.ItemCheck += (sender, e) => {
                System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
                timer.Interval = 100; // Small delay to avoid issues during check operation
                timer.Tick += (s, args) => {
                    timer.Stop();
                    timer.Dispose();
                    SaveUIStateDirectly();
                };
                timer.Start();
            };

            // ✅ FIX: Active Document should ONLY be in Reference Elements section, NOT in Host Elements
            // Removed: Active Document from Host Elements section
            
            // Get the other files list box
            var hostListBoxes = _bottomLeftPanel.Controls.OfType<WinForms.CheckedListBox>().ToList();
            var otherHostFilesListBox2 = hostListBoxes.Skip(1).FirstOrDefault();

            // Add files to appropriate sections
            if (_linkedFiles.Count > 0 && _linkedFileService != null)
            {
                foreach (var file in _linkedFiles)
                {
                    // ✅ FIX: Display only filename - no .rvt extension, element count, or location info
                    string displayText = System.IO.Path.GetFileNameWithoutExtension(file.FileName);
                    if (!file.IsLoaded)
                    {
                        displayText += " [NOT LOADED]";
                    }
                    
                    // ✅ Show Architecture/Structural files in main list
                    bool isAvailableForHost = IsFileAvailableForHost(file.FileType);
                    if (isAvailableForHost)
                    {
                        hostFilesListBox.Items.Add(displayText, false);
                    }
                    else
                    {
                        // ✅ "Other Files" section: Only show files that are NOT in Host Elements main list AND NOT in Reference Elements main list
                        bool isAvailableForReference = IsFileAvailableForReference(file.FileType);
                        if (!isAvailableForReference)
                        {
                            // File is not in either main section → add to "Other Files"
                            otherHostFilesListBox2?.Items.Add(displayText, false);
                        }
                        // If file is available for Reference, skip it (it will be shown in Reference Elements main list)
                    }
                }
                System.Diagnostics.Debug.WriteLine($"Added files to host elements sections");
            }
        }

        private void PopulateBottomRightSection()
        {
            // Title: "Host Categories" - smaller and positioned at top
            var title = new WinForms.Label
            {
                Text = "Host Categories",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51),
                Location = new System.Drawing.Point(5, 2),
                Size = new System.Drawing.Size(100, 16),
                AutoSize = false
            };
            _bottomRightPanel.Controls.Add(title);

            // Horizontal Openings Section - starts right after title
            var horizontalLabel = new WinForms.Label
            {
                Text = "Horizontal Openings:",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(0, 100, 200),
                Location = new System.Drawing.Point(5, 20),
                Size = new System.Drawing.Size(100, 14),
                AutoSize = false
            };
            _bottomRightPanel.Controls.Add(horizontalLabel);
            horizontalLabel.Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right;

            // ListBox for horizontal host categories
            var horizontalCategoriesListBox = new WinForms.CheckedListBox
            {
                Location = new System.Drawing.Point(5, 36),
                Size = new System.Drawing.Size(_bottomRightPanel.Width - 10, 70),
                CheckOnClick = true,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular)
            };
            _bottomRightPanel.Controls.Add(horizontalCategoriesListBox);
            _horizontalCategoriesListBox = horizontalCategoriesListBox; // ✅ Store reference for UI state persistence

            // ✅ AUTO-SAVE: Add ItemCheck handler to save UI state automatically when horizontal host categories are checked/unchecked
            horizontalCategoriesListBox.ItemCheck += (sender, e) => {
                System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
                timer.Interval = 100; // Small delay to avoid issues during check operation
                timer.Tick += (s, args) => {
                    timer.Stop();
                    timer.Dispose();
                    SaveUIStateDirectly();
                };
                timer.Start();
            };

            // Always show Walls and Structural Framing categories regardless of linked files
            horizontalCategoriesListBox.Items.Add("Walls", false);
            horizontalCategoriesListBox.Items.Add("Structural Framing", false);
            System.Diagnostics.Debug.WriteLine("Added Walls and Structural Framing to horizontal host categories");

            // Vertical Openings Section - positioned below horizontal
            var verticalLabel = new WinForms.Label
            {
                Text = "Vertical Openings:",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(200, 100, 0),
                Location = new System.Drawing.Point(5, 100),
                Size = new System.Drawing.Size(100, 14),
                AutoSize = false
            };
            _bottomRightPanel.Controls.Add(verticalLabel);
            verticalLabel.Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right;

            // ListBox for vertical host categories
            var verticalCategoriesListBox = new WinForms.CheckedListBox
            {
                Location = new System.Drawing.Point(5, 116),
                Size = new System.Drawing.Size(_bottomRightPanel.Width - 10, 70),
                CheckOnClick = true,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular)
            };
            _bottomRightPanel.Controls.Add(verticalCategoriesListBox);
            _verticalCategoriesListBox = verticalCategoriesListBox; // ✅ Store reference for UI state persistence

            // ✅ AUTO-SAVE: Add ItemCheck handler to save UI state automatically when vertical host categories are checked/unchecked
            verticalCategoriesListBox.ItemCheck += (sender, e) => {
                System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
                timer.Interval = 100; // Small delay to avoid issues during check operation
                timer.Tick += (s, args) => {
                    timer.Stop();
                    timer.Dispose();
                    SaveUIStateDirectly();
                };
                timer.Start();
            };

            // Add vertical categories based on available linked files
            if (_linkedFiles.Count > 0 && _linkedFileService != null)
            {
                var hostFiles = _linkedFileService.GetHostElementFiles(_linkedFiles);
                var verticalCategories = new HashSet<string>();

                foreach (var file in hostFiles)
                {
                    var categories = LinkedFileDetectionService.GetAvailableHostCategories(file.FileType);
                    foreach (var category in categories)
                    {
                        var displayName = LinkedFileDetectionService.GetHostCategoryDisplayName(category);
                        if (displayName == "Floors" || displayName == "Ceilings")
                        {
                            verticalCategories.Add(displayName);
                        }
                    }
                }

                foreach (var category in verticalCategories.OrderBy(c => c))
                {
                    verticalCategoriesListBox.Items.Add(category, false);
                }
                System.Diagnostics.Debug.WriteLine($"Added {verticalCategories.Count} vertical host categories");
            }
            else
            {
                // Show message when no linked files
                var noDataLabel = new WinForms.Label
                {
                    Text = "No linked files.\nVertical openings require architectural/structural files.",
                    Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Italic),
                    ForeColor = System.Drawing.Color.Gray,
                    Location = new System.Drawing.Point(10, 120),
                    Size = new System.Drawing.Size(_bottomRightPanel.Width - 20, 40),
                    AutoSize = false,
                    TextAlign = System.Drawing.ContentAlignment.MiddleCenter
                };
                _bottomRightPanel.Controls.Add(noDataLabel);
            }

         _hostCategoriesInitialized =
            (_horizontalCategoriesListBox != null && _horizontalCategoriesListBox.Items.Count > 0) ||
            (_verticalCategoriesListBox != null && _verticalCategoriesListBox.Items.Count > 0);

            if (!_hostCategoriesInitialized && !DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info("[UI-STATE] Host categories not initialized yet; UI state restore will be deferred");
            }
        }

        private void InitializeRightPanel()
        {
            // Title
            var rightTitle = new WinForms.Label
            {
                Text = "Opening Configuration & Clearances",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51),
                Location = new System.Drawing.Point(10, 10),
                Size = new System.Drawing.Size(_rightPanel.Width - 20, 25),
                AutoSize = false,
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right
            };
            _rightPanel.Controls.Add(rightTitle);

            // MEP Type Selection
            var mepTypeLabel = new WinForms.Label
            {
                Text = "MEP Type:",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51),
                Location = new System.Drawing.Point(10, 45),
                Size = new System.Drawing.Size(100, 20),
                AutoSize = false
            };
            _rightPanel.Controls.Add(mepTypeLabel);

            _mepTypeCombo = new WinForms.ComboBox
            {
                Location = new System.Drawing.Point(120, 45),
                Size = new System.Drawing.Size(_rightPanel.Width - 130, 20),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular),
                DropDownStyle = WinForms.ComboBoxStyle.DropDownList,
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right,
                Visible = true, // Always visible
                Enabled = false, // Initially disabled (greyed out)
                BackColor = System.Drawing.Color.LightGray // Grey background when disabled
            };
            _mepTypeCombo.Items.AddRange(new[] { "Pipe", "Duct", "Duct Accessories", "Duct Fittings", "Cable Trays", "Conduit" });
            _mepTypeCombo.SelectedIndex = 0;
            
            // Add event handler to update clearance visibility when MEP type changes
            _mepTypeCombo.SelectedIndexChanged += (s, e) => 
            {
                // ✅ FIX: Update both clearance visibility AND discipline prefix
                // This ensures discipline prefix textbox shows the correct prefix for the selected MEP type
                UpdateClearanceVisibility();
                UpdateDisciplinePrefix(); // ⚠️ CRITICAL: Update discipline prefix when MEP type changes
            };
            
            _rightPanel.Controls.Add(_mepTypeCombo);

            // Create clearance panels
            
            // 🔍 DIAGNOSTIC: Count sleeves BEFORE CreateClearancePanels
            try
            {
                if (_document != null)
                {
                    var sleevesBeforeClearance = new FilteredElementCollector(_document)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                        .Count();
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 BEFORE CreateClearancePanels: {sleevesBeforeClearance} sleeves\n");
                }
            }
            catch { }
            
            CreateClearancePanels();
            
            // 🔍 DIAGNOSTIC: Count sleeves AFTER CreateClearancePanels
            try
            {
                if (_document != null)
                {
                    var sleevesAfterClearance = new FilteredElementCollector(_document)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                        .Count();
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 AFTER CreateClearancePanels: {sleevesAfterClearance} sleeves\n");
                }
            }
            catch { }

            // Parameter service section removed - now available as separate dialog
            
            // ✅ INTEGRATED SETTINGS: Add Settings section below Clearance panels
            CreateSettingsSection();
            
            // Ensure correct panel visible at startup
            UpdateClearanceVisibility();
            // Initialize discipline prefix based on default MEP Type selection
            UpdateDisciplinePrefix(); // ⚠️ NEW: Set initial discipline prefix
        }
        /// <summary>
        /// ✅ INTEGRATED SETTINGS: Create Settings section in right panel below Clearance
        /// Shows "Adopt to modified document" and all Limits settings
        /// </summary>
        private void CreateSettingsSection()
        {
            // Settings start below Clearance panels (Y=75 + 110 height = 185, add 10px gap = 195)
            int settingsStartY = 195;
            
            // Settings GroupBox
            var settingsGroupBox = new WinForms.GroupBox
            {
                Text = "Settings",
                Location = new System.Drawing.Point(10, settingsStartY),
                Size = new System.Drawing.Size(_rightPanel.Width - 20, 310), // Increased height for Force Full Detection checkbox
                BackColor = System.Drawing.Color.FromArgb(248, 249, 250),
                // ✅ NOTE: GroupBox doesn't have BorderStyle property - it has a border by default
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right
            };
            _rightPanel.Controls.Add(settingsGroupBox);
            
            int yPos = 25;
            
            // ✅ VISIBLE: Adopt to modified document checkbox
            var adoptModificationLabel = new WinForms.Label
            {
                Text = "Adopt to modified document:",
                Location = new System.Drawing.Point(15, yPos),
                Size = new System.Drawing.Size(350, 20),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular)
            };
            settingsGroupBox.Controls.Add(adoptModificationLabel);
            
            _enableThreePointValidationCheckBox = new WinForms.CheckBox
            {
                Text = "",
                Location = new System.Drawing.Point(370, yPos),
                Size = new System.Drawing.Size(20, 20),
                Checked = true // Default enabled
            };
            settingsGroupBox.Controls.Add(_enableThreePointValidationCheckBox);
            yPos += 30;
            
            // ✅ NEW: Force Full Detection Mode checkbox
            var forceDetectionLabel = new WinForms.Label
            {
                Text = "Force Full Detection Mode:",
                Location = new System.Drawing.Point(15, yPos),
                Size = new System.Drawing.Size(350, 20),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular)
            };
            settingsGroupBox.Controls.Add(forceDetectionLabel);
            
            _forceDetectionModeCheckBox = new WinForms.CheckBox
            {
                Text = "",
                Location = new System.Drawing.Point(370, yPos),
                Size = new System.Drawing.Size(20, 20),
                Checked = false // Default disabled
            };
            settingsGroupBox.Controls.Add(_forceDetectionModeCheckBox);
            yPos += 30;
            
            // ✅ ALL LIMITS SETTINGS VISIBLE:
            
            // Ignore openings smaller than
            var ignoreSmallLabel = new WinForms.Label
            {
                Text = "Ignore openings smaller than (mm):",
                Location = new System.Drawing.Point(15, yPos),
                Size = new System.Drawing.Size(350, 20),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular)
            };
            settingsGroupBox.Controls.Add(ignoreSmallLabel);
            
            _ignoreOpeningsSmallerThanTextBox = new WinForms.TextBox
            {
                Text = "0.1",
                Location = new System.Drawing.Point(370, yPos),
                Size = new System.Drawing.Size(50, 20)
            };
            settingsGroupBox.Controls.Add(_ignoreOpeningsSmallerThanTextBox);
            yPos += 25;
            
            // Round openings rectangular
            var roundRectLabel = new WinForms.Label
            {
                Text = "Round → Rectangular if diameter > (mm):",
                Location = new System.Drawing.Point(15, yPos),
                Size = new System.Drawing.Size(350, 20),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular)
            };
            settingsGroupBox.Controls.Add(roundRectLabel);
            
            _roundOpeningsRectangularTextBox = new WinForms.TextBox
            {
                Text = "200",
                Location = new System.Drawing.Point(370, yPos),
                Size = new System.Drawing.Size(50, 20)
            };
            settingsGroupBox.Controls.Add(_roundOpeningsRectangularTextBox);
            yPos += 25;
            
            // Join openings distance
            var joinDistanceLabel = new WinForms.Label
            {
                Text = "Join openings if distance < (mm):",
                Location = new System.Drawing.Point(15, yPos),
                Size = new System.Drawing.Size(350, 20),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular)
            };
            settingsGroupBox.Controls.Add(joinDistanceLabel);
            
            _joinOpeningsDistanceTextBox = new WinForms.TextBox
            {
                Text = "200",
                Location = new System.Drawing.Point(370, yPos),
                Size = new System.Drawing.Size(50, 20)
            };
            settingsGroupBox.Controls.Add(_joinOpeningsDistanceTextBox);
            yPos += 25;
            
            // Round opening sizes
            var roundSizesLabel = new WinForms.Label
            {
                Text = "Round sizes to nearest (mm):",
                Location = new System.Drawing.Point(15, yPos),
                Size = new System.Drawing.Size(350, 20),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular)
            };
            settingsGroupBox.Controls.Add(roundSizesLabel);
            
            _roundingValueTextBox = new WinForms.TextBox
            {
                Text = "5",
                Location = new System.Drawing.Point(370, yPos),
                Size = new System.Drawing.Size(50, 20)
            };
            settingsGroupBox.Controls.Add(_roundingValueTextBox);
            yPos += 25;
            
            // Always round up
            var roundUpLabel = new WinForms.Label
            {
                Text = "Always round up:",
                Location = new System.Drawing.Point(15, yPos),
                Size = new System.Drawing.Size(350, 20),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular)
            };
            settingsGroupBox.Controls.Add(roundUpLabel);
            
            _roundAlwaysUpCheckBox = new WinForms.CheckBox
            {
                Text = "",
                Location = new System.Drawing.Point(370, yPos),
                Size = new System.Drawing.Size(20, 20),
                Checked = false
            };
            settingsGroupBox.Controls.Add(_roundAlwaysUpCheckBox);
            yPos += 25;
            
            // Ignore walls if thickness below
            var minWallThicknessLabel = new WinForms.Label
            {
                Text = "Ignore walls if thickness < (mm):",
                Location = new System.Drawing.Point(15, yPos),
                Size = new System.Drawing.Size(350, 20),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular)
            };
            settingsGroupBox.Controls.Add(minWallThicknessLabel);
            
            _minWallThicknessTextBox = new WinForms.TextBox
            {
                Text = "0",
                Location = new System.Drawing.Point(370, yPos),
                Size = new System.Drawing.Size(50, 20)
            };
            settingsGroupBox.Controls.Add(_minWallThicknessTextBox);
            yPos += 25;
            
            // Ignore architectural floors
            var ignoreArchFloorsLabel = new WinForms.Label
            {
                Text = "Ignore architectural floors:",
                Location = new System.Drawing.Point(15, yPos),
                Size = new System.Drawing.Size(350, 20),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular)
            };
            settingsGroupBox.Controls.Add(ignoreArchFloorsLabel);
            
            _ignoreArchitecturalFloorsCheckBox = new WinForms.CheckBox
            {
                Text = "",
                Location = new System.Drawing.Point(370, yPos),
                Size = new System.Drawing.Size(20, 20),
                Checked = false
            };
            settingsGroupBox.Controls.Add(_ignoreArchitecturalFloorsCheckBox);
            
            // Load settings from file
            LoadSettingsToRightPanel();
            
            // Auto-save settings when changed
            _enableThreePointValidationCheckBox.CheckedChanged += (s, e) => SaveSettingsFromRightPanel();
            _ignoreOpeningsSmallerThanTextBox.TextChanged += (s, e) => SaveSettingsFromRightPanel();
            _roundOpeningsRectangularTextBox.TextChanged += (s, e) => SaveSettingsFromRightPanel();
            _joinOpeningsDistanceTextBox.TextChanged += (s, e) => SaveSettingsFromRightPanel();
            _roundingValueTextBox.TextChanged += (s, e) => SaveSettingsFromRightPanel();
            _roundAlwaysUpCheckBox.CheckedChanged += (s, e) => SaveSettingsFromRightPanel();
            _minWallThicknessTextBox.TextChanged += (s, e) => SaveSettingsFromRightPanel();
            _ignoreArchitecturalFloorsCheckBox.CheckedChanged += (s, e) => SaveSettingsFromRightPanel();
            _forceDetectionModeCheckBox.CheckedChanged += (s, e) => SaveSettingsFromRightPanel();
        }
        
        /// <summary>
        /// Load settings from file into right panel controls
        /// </summary>
        private void LoadSettingsToRightPanel()
        {
            try
            {
                var settingsService = new SettingsService();
                var settings = settingsService.LoadSettings();
                
                if (_enableThreePointValidationCheckBox != null)
                    _enableThreePointValidationCheckBox.Checked = settings.EnableThreePointValidation;
                if (_forceDetectionModeCheckBox != null)
                    _forceDetectionModeCheckBox.Checked = settings.ForceDetectionMode;
                if (_ignoreOpeningsSmallerThanTextBox != null)
                    _ignoreOpeningsSmallerThanTextBox.Text = settings.IgnoreOpeningsSmallerThan.ToString();
                if (_roundOpeningsRectangularTextBox != null)
                    _roundOpeningsRectangularTextBox.Text = settings.RoundOpeningsRectangular.ToString();
                if (_joinOpeningsDistanceTextBox != null)
                    _joinOpeningsDistanceTextBox.Text = settings.JoinOpeningsDistance.ToString();
                if (_roundingValueTextBox != null)
                    _roundingValueTextBox.Text = settings.RoundingValue.ToString();
                if (_roundAlwaysUpCheckBox != null)
                    _roundAlwaysUpCheckBox.Checked = settings.RoundAlwaysUp;
                if (_minWallThicknessTextBox != null)
                    _minWallThicknessTextBox.Text = settings.MinWallThickness.ToString();
                if (_ignoreArchitecturalFloorsCheckBox != null)
                    _ignoreArchitecturalFloorsCheckBox.Checked = settings.IgnoreArchitecturalFloors;
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[EmergencyMainDialog] Error loading settings to right panel: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Save settings from right panel controls to file
        /// </summary>
        private void SaveSettingsFromRightPanel()
        {
            try
            {
                var settingsService = new SettingsService();
                var settings = settingsService.LoadSettings();
                
                if (_enableThreePointValidationCheckBox != null)
                    settings.EnableThreePointValidation = _enableThreePointValidationCheckBox.Checked;
                if (_forceDetectionModeCheckBox != null)
                    settings.ForceDetectionMode = _forceDetectionModeCheckBox.Checked;
                if (_ignoreOpeningsSmallerThanTextBox != null && double.TryParse(_ignoreOpeningsSmallerThanTextBox.Text, out double ignoreSmall))
                    settings.IgnoreOpeningsSmallerThan = ignoreSmall;
                if (_roundOpeningsRectangularTextBox != null && double.TryParse(_roundOpeningsRectangularTextBox.Text, out double roundRect))
                    settings.RoundOpeningsRectangular = roundRect;
                if (_joinOpeningsDistanceTextBox != null && double.TryParse(_joinOpeningsDistanceTextBox.Text, out double joinDist))
                    settings.JoinOpeningsDistance = joinDist;
                if (_roundingValueTextBox != null && double.TryParse(_roundingValueTextBox.Text, out double roundValue))
                    settings.RoundingValue = roundValue;
                if (_roundAlwaysUpCheckBox != null)
                    settings.RoundAlwaysUp = _roundAlwaysUpCheckBox.Checked;
                if (_minWallThicknessTextBox != null && double.TryParse(_minWallThicknessTextBox.Text, out double minWallThickness))
                    settings.MinWallThickness = minWallThickness;
                if (_ignoreArchitecturalFloorsCheckBox != null)
                    settings.IgnoreArchitecturalFloors = _ignoreArchitecturalFloorsCheckBox.Checked;
                
                settingsService.SaveSettings(settings);
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[EmergencyMainDialog] Error saving settings from right panel: {ex.Message}");
            }
        }



        private void UpdateClearanceVisibility()
        {
            var mep = _mepTypeCombo.SelectedItem?.ToString() ?? string.Empty;
            UpdateClearanceVisibilityForCategory(mep);
        }
        /// <summary>
        /// Updates clearance panel visibility based on selected MEP category
        /// 
        /// ⚠️ CRITICAL FEATURE - DO NOT MODIFY WITHOUT APPROVAL ⚠️
        /// This method controls which clearance panel is visible:
        /// - "Cable Trays" → Shows _cableTrayPanel (Top Side + Other Sides)
        /// - "Duct Accessories" → Shows _damperPanel (MEP Connector Side + Other Sides)
        /// - "Ducts" → Shows _clearancePanel (standard clearance)
        /// - "Pipes" → Shows _pipePanel (clearance + opening type selection)
        /// 
        /// String matching is case-insensitive and must match exactly with combo box items.
        /// </summary>
        private void UpdateClearanceVisibilityForCategory(string category)
        {
            // ✅ SAVE: Save current clearance values before switching categories
            if (!string.IsNullOrEmpty(_currentCategory))
            {
                SaveCurrentClearanceValues();
            }
            
            // Hide all
            _clearancePanel.Visible = false;
            _cableTrayPanel.Visible = false;
            _damperPanel.Visible = false;
            _pipePanel.Visible = false;

            // Update current category
            _currentCategory = category;

            if (category.Equals("Cable Trays", StringComparison.OrdinalIgnoreCase))
            {
                _cableTrayPanel.Visible = true;
                RestoreClearanceValues("Cable Trays");
            }
            else if (category.Equals("Duct Accessories", StringComparison.OrdinalIgnoreCase))
            {
                _damperPanel.Visible = true;
                RestoreClearanceValues("Duct Accessories");
            }
            else if (category.Equals("Ducts", StringComparison.OrdinalIgnoreCase))
            {
                _clearancePanel.Visible = true;
                RestoreClearanceValues("Ducts");
            }
            else if (category.Equals("Pipes", StringComparison.OrdinalIgnoreCase) || category.Equals("Pipe", StringComparison.OrdinalIgnoreCase))
            {
                _pipePanel.Visible = true;
                // Normalize to "Pipes" for consistency
                RestoreClearanceValues("Pipes");
            }
            else
            {
                _clearancePanel.Visible = true; // default
                RestoreClearanceValues("Default");
            }
        }
        /// <summary>
        /// Updates the discipline prefix textbox based on the selected MEP Type dropdown
        /// ⚠️ NEW METHOD - Updates discipline prefix dynamically when MEP Type changes
        /// </summary>
        private void UpdateDisciplinePrefix()
        {
            try
            {
                if (_disciplinePrefixTextBox == null) return;

                var selectedMepType = _mepTypeCombo?.SelectedItem?.ToString() ?? string.Empty;
                
                // Map MEP Type dropdown values to discipline prefixes with proper fallback
                // Use MepCategoryConstants.Normalize() to handle singular/plural variations
                string normalizedCategory = MepCategoryConstants.Normalize(selectedMepType);
                string disciplinePrefix = normalizedCategory switch
                {
                    MepCategoryConstants.DUCTS => _categoryPrefixes.ContainsKey(MepCategoryConstants.DUCTS) 
                        ? _categoryPrefixes[MepCategoryConstants.DUCTS] 
                        : "DCT",
                    MepCategoryConstants.PIPES => _categoryPrefixes.ContainsKey(MepCategoryConstants.PIPES) 
                        ? _categoryPrefixes[MepCategoryConstants.PIPES] 
                        : "PLU",
                    MepCategoryConstants.CABLE_TRAYS => _categoryPrefixes.ContainsKey(MepCategoryConstants.CABLE_TRAYS) 
                        ? _categoryPrefixes[MepCategoryConstants.CABLE_TRAYS] 
                        : "ELE",
                    MepCategoryConstants.DUCT_ACCESSORIES => _categoryPrefixes.ContainsKey(MepCategoryConstants.DUCT_ACCESSORIES) 
                        ? _categoryPrefixes[MepCategoryConstants.DUCT_ACCESSORIES] 
                        : "DMP",
                    _ => "D" // Default fallback
                };

                _disciplinePrefixTextBox.Text = disciplinePrefix;
                DebugLogger.Info($"[UpdateDisciplinePrefix] Updated discipline prefix to '{disciplinePrefix}' for MEP Type '{selectedMepType}'");
                
                // ✅ DEBUG: Log the current state of _categoryPrefixes
                DebugLogger.Info($"[UpdateDisciplinePrefix] Current _categoryPrefixes state:");
                foreach (var kvp in _categoryPrefixes)
                {
                    DebugLogger.Info($"[UpdateDisciplinePrefix]   {kvp.Key} = '{kvp.Value}'");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UpdateDisciplinePrefix] Error updating discipline prefix: {ex.Message}");
            }
        }
        
        /// <summary>
        /// ✅ FIX: Save discipline prefix for the current category when user changes it
        /// This ensures discipline prefixes persist per category throughout the session
        /// </summary>
        private void SaveDisciplinePrefixForCurrentCategory()
        {
            try
            {
                if (_disciplinePrefixTextBox == null || _mepTypeCombo == null) return;
                
                var selectedMepType = _mepTypeCombo.SelectedItem?.ToString() ?? string.Empty;
                var disciplinePrefix = _disciplinePrefixTextBox.Text?.Trim() ?? string.Empty;
                
                if (string.IsNullOrEmpty(disciplinePrefix)) return;
                
                // Save the discipline prefix for the current category
                // Use MepCategoryConstants.Normalize() to handle singular/plural variations
                string normalizedCategory = MepCategoryConstants.Normalize(selectedMepType);
                switch (normalizedCategory)
                {
                    case MepCategoryConstants.DUCTS:
                        _categoryPrefixes[MepCategoryConstants.DUCTS] = disciplinePrefix;
                        break;
                    case MepCategoryConstants.PIPES:
                        _categoryPrefixes[MepCategoryConstants.PIPES] = disciplinePrefix;
                        break;
                    case MepCategoryConstants.CABLE_TRAYS:
                        _categoryPrefixes[MepCategoryConstants.CABLE_TRAYS] = disciplinePrefix;
                        break;
                    case MepCategoryConstants.DUCT_ACCESSORIES:
                        _categoryPrefixes[MepCategoryConstants.DUCT_ACCESSORIES] = disciplinePrefix;
                        break;
                }
                
                DebugLogger.Info($"[SaveDisciplinePrefixForCurrentCategory] Saved '{disciplinePrefix}' for category '{selectedMepType}'");
                
                // ✅ DEBUG: Log the updated state of _categoryPrefixes
                DebugLogger.Info($"[SaveDisciplinePrefixForCurrentCategory] Updated _categoryPrefixes state:");
                foreach (var kvp in _categoryPrefixes)
                {
                    DebugLogger.Info($"[SaveDisciplinePrefixForCurrentCategory]   {kvp.Key} = '{kvp.Value}'");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[SaveDisciplinePrefixForCurrentCategory] Error saving discipline prefix: {ex.Message}");
            }
        }
        /// <summary>
        /// Creates clearance panels for different MEP types
        /// 
        /// ⚠️ CRITICAL FEATURE - DO NOT MODIFY WITHOUT APPROVAL ⚠️
        /// This method creates specialized clearance panels for:
        /// - Cable Trays: Shows "Top Side" and "Other Sides" clearance rows
        /// - Duct Accessories: Shows "MEP Connector Side" and "Other Sides" clearance rows
        /// - Standard panels for Ducts and Pipes
        /// 
        /// The visibility is controlled by UpdateClearanceVisibilityForCategory() method
        /// and the MEP type combo box SelectedIndexChanged event handler.
        /// 
        /// Any changes to this method may break the 2-row clearance display functionality.
        /// </summary>
        private void CreateClearancePanels()
        {
            // Standard Clearance Panel (for Ducts) - Same size as other panels
            _clearancePanel = new WinForms.Panel
            {
                Location = new System.Drawing.Point(10, 75),
                Size = new System.Drawing.Size(_rightPanel.Width - 20, 110), // Same height as cable tray and damper panels
                BackColor = System.Drawing.Color.FromArgb(248, 249, 250),
                BorderStyle = WinForms.BorderStyle.FixedSingle,
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right
            };
            _rightPanel.Controls.Add(_clearancePanel);

            // Sub-headers
            var normalHeader = new WinForms.Label
            {
                Text = "Normal (mm)",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold),
                Location = new System.Drawing.Point(170, 10),
                Size = new System.Drawing.Size(100, 18)
            };
            _clearancePanel.Controls.Add(normalHeader);

            var insulatedHeader = new WinForms.Label
            {
                Text = "Insulated (mm)",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold),
                Location = new System.Drawing.Point(300, 10),
                Size = new System.Drawing.Size(110, 18)
            };
            _clearancePanel.Controls.Add(insulatedHeader);

            // Row label
            var clearanceRowLbl = new WinForms.Label
            {
                Text = "Clearance per Side:",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular),
                Location = new System.Drawing.Point(10, 35),
                Size = new System.Drawing.Size(150, 18)
            };
            _clearancePanel.Controls.Add(clearanceRowLbl);

            // Inputs
            var normalText = new WinForms.TextBox
            {
                Location = new System.Drawing.Point(170, 33),
                Size = new System.Drawing.Size(50, 20),
                Text = "50",
                Tag = "normal_clearance",
                Enabled = false,
                BackColor = System.Drawing.Color.LightGray
            };
            _clearancePanel.Controls.Add(normalText);

            var normalLockBtn = new WinForms.Button
            {
                Location = new System.Drawing.Point(225, 33),
                Size = new System.Drawing.Size(25, 20),
                Text = "L",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F),
                Tag = "normal_lock",
                BackColor = System.Drawing.Color.LightGreen
            };
            normalLockBtn.Click += (s, e) => ToggleLock(normalLockBtn, normalText);
            _clearancePanel.Controls.Add(normalLockBtn);

            var insulatedText = new WinForms.TextBox
            {
                Location = new System.Drawing.Point(300, 33),
                Size = new System.Drawing.Size(50, 20),
                Text = "25",
                Tag = "insulated_clearance",
                Enabled = false,
                BackColor = System.Drawing.Color.LightGray
            };
            _clearancePanel.Controls.Add(insulatedText);

            var insulatedLockBtn = new WinForms.Button
            {
                Location = new System.Drawing.Point(355, 33),
                Size = new System.Drawing.Size(25, 20),
                Text = "🔒",
                Font = new System.Drawing.Font("Segoe UI Emoji", 8F),
                Tag = "insulated_lock",
                BackColor = System.Drawing.Color.LightGreen
            };
            insulatedLockBtn.Click += (s, e) => ToggleLock(insulatedLockBtn, insulatedText);
            _clearancePanel.Controls.Add(insulatedLockBtn);

            // Duct Opening Type Selection (for round ducts only) - Row 2 (COMPACT)
            var ductOpeningTypeLabel = new WinForms.Label
            {
                Text = "Opening Type for Round Ducts:",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular),
                Location = new System.Drawing.Point(10, 60), // Row 2 - COMPACT spacing (was 105)
                Size = new System.Drawing.Size(200, 16)
            };
            _clearancePanel.Controls.Add(ductOpeningTypeLabel);

            var ductCircularRadio = new WinForms.RadioButton
            {
                Text = "Circular",
                Location = new System.Drawing.Point(220, 58),
                Size = new System.Drawing.Size(80, 20),
                Checked = true,
                Tag = "duct_opening_circular"
            };
            _clearancePanel.Controls.Add(ductCircularRadio);

            var ductRectangularRadio = new WinForms.RadioButton
            {
                Text = "Rectangular",
                Location = new System.Drawing.Point(310, 58),
                Size = new System.Drawing.Size(100, 20),
                Checked = false,
                Tag = "duct_opening_rectangular"
            };
            _clearancePanel.Controls.Add(ductRectangularRadio);

            ductCircularRadio.CheckedChanged += (s, e) =>
            {
                if (ductCircularRadio.Checked)
                    ductRectangularRadio.Checked = false;
            };
            ductRectangularRadio.CheckedChanged += (s, e) =>
            {
                if (ductRectangularRadio.Checked)
                    ductCircularRadio.Checked = false;
            };

            // Round Duct Clearance Section - Row 3 (IDENTICAL to Row 1 - same X positions)
            var roundDuctLabel = new WinForms.Label
            {
                Text = "Round Duct:",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular),
                Location = new System.Drawing.Point(10, 85),
                Size = new System.Drawing.Size(120, 18)
            };
            _clearancePanel.Controls.Add(roundDuctLabel);

            // Normal clearance input (aligned with Row 1)
            var roundDuctNormalText = new WinForms.TextBox
            {
                Location = new System.Drawing.Point(170, 85), // Same X as Row 1 normal input (was 140)
                Size = new System.Drawing.Size(50, 20),
                Text = "50",
                Tag = "round_duct_normal_clearance",
                Enabled = false,
                BackColor = System.Drawing.Color.LightGray
            };
            _clearancePanel.Controls.Add(roundDuctNormalText);

            var roundDuctNormalLockBtn = new WinForms.Button
            {
                Location = new System.Drawing.Point(225, 85), // Same X as Row 1 normal lock (was 195)
                Size = new System.Drawing.Size(25, 20),
                Text = "🔒",
                Font = new System.Drawing.Font("Segoe UI Emoji", 8F),
                Tag = "round_duct_normal_lock",
                BackColor = System.Drawing.Color.LightGreen
            };
            roundDuctNormalLockBtn.Click += (s, e) => ToggleLock(roundDuctNormalLockBtn, roundDuctNormalText);
            _clearancePanel.Controls.Add(roundDuctNormalLockBtn);

            // Insulated clearance input (aligned with Row 1)
            var roundDuctInsulatedText = new WinForms.TextBox
            {
                Location = new System.Drawing.Point(300, 85), // Same X as Row 1 insulated input (was 270)
                Size = new System.Drawing.Size(50, 20),
                Text = "50",
                Tag = "round_duct_insulated_clearance",
                Enabled = false,
                BackColor = System.Drawing.Color.LightGray
            };
            _clearancePanel.Controls.Add(roundDuctInsulatedText);

            var roundDuctInsulatedLockBtn = new WinForms.Button
            {
                Location = new System.Drawing.Point(355, 85), // Same X as Row 1 insulated lock (was 325)
                Size = new System.Drawing.Size(25, 20),
                Text = "🔒",
                Font = new System.Drawing.Font("Segoe UI Emoji", 8F),
                Tag = "round_duct_insulated_lock",
                BackColor = System.Drawing.Color.LightGreen
            };
            roundDuctInsulatedLockBtn.Click += (s, e) => ToggleLock(roundDuctInsulatedLockBtn, roundDuctInsulatedText);
            _clearancePanel.Controls.Add(roundDuctInsulatedLockBtn);

            // Cable Tray Panel (initially hidden) - Top Side + Other Sides
            _cableTrayPanel = new WinForms.Panel
            {
                Location = new System.Drawing.Point(10, 75),
                Size = new System.Drawing.Size(_rightPanel.Width - 20, 110),
                BackColor = System.Drawing.Color.FromArgb(248, 249, 250),
                BorderStyle = WinForms.BorderStyle.FixedSingle,
                Visible = false,
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right
            };
            _rightPanel.Controls.Add(_cableTrayPanel);
            var ctLabel = new WinForms.Label { Text = "Clearances:", Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold), Location = new System.Drawing.Point(10, 10), Size = new System.Drawing.Size(130, 18) };
            _cableTrayPanel.Controls.Add(ctLabel);
            var topSideLbl = new WinForms.Label { Text = "Top Side:", Location = new System.Drawing.Point(10, 40), Size = new System.Drawing.Size(120, 18) };
            _cableTrayPanel.Controls.Add(topSideLbl);
            var ctNormalHeader = new WinForms.Label { Text = "Normal (mm)", Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold), Location = new System.Drawing.Point(170, 18), Size = new System.Drawing.Size(110, 18) };
            _cableTrayPanel.Controls.Add(ctNormalHeader);
            var ctInsHeader = new WinForms.Label { Text = "Insulated (mm)", Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold), Location = new System.Drawing.Point(300, 18), Size = new System.Drawing.Size(120, 18) };
            _cableTrayPanel.Controls.Add(ctInsHeader);

            var topSideTxt = new WinForms.TextBox { Location = new System.Drawing.Point(170, 38), Size = new System.Drawing.Size(50, 20), Text = "75", Tag = "cabletray_top_clearance", Enabled = false, BackColor = System.Drawing.Color.LightGray };
            _cableTrayPanel.Controls.Add(topSideTxt);

            var topLockBtn = new WinForms.Button
            {
                Location = new System.Drawing.Point(225, 38),
                Size = new System.Drawing.Size(25, 20),
                Text = "L",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F),
                Tag = "cabletray_top_lock",
                BackColor = System.Drawing.Color.LightGreen
            };
            topLockBtn.Click += (s, e) => ToggleLock(topLockBtn, topSideTxt);
            _cableTrayPanel.Controls.Add(topLockBtn);

            var ctInsTopTxt = new WinForms.TextBox { Location = new System.Drawing.Point(300, 38), Size = new System.Drawing.Size(50, 20), Text = "75", Tag = "cabletray_top_insulated", Enabled = false, BackColor = System.Drawing.Color.LightGray };
            _cableTrayPanel.Controls.Add(ctInsTopTxt);

            var topInsLockBtn = new WinForms.Button
            {
                Location = new System.Drawing.Point(355, 38),
                Size = new System.Drawing.Size(25, 20),
                Text = "🔒",
                Font = new System.Drawing.Font("Segoe UI Emoji", 8F),
                Tag = "cabletray_top_ins_lock",
                BackColor = System.Drawing.Color.LightGreen
            };
            topInsLockBtn.Click += (s, e) => ToggleLock(topInsLockBtn, ctInsTopTxt);
            _cableTrayPanel.Controls.Add(topInsLockBtn);

            // Stack 'Other Sides' below
            var otherLbl = new WinForms.Label { Text = "Other Sides:", Location = new System.Drawing.Point(10, 70), Size = new System.Drawing.Size(120, 18) };
            _cableTrayPanel.Controls.Add(otherLbl);
            var otherTxt = new WinForms.TextBox { Location = new System.Drawing.Point(170, 68), Size = new System.Drawing.Size(50, 20), Text = "25", Tag = "cabletray_other_clearance", Enabled = false, BackColor = System.Drawing.Color.LightGray };
            _cableTrayPanel.Controls.Add(otherTxt);

            var otherLockBtn = new WinForms.Button
            {
                Location = new System.Drawing.Point(225, 68),
                Size = new System.Drawing.Size(25, 20),
                Text = "🔒",
                Font = new System.Drawing.Font("Segoe UI Emoji", 8F),
                Tag = "cabletray_other_lock",
                BackColor = System.Drawing.Color.LightGreen
            };
            otherLockBtn.Click += (s, e) => ToggleLock(otherLockBtn, otherTxt);
            _cableTrayPanel.Controls.Add(otherLockBtn);

            var ctInsOtherTxt = new WinForms.TextBox { Location = new System.Drawing.Point(300, 68), Size = new System.Drawing.Size(50, 20), Text = "25", Tag = "cabletray_other_insulated", Enabled = false, BackColor = System.Drawing.Color.LightGray };
            _cableTrayPanel.Controls.Add(ctInsOtherTxt);

            var otherInsLockBtn = new WinForms.Button
            {
                Location = new System.Drawing.Point(355, 68),
                Size = new System.Drawing.Size(25, 20),
                Text = "🔒",
                Font = new System.Drawing.Font("Segoe UI Emoji", 8F),
                Tag = "cabletray_other_ins_lock",
                BackColor = System.Drawing.Color.LightGreen
            };
            otherInsLockBtn.Click += (s, e) => ToggleLock(otherInsLockBtn, ctInsOtherTxt);
            _cableTrayPanel.Controls.Add(otherInsLockBtn);

            // Damper Panel (initially hidden) - MEP Side + Other Sides
            _damperPanel = new WinForms.Panel
            {
                Location = new System.Drawing.Point(10, 75),
                Size = new System.Drawing.Size(_rightPanel.Width - 20, 110),
                BackColor = System.Drawing.Color.FromArgb(248, 249, 250),
                BorderStyle = WinForms.BorderStyle.FixedSingle,
                Visible = false,
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right
            };
            _rightPanel.Controls.Add(_damperPanel);
            var dpLabel = new WinForms.Label { Text = "Clearances:", Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold), Location = new System.Drawing.Point(10, 10), Size = new System.Drawing.Size(160, 18) };
            _damperPanel.Controls.Add(dpLabel);
            var mepSideLbl = new WinForms.Label { Text = "MEP Connector Side:", Location = new System.Drawing.Point(10, 40), Size = new System.Drawing.Size(160, 18) };
            _damperPanel.Controls.Add(mepSideLbl);
            var dNormalHeader = new WinForms.Label { Text = "Normal (mm)", Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold), Location = new System.Drawing.Point(170, 18), Size = new System.Drawing.Size(110, 18) };
            _damperPanel.Controls.Add(dNormalHeader);
            var dInsHeader = new WinForms.Label { Text = "Insulated (mm)", Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold), Location = new System.Drawing.Point(300, 18), Size = new System.Drawing.Size(120, 18) };
            _damperPanel.Controls.Add(dInsHeader);

            var mepSideTxt = new WinForms.TextBox { Location = new System.Drawing.Point(170, 38), Size = new System.Drawing.Size(50, 20), Text = "100", Tag = "ductaccessories_mep_normal", Enabled = false, BackColor = System.Drawing.Color.LightGray };
            _damperPanel.Controls.Add(mepSideTxt);

            var mepLockBtn = new WinForms.Button
            {
                Location = new System.Drawing.Point(225, 38),
                Size = new System.Drawing.Size(25, 20),
                Text = "🔒",
                Font = new System.Drawing.Font("Segoe UI Emoji", 8F),
                Tag = "ductaccessories_mep_lock",
                BackColor = System.Drawing.Color.LightGreen
            };
            mepLockBtn.Click += (s, e) => ToggleLock(mepLockBtn, mepSideTxt);
            _damperPanel.Controls.Add(mepLockBtn);

            var dInsTopTxt = new WinForms.TextBox { Location = new System.Drawing.Point(300, 38), Size = new System.Drawing.Size(50, 20), Text = "100", Tag = "ductaccessories_mep_insulated", Enabled = false, BackColor = System.Drawing.Color.LightGray };
            _damperPanel.Controls.Add(dInsTopTxt);

            var mepInsLockBtn = new WinForms.Button
            {
                Location = new System.Drawing.Point(355, 38),
                Size = new System.Drawing.Size(25, 20),
                Text = "🔒",
                Font = new System.Drawing.Font("Segoe UI Emoji", 8F),
                Tag = "ductaccessories_mep_ins_lock",
                BackColor = System.Drawing.Color.LightGreen
            };
            mepInsLockBtn.Click += (s, e) => ToggleLock(mepInsLockBtn, dInsTopTxt);
            _damperPanel.Controls.Add(mepInsLockBtn);

            // Stack 'Other Sides' below
            var otherDLbl = new WinForms.Label { Text = "Other Sides:", Location = new System.Drawing.Point(10, 70), Size = new System.Drawing.Size(120, 18) };
            _damperPanel.Controls.Add(otherDLbl);
            var otherDTxt = new WinForms.TextBox { Location = new System.Drawing.Point(170, 68), Size = new System.Drawing.Size(50, 20), Text = "50", Tag = "ductaccessories_other_normal", Enabled = false, BackColor = System.Drawing.Color.LightGray };
            _damperPanel.Controls.Add(otherDTxt);

            var otherDLockBtn = new WinForms.Button
            {
                Location = new System.Drawing.Point(225, 68),
                Size = new System.Drawing.Size(25, 20),
                Text = "🔒",
                Font = new System.Drawing.Font("Segoe UI Emoji", 8F),
                Tag = "ductaccessories_other_lock",
                BackColor = System.Drawing.Color.LightGreen
            };
            otherDLockBtn.Click += (s, e) => ToggleLock(otherDLockBtn, otherDTxt);
            _damperPanel.Controls.Add(otherDLockBtn);

            var dInsOtherTxt = new WinForms.TextBox { Location = new System.Drawing.Point(300, 68), Size = new System.Drawing.Size(50, 20), Text = "50", Tag = "ductaccessories_other_insulated", Enabled = false, BackColor = System.Drawing.Color.LightGray };
            _damperPanel.Controls.Add(dInsOtherTxt);

            var otherDInsLockBtn = new WinForms.Button
            {
                Location = new System.Drawing.Point(355, 68),
                Size = new System.Drawing.Size(25, 20),
                Text = "🔒",
                Font = new System.Drawing.Font("Segoe UI Emoji", 8F),
                Tag = "ductaccessories_other_ins_lock",
                BackColor = System.Drawing.Color.LightGreen
            };
            otherDInsLockBtn.Click += (s, e) => ToggleLock(otherDInsLockBtn, dInsOtherTxt);
            _damperPanel.Controls.Add(otherDInsLockBtn);
            var dInsUnit = new WinForms.Label { Text = "mm", Location = new System.Drawing.Point(400, 41), Size = new System.Drawing.Size(30, 16) };
            _damperPanel.Controls.Add(dInsUnit);

            // Pipe Panel (initially hidden) - Standard clearance + Opening Type selection
            _pipePanel = new WinForms.Panel
            {
                Location = new System.Drawing.Point(10, 75),
                Size = new System.Drawing.Size(_rightPanel.Width - 20, 110),
                BackColor = System.Drawing.Color.FromArgb(248, 249, 250),
                BorderStyle = WinForms.BorderStyle.FixedSingle,
                Visible = false,
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right
            };
            _rightPanel.Controls.Add(_pipePanel);

            // Clearance Section
            var pipeClearanceLabel = new WinForms.Label
            {
                Text = "Clearances:",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold),
                Location = new System.Drawing.Point(10, 10),
                Size = new System.Drawing.Size(130, 18)
            };
            _pipePanel.Controls.Add(pipeClearanceLabel);

            var pipeNormalHeader = new WinForms.Label
            {
                Text = "Normal (mm)",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold),
                Location = new System.Drawing.Point(170, 10),
                Size = new System.Drawing.Size(110, 18)
            };
            _pipePanel.Controls.Add(pipeNormalHeader);

            var pipeInsHeader = new WinForms.Label
            {
                Text = "Insulated (mm)",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold),
                Location = new System.Drawing.Point(300, 10),
                Size = new System.Drawing.Size(120, 18)
            };
            _pipePanel.Controls.Add(pipeInsHeader);

            var pipeClearanceRowLbl = new WinForms.Label
            {
                Text = "Clearance per Side:",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular),
                Location = new System.Drawing.Point(10, 35),
                Size = new System.Drawing.Size(150, 18)
            };
            _pipePanel.Controls.Add(pipeClearanceRowLbl);

            var pipeNormalText = new WinForms.TextBox
            {
                Location = new System.Drawing.Point(170, 33),
                Size = new System.Drawing.Size(50, 20),
                Text = "50",
                Tag = "pipes_normal_clearance",
                Enabled = false,
                BackColor = System.Drawing.Color.LightGray
            };
            _pipePanel.Controls.Add(pipeNormalText);

            var pipeNormalLockBtn = new WinForms.Button
            {
                Location = new System.Drawing.Point(225, 33),
                Size = new System.Drawing.Size(25, 20),
                Text = "🔒",
                Font = new System.Drawing.Font("Segoe UI Emoji", 8F),
                Tag = "pipes_normal_lock",
                BackColor = System.Drawing.Color.LightGreen
            };
            pipeNormalLockBtn.Click += (s, e) => ToggleLock(pipeNormalLockBtn, pipeNormalText);
            _pipePanel.Controls.Add(pipeNormalLockBtn);

            var pipeInsText = new WinForms.TextBox
            {
                Location = new System.Drawing.Point(300, 33),
                Size = new System.Drawing.Size(50, 20),
                Text = "50",
                Tag = "pipes_insulated_clearance",
                Enabled = false,
                BackColor = System.Drawing.Color.LightGray
            };
            _pipePanel.Controls.Add(pipeInsText);

            var pipeInsLockBtn = new WinForms.Button
            {
                Location = new System.Drawing.Point(355, 33),
                Size = new System.Drawing.Size(25, 20),
                Text = "🔒",
                Font = new System.Drawing.Font("Segoe UI Emoji", 8F),
                Tag = "pipes_insulated_lock",
                BackColor = System.Drawing.Color.LightGreen
            };
            pipeInsLockBtn.Click += (s, e) => ToggleLock(pipeInsLockBtn, pipeInsText);
            _pipePanel.Controls.Add(pipeInsLockBtn);

            // ✅ Opening Type Section - FIXED: Move to Row 2 (like Round Duct layout)
            var pipeOpeningTypeLabel = new WinForms.Label
            {
                Text = "Opening Type for Pipes:",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular),
                Location = new System.Drawing.Point(10, 60), // Row 2 (same as Round Duct)
                Size = new System.Drawing.Size(200, 16)
            };
            _pipePanel.Controls.Add(pipeOpeningTypeLabel);

            var pipeCircularRadio = new WinForms.RadioButton
            {
                Text = "Circular",
                Location = new System.Drawing.Point(220, 58), // Row 2 (same as Round Duct)
                Size = new System.Drawing.Size(80, 20),
                Checked = true, // Default to circular
                Tag = "pipe_opening_circular"
            };
            _pipePanel.Controls.Add(pipeCircularRadio);

            var pipeRectangularRadio = new WinForms.RadioButton
            {
                Text = "Rectangular",
                Location = new System.Drawing.Point(310, 58), // Row 2 (same as Round Duct)
                Size = new System.Drawing.Size(100, 20),
                Checked = false,
                Tag = "pipe_opening_rectangular"
            };
            _pipePanel.Controls.Add(pipeRectangularRadio);

            // Group the radio buttons so only one can be selected
            pipeCircularRadio.CheckedChanged += (s, e) =>
            {
                if (pipeCircularRadio.Checked)
                    pipeRectangularRadio.Checked = false;
            };
            pipeRectangularRadio.CheckedChanged += (s, e) =>
            {
                if (pipeRectangularRadio.Checked)
                    pipeCircularRadio.Checked = false;
            };
        }

        private void ToggleLock(WinForms.Button lockBtn, WinForms.TextBox textBox)
        {
            if (lockBtn.Text == "🔒")
            {
                // Unlock - allow editing
                lockBtn.Text = "🔓";
                lockBtn.BackColor = System.Drawing.Color.LightCoral;
                textBox.Enabled = true;
                textBox.BackColor = System.Drawing.Color.White;
            }
            else
            {
                // Lock - disable editing AND SAVE THE VALUE
                lockBtn.Text = "🔒";
                lockBtn.BackColor = System.Drawing.Color.LightGreen;
                textBox.Enabled = false;
                textBox.BackColor = System.Drawing.Color.LightGray;
                
                // ✅ SAVE: When locking, save the current clearance value for this category
                SaveCurrentClearanceValues();
            }
        }
        private void SetDefaultClearanceValues(string category)
        {
            try
            {
                switch (category.ToLower())
                {
                    case "ducts":
                        SetClearancePanelValues("50", "25");
                        break;
                    case "pipes":
                        SetPipePanelValues("50", "25");
                        break;
                    case "cable trays":
                        SetCableTrayPanelValues("75", "25", "75", "25");
                        break;
                    case "duct accessories":
                        SetDamperPanelValues("100", "50", "100", "50");
                        break;
                    default:
                        SetClearancePanelValues("50", "25");
                        break;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error setting default clearance values: {ex.Message}");
            }
        }

        private void SetClearancePanelValues(string normalValue, string insulatedValue)
        {
            if (_clearancePanel?.Controls.Count > 0)
            {
                foreach (var control in _clearancePanel.Controls)
                {
                    if (control is WinForms.TextBox textBox)
                    {
                        if (textBox.Tag?.ToString() == "normal_clearance")
                        {
                            textBox.Text = normalValue;
                        }
                        else if (textBox.Tag?.ToString() == "insulated_clearance")
                        {
                            textBox.Text = insulatedValue;
                        }
                    }
                }
            }
        }

        private void SetCableTrayPanelValues(string topNormal, string otherNormal, string topInsulated, string otherInsulated)
        {
            if (_cableTrayPanel?.Controls.Count > 0)
            {
                foreach (var control in _cableTrayPanel.Controls)
                {
                    if (control is WinForms.TextBox textBox)
                    {
                        switch (textBox.Tag?.ToString())
                        {
                            case "cabletray_top_normal":
                                textBox.Text = topNormal;
                                break;
                            case "cabletray_other_normal":
                                textBox.Text = otherNormal;
                                break;
                            case "cabletray_top_insulated":
                                textBox.Text = topInsulated;
                                break;
                            case "cabletray_other_insulated":
                                textBox.Text = otherInsulated;
                                break;
                        }
                    }
                }
            }
        }

        private void SetDamperPanelValues(string mepNormal, string otherNormal, string mepInsulated, string otherInsulated)
        {
            if (_damperPanel?.Controls.Count > 0)
            {
                foreach (var control in _damperPanel.Controls)
                {
                    if (control is WinForms.TextBox textBox)
                    {
                        switch (textBox.Tag?.ToString())
                        {
                            case "ductaccessories_mep_normal":
                                textBox.Text = mepNormal;
                                break;
                            case "ductaccessories_other_normal":
                                textBox.Text = otherNormal;
                                break;
                            case "ductaccessories_mep_insulated":
                                textBox.Text = mepInsulated;
                                break;
                            case "ductaccessories_other_insulated":
                                textBox.Text = otherInsulated;
                                break;
                        }
                    }
                }
            }
        }
        private void SetPipePanelValues(string normalValue, string insulatedValue)
        {
            if (_pipePanel?.Controls.Count > 0)
            {
                foreach (var control in _pipePanel.Controls)
                {
                    if (control is WinForms.TextBox textBox)
                    {
                        if (textBox.Tag?.ToString() == "pipes_normal_clearance")
                        {
                            textBox.Text = normalValue;
                        }
                        else if (textBox.Tag?.ToString() == "pipes_insulated_clearance")
                        {
                            textBox.Text = insulatedValue;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Save current clearance values from the visible panel to the category storage dictionary
        /// </summary>
        private void SaveCurrentClearanceValues()
        {
            if (string.IsNullOrEmpty(_currentCategory))
                return;

            try
            {
                // Initialize category dictionary if it doesn't exist
                if (!_categoryClearanceValues.ContainsKey(_currentCategory))
                {
                    _categoryClearanceValues[_currentCategory] = new Dictionary<string, string>();
                }

                var categoryValues = _categoryClearanceValues[_currentCategory];

                // Save values from the currently visible panel
                WinForms.Panel activePanel = null;
                if (_clearancePanel.Visible)
                    activePanel = _clearancePanel;
                else if (_cableTrayPanel.Visible)
                    activePanel = _cableTrayPanel;
                else if (_damperPanel.Visible)
                    activePanel = _damperPanel;
                else if (_pipePanel.Visible)
                    activePanel = _pipePanel;

                if (activePanel != null)
                {
                    foreach (var control in activePanel.Controls)
                    {
                        if (control is WinForms.TextBox textBox && textBox.Tag != null)
                        {
                            string tag = textBox.Tag.ToString() ?? "";
                            categoryValues[tag] = textBox.Text;
                            DebugLogger.Info($"[SaveCurrentClearanceValues] Saved {_currentCategory}.{tag} = {textBox.Text}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[SaveCurrentClearanceValues] Error: {ex.Message}");
            }
        }
        /// <summary>
        /// Restore clearance values for a category from storage, or set defaults if not saved
        /// ✅ FIX: Also tries to load from database if in-memory values are not available
        /// </summary>
        private void RestoreClearanceValues(string category)
        {
            try
            {
                // First, try to load from database if not already in memory
                if (!_categoryClearanceValues.ContainsKey(category) || _categoryClearanceValues[category].Count == 0)
                {
                    LoadClearanceValuesFromDatabase(category);
                }
                
                // Check if we have saved values for this category
                if (_categoryClearanceValues.ContainsKey(category) && _categoryClearanceValues[category].Count > 0)
                {
                    // Restore saved values
                    var savedValues = _categoryClearanceValues[category];
                    WinForms.Panel activePanel = null;
                    
                    if (category.Equals("Cable Trays", StringComparison.OrdinalIgnoreCase))
                        activePanel = _cableTrayPanel;
                    else if (category.Equals("Duct Accessories", StringComparison.OrdinalIgnoreCase))
                        activePanel = _damperPanel;
                    else if (category.Equals("Ducts", StringComparison.OrdinalIgnoreCase))
                        activePanel = _clearancePanel;
                    else if (category.Equals("Pipes", StringComparison.OrdinalIgnoreCase))
                        activePanel = _pipePanel;
                    else
                        activePanel = _clearancePanel; // default

                    if (activePanel != null)
                    {
                        foreach (var control in activePanel.Controls)
                        {
                            if (control is WinForms.TextBox textBox && textBox.Tag != null)
                            {
                                string tag = textBox.Tag.ToString() ?? "";
                                if (savedValues.ContainsKey(tag))
                                {
                                    textBox.Text = savedValues[tag];
                                    DebugLogger.Info($"[RestoreClearanceValues] Restored {category}.{tag} = {savedValues[tag]}");
                                }
                            }
                        }
                    }
                }
                else
                {
                    // No saved values - use defaults
                    SetDefaultClearanceValues(category);
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[RestoreClearanceValues] Error: {ex.Message}");
                // Fallback to defaults on error
                SetDefaultClearanceValues(category);
            }
        }
        
        /// <summary>
        /// ✅ NEW: Load clearance values from database and populate _categoryClearanceValues
        /// </summary>
        private void LoadClearanceValuesFromDatabase(string category)
        {
            try
            {
                if (_document == null) return;
                
                // Get selected filters
                var selectedFilters = GetSelectedFilters();
                if (selectedFilters.Count == 0) return;
                
                var filter = selectedFilters.FirstOrDefault();
                if (filter == null) return;
                
                // Build combined key using same logic as SaveConditionsToXml
                // ✅ FIX: Strip .xml extension from filter name if present (prevents duplicate database entries)
                string cleanFilterName = filter.Name;
                if (cleanFilterName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                {
                    cleanFilterName = cleanFilterName.Substring(0, cleanFilterName.Length - 4);
                }
                string normalizedCategory = MepCategoryConstants.GetXmlSuffix(category);
                string combinedKey = $"{cleanFilterName}_{normalizedCategory}";
                
                // Load conditions from database
                using (var dbContext = new SleeveDbContext(_document))
                {
                    var conditionRepository = new ConditionRepository(dbContext);
                    var conditions = conditionRepository.GetConditions(combinedKey);
                    
                    if (conditions?.ClearanceSettings != null)
                    {
                        var clearance = conditions.ClearanceSettings;
                        
                        // Initialize category dictionary if it doesn't exist
                        if (!_categoryClearanceValues.ContainsKey(category))
                        {
                            _categoryClearanceValues[category] = new Dictionary<string, string>();
                        }
                        
                        var categoryValues = _categoryClearanceValues[category];
                        
                        // Map database clearance values to UI tags based on category
                        // Normalize category name to handle both "Pipe" and "Pipes"
                        string normalizedCategoryForMapping = category.Equals("Pipe", StringComparison.OrdinalIgnoreCase) ? "Pipes" : category;
                        
                        if (normalizedCategoryForMapping.Equals("Pipes", StringComparison.OrdinalIgnoreCase))
                        {
                            categoryValues["pipes_normal_clearance"] = clearance.PipesNormal.ToString("F0");
                            categoryValues["pipes_insulated_clearance"] = clearance.PipesInsulated.ToString("F0");
                            DebugLogger.Info($"[LoadClearanceValuesFromDatabase] Loaded Pipes: Normal={clearance.PipesNormal}mm, Insulated={clearance.PipesInsulated}mm");
                        }
                        else if (category.Equals("Cable Trays", StringComparison.OrdinalIgnoreCase))
                        {
                            categoryValues["cabletray_top_normal"] = clearance.CableTrayTop.ToString("F0");
                            categoryValues["cabletray_other_normal"] = clearance.CableTrayOther.ToString("F0");
                            // Note: Cable tray doesn't have separate insulated values in current model
                            DebugLogger.Info($"[LoadClearanceValuesFromDatabase] Loaded Cable Trays: Top={clearance.CableTrayTop}mm, Other={clearance.CableTrayOther}mm");
                        }
                        else if (category.Equals("Duct Accessories", StringComparison.OrdinalIgnoreCase))
                        {
                            categoryValues["ductaccessories_mep_normal"] = clearance.DuctAccessoryMepNormal.ToString("F0");
                            categoryValues["ductaccessories_other_normal"] = clearance.DuctAccessoryOtherNormal.ToString("F0");
                            DebugLogger.Info($"[LoadClearanceValuesFromDatabase] Loaded Duct Accessories: MEP={clearance.DuctAccessoryMepNormal}mm, Other={clearance.DuctAccessoryOtherNormal}mm");
                        }
                        else if (category.Equals("Ducts", StringComparison.OrdinalIgnoreCase))
                        {
                            categoryValues["normal_clearance"] = clearance.RectangularNormal.ToString("F0");
                            categoryValues["insulated_clearance"] = clearance.RectangularInsulated.ToString("F0");
                            categoryValues["round_duct_normal_clearance"] = clearance.RoundNormal.ToString("F0");
                            categoryValues["round_duct_insulated_clearance"] = clearance.RoundInsulated.ToString("F0");
                            DebugLogger.Info($"[LoadClearanceValuesFromDatabase] Loaded Ducts: Rect Normal={clearance.RectangularNormal}mm, Rect Ins={clearance.RectangularInsulated}mm, Round Normal={clearance.RoundNormal}mm, Round Ins={clearance.RoundInsulated}mm");
                        }
                        
                        DebugLogger.Info($"[LoadClearanceValuesFromDatabase] ✅ Loaded clearance values from database for category '{category}' (combinedKey='{combinedKey}')");
                    }
                    else
                    {
                        DebugLogger.Info($"[LoadClearanceValuesFromDatabase] No conditions found in database for category '{category}' (combinedKey='{combinedKey}') - will use defaults");
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[LoadClearanceValuesFromDatabase] Error loading clearance values from database for category '{category}': {ex.Message}");
            }
        }

        /// <summary>
        /// Get the selected opening type for pipes (Circular or Rectangular)
        /// </summary>
        private string GetPipeOpeningType()
        {
            if (_pipePanel?.Controls.Count > 0)
            {
                foreach (var control in _pipePanel.Controls)
                {
                    if (control is WinForms.RadioButton radioButton)
                    {
                        if (radioButton.Checked)
                        {
                            if (radioButton.Tag?.ToString() == "pipe_opening_circular")
                                return "Circular";
                            else if (radioButton.Tag?.ToString() == "pipe_opening_rectangular")
                                return "Rectangular";
                        }
                    }
                }
            }
            return "Circular"; // Default to circular
        }

        /// <summary>
        /// Get the selected opening type for ducts (Circular or Rectangular)
        /// </summary>
        public string GetDuctOpeningType()
        {
            if (_clearancePanel?.Controls.Count > 0)
            {
                foreach (var control in _clearancePanel.Controls)
                {
                    if (control is WinForms.RadioButton radioButton)
                    {
                        if (radioButton.Checked)
                        {
                            if (radioButton.Tag?.ToString() == "duct_opening_circular")
                                return "Circular";
                            else if (radioButton.Tag?.ToString() == "duct_opening_rectangular")
                                return "Rectangular";
                        }
                    }
                }
            }
            return "Circular"; // Default to circular
        }

        // Parameter service methods removed - now available as separate dialog

        // Parameter marking section removed - now available in separate parameter service dialog


        // Reference elements master tab removed - now in separate parameter service dialog
        // Host elements master tab removed - now in separate parameter service dialog
        // Service tab creation removed - now in separate parameter service dialog

        // AddServiceParameterRow method removed - now in separate parameter service dialog

        /// <summary>
        /// Gets the current MEP parameters that were collected during the latest refresh
        /// </summary>
        private List<string> GetCurrentMepParameters()
        {
            try
            {
                // Get parameters from the latest refresh
                var mepParameters = new List<string>();
                
                // Get selected MEP categories
                var selectedCategories = GetSelectedMepCategories();
                if (selectedCategories.Count == 0)
                {
                    DebugLogger.Warning("[PARAMETER_CURRENT] No MEP categories selected, returning empty list");
                    return mepParameters;
                }
                
                // Get current document
                var document = _uiDocument?.Document;
                if (document == null)
                {
                    DebugLogger.Warning("[PARAMETER_CURRENT] No active document, returning empty list");
                    return mepParameters;
                }
                
                // Convert string categories to enum
                var mepCategories = new List<Services.MepCategory>();
                foreach (var categoryName in selectedCategories)
                {
                    if (Enum.TryParse<Services.MepCategory>(categoryName, out var category))
                    {
                        mepCategories.Add(category);
                    }
                }
                
                if (mepCategories.Count == 0)
                {
                    DebugLogger.Warning("[PARAMETER_CURRENT] No valid MEP categories found, returning empty list");
                    return mepParameters;
                }
                
                // Use ParameterExtractionService to get parameters from linked files
                var parameterService = new Services.ParameterExtractionService();
                var linkedFileService = new Services.LinkedFileService();
                var linkedFiles = linkedFileService.GetLinkedFiles(document);
                
                if (linkedFiles.Count > 0)
                {
                    // Get parameters from linked files
                    foreach (var linkedFile in linkedFiles)
                    {
                        try
                        {
                            var linkedDoc = linkedFile.LinkInstance?.GetLinkDocument();
                            if (linkedDoc != null)
                            {
                                var parameters = parameterService.GetParametersForMepCategories(linkedDoc, mepCategories);
                                var parameterNames = parameters.Select(p => p.Name).ToList();
                                
                                // Add unique parameters
                                foreach (var paramName in parameterNames)
                                {
                                    if (!mepParameters.Contains(paramName))
                                    {
                                        mepParameters.Add(paramName);
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            DebugLogger.Warning($"[PARAMETER_CURRENT] Error getting parameters from linked file '{linkedFile.FileName}': {ex.Message}");
                        }
                    }
                }
                else
                {
                    // Fallback to current document if no linked files
                    var parameters = parameterService.GetParametersForMepCategories(document, mepCategories);
                    mepParameters = parameters.Select(p => p.Name).ToList();
                }
                DebugLogger.Info($"[PARAMETER_CURRENT] Retrieved {mepParameters.Count} MEP parameters for new row");
                
                return mepParameters;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[PARAMETER_CURRENT] Error getting current MEP parameters: {ex.Message}");
                return new List<string>();
            }
        }

        /// <summary>
        /// Transfer all mappings in the given service panel: for each row, copy the selected
        /// item from the combo tagged "mep" into the combo tagged "opening" (if possible).
        /// This implements the tab-level "Transfer All" behavior requested by UX.
        /// NOTE: This method is no longer used as parameter service has been moved to separate dialog.
        /// </summary>
        private void TransferAllMappingsFromPanel(WinForms.Panel servicePanel)
        {
            // Parameter service moved to separate dialog - this method is no longer used
            DebugLogger.Info("[TRANSFER_ALL] Parameter service moved to separate dialog - method no longer used");
            return;
        }

        private List<ElementId> GetAllOpeningInstanceIds(Autodesk.Revit.DB.Document doc)
        {
            var ids = new List<ElementId>();
            try
            {
                // Target opening families used throughout the app
                var targetFamilyNames = new List<string>
                {
                    "RectangularOpeningOnWall",
                    "RectangularOpeningOnSlab",
                    "CircularOpeningOnWall",
                    "CircularOpeningOnSlab"
                };

                var instances = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>();

                foreach (var fi in instances)
                {
                    var famName = fi.Symbol?.Family?.Name ?? string.Empty;
                    if (targetFamilyNames.Any(n => famName.IndexOf(n, StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        ids.Add(fi.Id);
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[TRANSFER_ALL] Error collecting opening instance ids: {ex.Message}");
            }
            return ids;
        }
        
        /// <summary>
        /// Gets the current opening parameters that were collected during the latest refresh
        /// </summary>
        private List<string> GetCurrentOpeningParameters()
        {
            try
            {
                // Get parameters from the latest refresh
                var openingParameters = new List<string>();
                
                // Get current document
                var document = _uiDocument?.Document;
                if (document == null)
                {
                    DebugLogger.Warning("[PARAMETER_CURRENT] No active document, returning empty list");
                    return openingParameters;
                }
                
                // Get opening families
                var openingFamilies = GetOpeningFamilies(document);
                if (openingFamilies.Count == 0)
                {
                    DebugLogger.Warning("[PARAMETER_CURRENT] No opening families found, returning empty list");
                    return openingParameters;
                }
                
                // Get parameters from opening families directly
                foreach (var family in openingFamilies.Take(5)) // Sample first 5 families
                {
                    var parameters = family.Parameters;
                    foreach (Parameter param in parameters)
                    {
                        if (!openingParameters.Contains(param.Definition.Name))
                        {
                            openingParameters.Add(param.Definition.Name);
                        }
                    }
                }
                
                DebugLogger.Info($"[PARAMETER_CURRENT] Retrieved {openingParameters.Count} opening parameters for new row");
                
                return openingParameters;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[PARAMETER_CURRENT] Error getting current opening parameters: {ex.Message}");
                return new List<string>();
            }
        }

        private void RepositionServiceParameterRows(WinForms.Panel servicePanel)
        {
            var rows = servicePanel.Controls.OfType<WinForms.Panel>().ToList();
            for (int i = 0; i < rows.Count; i++)
            {
                rows[i].Location = new System.Drawing.Point(8, 25 + (i * 27));
            }
        }

        private void AddParameterRow(string parameterName, string value)
        {
            int rowHeight = 28;
            // int top = 30 + (_parameterRows.Count * (rowHeight + 6)); // Parameter service moved to separate dialog

            var row = new WinForms.Panel
            {
                // Location = new System.Drawing.Point(8, top), // Parameter service moved to separate dialog
                // Size = new System.Drawing.Size(_parameterFilterPanel.Width - 16, rowHeight), // Parameter service moved to separate dialog
                // Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right // Parameter service moved to separate dialog
            };
            // _parameterFilterPanel.Controls.Add(row); // Parameter service moved to separate dialog
            // _parameterRows.Add(row); // Parameter service moved to separate dialog

            var nameCombo = new WinForms.ComboBox
            {
                Location = new System.Drawing.Point(0, 3),
                Size = new System.Drawing.Size(150, 21),
                DropDownStyle = WinForms.ComboBoxStyle.DropDownList
            };
            
            // Get real parameters from the model instead of hardcoded values
            var availableParameters = GetAvailableParameters();
            nameCombo.Items.AddRange(availableParameters.ToArray());
            nameCombo.SelectedItem = parameterName;
            
            row.Controls.Add(nameCombo);

            var valueCombo = new WinForms.ComboBox
            {
                Location = new System.Drawing.Point(160, 3),
                Size = new System.Drawing.Size(row.Width - 160 - 60, 21),
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right,
                DropDownStyle = WinForms.ComboBoxStyle.DropDownList
            };
            
            // Get real parameter values instead of hardcoded values
            var availableValues = GetAvailableParameterValues(parameterName);
            valueCombo.Items.AddRange(availableValues.ToArray());
            valueCombo.SelectedIndex = 0;
            
            // Add event handler to update parameter values when parameter selection changes
            nameCombo.SelectedIndexChanged += (sender, e) => {
                UpdateParameterValues(valueCombo, nameCombo.SelectedItem?.ToString());
            };
            
            row.Controls.Add(valueCombo);

            var removeBtn = new WinForms.Button
            {
                Text = "-",
                Location = new System.Drawing.Point(row.Width - 50, 2),
                Size = new System.Drawing.Size(24, 24),
                BackColor = System.Drawing.Color.FromArgb(255, 230, 230),
                FlatStyle = WinForms.FlatStyle.Flat,
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Right
            };
            removeBtn.Click += (_, __) =>
            {
                // _parameterFilterPanel.Controls.Remove(row); // Parameter service moved to separate dialog
                // _parameterRows.Remove(row); // Parameter service moved to separate dialog
                ReflowParameterRows();
            };
            row.Controls.Add(removeBtn);
        }

        private void ReflowParameterRows()
        {
            int rowHeight = 28;
            // for (int i = 0; i < _parameterRows.Count; i++) // Parameter service moved to separate dialog
            // {
            //     var row = _parameterRows[i]; // Parameter service moved to separate dialog
            //     row.Location = new System.Drawing.Point(8, 30 + (i * (rowHeight + 6))); // Parameter service moved to separate dialog
            //     row.Size = new System.Drawing.Size(_parameterFilterPanel.Width - 16, rowHeight); // Parameter service moved to separate dialog
            // }
        }

        /// <summary>
        /// Updates parameter values dropdown based on selected parameter
        /// </summary>
        private void UpdateParameterValues(WinForms.ComboBox valueCombo, string? selectedParameter)
        {
            try
            {
                if (valueCombo == null || string.IsNullOrEmpty(selectedParameter) || selectedParameter == "<Select>")
                {
                    if (valueCombo != null)
                    {
                        valueCombo.Items.Clear();
                        valueCombo.Items.Add("<Auto Selection>");
                        valueCombo.SelectedIndex = 0;
                    }
                    return;
                }

                // Get parameter values for the selected parameter
                var availableValues = GetAvailableParameterValues(selectedParameter);
                
                // Update the value combo box
                valueCombo.Items.Clear();
                valueCombo.Items.AddRange(availableValues.ToArray());
                valueCombo.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                // Log error but don't crash
                try
                {
                    string logPath = SafeFileLogger.GetLogFilePath("parameter_update_error.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now}] Error updating parameter values for '{selectedParameter}': {ex.Message}\n");
                }
                catch { }
                
                // Fallback to basic values
                valueCombo.Items.Clear();
                valueCombo.Items.Add("<Auto Selection>");
                valueCombo.SelectedIndex = 0;
            }
        }

        /// <summary>
        /// Gets available parameters from the current model for MEP categories
        /// </summary>
        private List<string> GetAvailableParameters()
        {
            return GetAvailableParametersForCategories(null); // Get all categories
        }

        private List<string> GetOpeningSleeveParameters()
        {
            return new List<string>
            {
                "Mark",
                "Type Mark",
                "Assembly Code",
                "Assembly Description",
                "Type Comments",
                "URL",
                "Description",
                "Type Image",
                "Keynote",
                "Manufacturer",
                "Model",
                "Comments",
                "Height",
                "Width",
                "Diameter",
                "Outside Diameter",
                "Inside Diameter",
                "Clearance",
                "Opening Type",
                "Service Type",
                "Level",
                "Host Element",
                "Reference Element",
                "Installation Date",
                "Installation Notes"
            };
        }

        /// <summary>
        /// Gets available parameters for specific MEP categories
        /// </summary>
        private List<string> GetAvailableParametersForCategories(List<Models.MepCategory>? specificCategories)
        {
            var parameters = new List<string> { "<Select>" };
            
            try
            {
                if (_activeDocument != null)
                {
                    var parameterService = new ParameterExtractionService();
                    
                    // Use specific categories if provided, otherwise use all common MEP categories
                    var mepCategories = specificCategories ?? new List<Models.MepCategory>
                    {
                        Models.MepCategory.Pipes,
                        Models.MepCategory.Ducts,
                        Models.MepCategory.CableTrays,
                        Models.MepCategory.CableTrays // Conduits not available, use CableTrays instead
                    };
                    
                    var parameterInfos = parameterService.GetParametersForMepCategories(_activeDocument, mepCategories.Cast<Services.MepCategory>().ToList());
                    var parameterNames = parameterService.GetParameterNamesForDisplay(parameterInfos.Cast<Services.ParameterInfo>().ToList());
                    
                    // Add common useful parameters first
                    var commonParameters = new List<string>
                    {
                        "Reference Level",
                        "Size",
                        "Service Type",
                        "Category",
                        "System Type",
                        "System Name",
                        "Comments",
                        "Mark",
                        "Type Mark",
                        "Family",
                        "Type"
                    };
                    
                    // Add common parameters that exist in the model
                    foreach (var commonParam in commonParameters)
                    {
                        if (parameterNames.Contains(commonParam) && !parameters.Contains(commonParam))
                        {
                            parameters.Add(commonParam);
                        }
                    }
                    
                    // Add other parameters from the model
                    foreach (var paramName in parameterNames)
                    {
                        if (!parameters.Contains(paramName) && !string.IsNullOrEmpty(paramName))
                        {
                            parameters.Add(paramName);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log error but don't crash - fall back to basic parameters
                try
                {
                    string logPath = SafeFileLogger.GetLogFilePath("parameter_extraction_error.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now}] Error getting parameters: {ex.Message}\n");
                }
                catch { }
                
                // Fallback to basic parameters
                parameters.AddRange(new[] { "Reference Level", "Size", "Service Type", "Category" });
            }
            
            return parameters;
        }

        /// <summary>
        /// Gets available parameters for specific categories (e.g., only Pipes, only Ducts, etc.)
        /// </summary>
        private List<string> GetAvailableParametersForSpecificCategories(List<string> categoryNames)
        {
            var parameters = new List<string> { "<Select>" };
            
            try
            {
                if (_activeDocument != null && categoryNames.Any())
                {
                    // ⚠️ CRITICAL FIX: Skip parameter extraction during initialization to prevent sleeve deletion
                    if (_isInitializing)
                    {
                        // During initialization, use cached parameters to avoid triggering refresh operations
                        parameters.AddRange(new[] { "Width", "Height", "Length", "Level", "Reference Level", "Schedule Level" });
                        DebugLogger.Info("🔍 SKIPPING GetAvailableParametersForSpecificCategories during initialization - using cached parameters");
                        return parameters;
                    }
                    
                    var parameterService = new ParameterExtractionService();
                    
                    // Convert category names to MepCategory enum
                    var mepCategories = new List<Models.MepCategory>();
                    foreach (var categoryName in categoryNames)
                    {
                        var category = GetMepCategoryFromName(categoryName);
                        if (category.HasValue)
                        {
                            mepCategories.Add(category.Value);
                        }
                    }
                    
                    if (mepCategories.Any())
                    {
                        var parameterInfos = parameterService.GetParametersForMepCategories(_activeDocument, mepCategories.Cast<Services.MepCategory>().ToList());
                        var parameterNames = parameterService.GetParameterNamesForDisplay(parameterInfos);
                        
                        // Add category-specific parameters
                        foreach (var paramName in parameterNames)
                        {
                            if (!string.IsNullOrEmpty(paramName) && !parameters.Contains(paramName))
                            {
                                parameters.Add(paramName);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log error but don't crash
                try
                {
                    string logPath = SafeFileLogger.GetLogFilePath("category_parameter_error.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now}] Error getting parameters for categories {string.Join(", ", categoryNames)}: {ex.Message}\n");
                }
                catch { }
            }
            
            return parameters;
        }

        /// <summary>
        /// Converts category name string to MepCategory enum
        /// </summary>
        private Models.MepCategory? GetMepCategoryFromName(string categoryName)
        {
            return categoryName.ToLower() switch
            {
                "pipes" or "pipe" => Models.MepCategory.Pipes,
                "ducts" or "duct" => Models.MepCategory.Ducts,
                "cable trays" or "cable tray" => Models.MepCategory.CableTrays,
                "conduits" or "conduit" => Models.MepCategory.CableTrays, // Conduits not available, use CableTrays instead
                "duct accessories" or "duct accessory" => Models.MepCategory.DuctAccessories,
                "duct fittings" or "duct fitting" => Models.MepCategory.DuctAccessories, // DuctFittings not available, use DuctAccessories instead
                _ => null
            };
        }
        /// <summary>
        /// Gets available values for a specific parameter from the current model
        /// </summary>
        private List<string> GetAvailableParameterValues(string parameterName)
        {
            var values = new List<string> { "<Auto Selection>" };
            
            try
            {
                if (_activeDocument != null && !string.IsNullOrEmpty(parameterName) && parameterName != "<Select>")
                {
                    var parameterService = new ParameterExtractionService();
                    
                    // Get parameters for common MEP categories
                    var mepCategories = new List<Models.MepCategory>
                    {
                        Models.MepCategory.Pipes,
                        Models.MepCategory.Ducts,
                        Models.MepCategory.CableTrays,
                        Models.MepCategory.CableTrays // Conduits not available, use CableTrays instead
                    };
                    
                    var parameterInfos = parameterService.GetParametersForMepCategories(_activeDocument, mepCategories.Cast<Services.MepCategory>().ToList());
                    var targetParameter = parameterInfos.Cast<Services.ParameterInfo>().FirstOrDefault(p => p.Name == parameterName);
                    
                    if (targetParameter != null && targetParameter.Values != null && targetParameter.Values.Any())
                    {
                        // Add unique values from the model
                        foreach (var value in targetParameter.Values)
                        {
                            if (!string.IsNullOrEmpty(value) && !values.Contains(value))
                            {
                                values.Add(value);
                            }
                        }
                    }
                    
                    // Add some common values based on parameter type
                    AddCommonValuesForParameter(parameterName, values);
                }
            }
            catch (Exception ex)
            {
                // Log error but don't crash - fall back to basic values
                try
                {
                    string logPath = SafeFileLogger.GetLogFilePath("parameter_values_error.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now}] Error getting parameter values for '{parameterName}': {ex.Message}\n");
                }
                catch { }
                
                // Fallback to basic values
                AddCommonValuesForParameter(parameterName, values);
            }
            
            return values;
        }

        /// <summary>
        /// Adds common values for specific parameter types
        /// </summary>
        private void AddCommonValuesForParameter(string parameterName, List<string> values)
        {
            switch (parameterName.ToLower())
            {
                case "size":
                    var sizes = new[] { "100x100", "200x200", "300x300", "400x400", "100", "200", "300", "400" };
                    foreach (var size in sizes)
                    {
                        if (!values.Contains(size)) values.Add(size);
                    }
                    break;
                    
                case "reference level":
                case "level":
                    var levels = new[] { "A_GARDEN LEVEL", "GROUND FLOOR", "FIRST FLOOR", "SECOND FLOOR" };
                    foreach (var level in levels)
                    {
                        if (!values.Contains(level)) values.Add(level);
                    }
                    break;
                    
                case "service type":
                case "system type":
                    var services = new[] { "ESSENTIAL POWER", "LIGHTING", "HVAC", "PLUMBING", "FIRE PROTECTION" };
                    foreach (var service in services)
                    {
                        if (!values.Contains(service)) values.Add(service);
                    }
                    break;
                    
                case "category":
                    var categories = new[] { "Pipes", "Ducts", "Cable Trays", "Conduits" };
                    foreach (var category in categories)
                    {
                        if (!values.Contains(category)) values.Add(category);
                    }
                    break;
            }
        }
        /// <summary>
        /// Refreshes parameter dropdowns based on selected reference file and categories
        /// </summary>
        public void RefreshParametersForSelection(List<string> selectedCategories)
        {
            try
            {
                // Get category-specific parameters
                var categoryParameters = GetAvailableParametersForSpecificCategories(selectedCategories);
                
                // Update all parameter name dropdowns in existing rows
                // Parameter service moved to separate dialog - this functionality is now handled in ParameterServiceDialog
                // foreach (var row in _parameterRows)
                // {
                //     var nameCombo = row.Controls.OfType<WinForms.ComboBox>().FirstOrDefault();
                //     if (nameCombo != null)
                //     {
                //         var currentSelection = nameCombo.SelectedItem?.ToString();
                //         
                //         // Update the dropdown items
                //         nameCombo.Items.Clear();
                //         nameCombo.Items.AddRange(categoryParameters.ToArray());
                //         
                //         // Restore selection if it still exists in the new list
                //         if (!string.IsNullOrEmpty(currentSelection) && categoryParameters.Contains(currentSelection))
                //         {
                //             nameCombo.SelectedItem = currentSelection;
                //         }
                //         else
                //         {
                //             nameCombo.SelectedIndex = 0; // Select "<Select>"
                //         }
                //     }
                // }
            }
            catch (Exception ex)
            {
                // Log error but don't crash
                try
                {
                    string logPath = SafeFileLogger.GetLogFilePath("parameter_refresh_error.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now}] Error refreshing parameters for categories {string.Join(", ", selectedCategories)}: {ex.Message}\n");
                }
                catch { }
            }
        }
        private void LoadProfileInfo()
        {
            try
            {
                var currentProfile = _appProfileService.GetCurrentProfile();
                if (currentProfile != null)
                {
                    _profileLabel.Text = $"Profile: {currentProfile.Name}";
                    DebugLogger.Info($"Profile loaded: {currentProfile.Name}");
                    
                    // CRITICAL: Load saved configuration into UI
                    LoadConfigurationFromProfile(currentProfile);
                }
                else
                {
                    _profileLabel.Text = "Profile: Default";
                    DebugLogger.Info("No profile found - using default");
                }
            }
            catch (Exception ex)
            {
                _profileLabel.Text = $"Profile: Error - {ex.Message}";
                DebugLogger.Error($"Profile loading error: {ex.Message}");
            }
        }

        private void LoadRealLinkedFiles(Document document)
        {
            // 🔍 DIAGNOSTIC: Log LoadRealLinkedFiles start
            try
            {
                DebugLogger.Info($"[{DateTime.Now}] 🔍 LoadRealLinkedFiles: STARTED\n");
                
                // Count sleeves BEFORE LoadRealLinkedFiles
                var sleevesBeforeCount = new FilteredElementCollector(document)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                    .Count();
                DebugLogger.Info($"[{DateTime.Now}] 🔍 LoadRealLinkedFiles: Sleeves BEFORE: {sleevesBeforeCount}\n");
            }
            catch { }
            
            DebugLogger.Info("=== STARTING LoadRealLinkedFiles ===");
            DebugLogger.Info($"Document: {document?.Title ?? "null"}");

            try
            {
                if (_linkedFileService != null)
                {
                    DebugLogger.Info("_linkedFileService is available");
                    _linkedFiles = _linkedFileService.GetLinkedFiles(document);
                    DebugLogger.Info($"Loaded {_linkedFiles.Count} linked files from Revit project");

                    // Check if no linked files were found and show a prominent message
                    if (_linkedFiles.Count == 0)
                    {
                        DebugLogger.Info("No linked files found - showing message");
                        ShowNoLinkedFilesMessage();
                        ShowStatusBanner();
                        // Keep the existing fallback labels already rendered
                    }
                    else
                    {
                        DebugLogger.Info($"Successfully loaded {_linkedFiles.Count} linked files from project");
                        foreach (var file in _linkedFiles)
                        {
                            DebugLogger.Info($"  - {file.FileName} ({file.FileType}) - {file.ElementCount} elements");
                        }

                        HideStatusBanner();
                        // Repopulate UI sections now that data is available
                        DebugLogger.Info("Calling RepopulateSectionsAfterLoad");
                        RepopulateSectionsAfterLoad();
                    }
                }
                else
                {
                    DebugLogger.Info("_linkedFileService is null!");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error loading linked files: {ex.Message}");
                DebugLogger.Error($"Stack trace: {ex.StackTrace}");
                ShowNoLinkedFilesMessage();
                ShowStatusBanner();
            }

                            DebugLogger.Info("=== LoadRealLinkedFiles COMPLETED ===");
                
                // 🔍 DIAGNOSTIC: Log LoadRealLinkedFiles completion
                try
                {
                    // Count sleeves AFTER LoadRealLinkedFiles
                    var sleevesAfterCount = new FilteredElementCollector(document)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                        .Count();
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 LoadRealLinkedFiles: Sleeves AFTER: {sleevesAfterCount}\n");
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 LoadRealLinkedFiles: COMPLETED\n");
                }
                catch { }
                
                // CRITICAL: Restore configuration AFTER UI is fully populated
                RestoreConfigurationToUI();
        }
        private void RepopulateSectionsAfterLoad()
        {
            DebugLogger.Info("=== STARTING RepopulateSectionsAfterLoad ===");

            if (this.InvokeRequired)
            {
                DebugLogger.Info("InvokeRequired - deferring to UI thread");
                this.BeginInvoke(new Action(RepopulateSectionsAfterLoad));
                return;
            }

            try
            {
                DebugLogger.Info($"_leftPanel exists: {_leftPanel != null}");
                DebugLogger.Info($"_leftPanel.Controls.Count before clear: {_leftPanel.Controls.Count}");

                // Rebuild the full left-side 2x2 layout to ensure splitters and inner right panels are present
                _leftPanel.Controls.Clear();
                DebugLogger.Info("_leftPanel.Controls cleared");

                CreateFourSectionLayout();
                DebugLogger.Info("CreateFourSectionLayout completed");

                _hostCategoriesInitialized =
                    (_horizontalCategoriesListBox != null && _horizontalCategoriesListBox.Items.Count > 0) ||
                    (_verticalCategoriesListBox != null && _verticalCategoriesListBox.Items.Count > 0);
                if (_hostCategoriesInitialized)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info("[UI-STATE] Host categories ready after repopulate - restoring UI selections");
                    LoadUIStateDirectly();
                }
                else if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info("[UI-STATE] Host categories still not available after repopulate - UI state restore deferred");
                }

                // CRITICAL: Reposition and resize panels after recreation
                PositionPanels();
                DebugLogger.Info("PositionPanels completed");

                BalanceLeftLayout();
                DebugLogger.Info("BalanceLeftLayout completed");

                DebugLogger.Info($"_leftPanel.Controls.Count after recreation: {_leftPanel.Controls.Count}");
                DebugLogger.Info($"_leftPanel.Width: {_leftPanel.Width}, Height: {_leftPanel.Height}");
                DebugLogger.Info($"_leftPanel.Location: {_leftPanel.Location}");

                // Log panel visibility and sizes
                if (_topLeftPanel != null)
                    DebugLogger.Info($"_topLeftPanel: Visible={_topLeftPanel.Visible}, Size={_topLeftPanel.Size}, Location={_topLeftPanel.Location}");
                if (_topRightPanel != null)
                    DebugLogger.Info($"_topRightPanel: Visible={_topRightPanel.Visible}, Size={_topRightPanel.Size}, Location={_topRightPanel.Location}");
                if (_bottomLeftPanel != null)
                    DebugLogger.Info($"_bottomLeftPanel: Visible={_bottomLeftPanel.Visible}, Size={_bottomLeftPanel.Size}, Location={_bottomLeftPanel.Location}");
                if (_bottomRightPanel != null)
                    DebugLogger.Info($"_bottomRightPanel: Visible={_bottomRightPanel.Visible}, Size={_bottomRightPanel.Size}, Location={_bottomRightPanel.Location}");

                DebugLogger.Info("=== RepopulateSectionsAfterLoad COMPLETED ===");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"ERROR in RepopulateSectionsAfterLoad: {ex.Message}");
                DebugLogger.Error($"Stack trace: {ex.StackTrace}");
            }
        }

        private void ShowNoLinkedFilesMessage()
        {
            var result = WinForms.MessageBox.Show(
                this,
                "No linked files were found in the current Revit project.\n\n" +
                "This dialog requires linked MEP, architectural, and structural files to function properly.\n\n" +
                "Please ensure you have linked the necessary files to your project before using this feature.\n\n" +
                "Would you like to continue anyway?",
                "No Linked Files Found",
                WinForms.MessageBoxButtons.YesNo,
                WinForms.MessageBoxIcon.Warning,
                WinForms.MessageBoxDefaultButton.Button2);

            if (result == WinForms.DialogResult.No)
            {
                this.Close();
            }
        }

        private void ShowStatusBanner()
        {
            if (_statusBannerPanel != null)
            {
                _statusBannerPanel.Visible = true;
                // Adjust other panels to account for the banner
                PositionPanels();
            }
        }

        private void HideStatusBanner()
        {
            if (_statusBannerPanel != null)
            {
                _statusBannerPanel.Visible = false;
                // Adjust other panels after hiding the banner
                PositionPanels();
            }
        }

        // Event handlers
        
        /// <summary>
        /// ⚠️ CRITICAL METHOD - DO NOT REMOVE ⚠️
        /// Save opening conditions to XML for each selected category/filter
        /// This implements proper architecture separation: CONDITIONS (XML) vs UI (static)
        /// </summary>
        private void SaveConditionsToXml(List<string> selectedCategories)
        {
            try
            {
                // Use project-specific Filters directory
                string projectFiltersDir = _document != null ? ProjectPathService.GetFiltersDirectory(_document) : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Projects", "Default", "Filters");
                var conditionsService = new ConditionsService(_document, projectFiltersDir, msg => DebugLogger.Info(msg));
                
                // Get selected filters
                var selectedFilters = GetSelectedFilters();
                
                foreach (var filter in selectedFilters)
                {
                    // Create conditions object from UI values
                    var conditions = new OpeningConditions
                    {
                        FilterName = filter.Name,
                        Category = filter.Category.ToString(), // Convert enum to string
                        ClearanceSettings = ReadClearanceSettingsFromUI(),
                        OpeningTypePreferences = ReadOpeningTypePreferencesFromUI()
                        // ⚠️ NOTE: MarkPrefixes are UI state, not opening conditions - handled separately
                    };
                    
                    // 🛡️ ARCHITECTURE FIX: Save CONDITIONS XML using combined key (FilterName_Category)
                    // This allows different clearance/opening types per category within the same filter
                    // ✅ STANDARDIZED: Use MepCategoryConstants.GetXmlSuffix() for consistent naming (same as RefreshService)
                    // ✅ FIX: Strip .xml extension from filter name if present (prevents duplicate database entries)
                    string cleanFilterName = filter.Name;
                    if (cleanFilterName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                    {
                        cleanFilterName = cleanFilterName.Substring(0, cleanFilterName.Length - 4);
                    }
                    string normalizedCategory = MepCategoryConstants.GetXmlSuffix(filter.Category.ToString());
                    string combinedKey = $"{cleanFilterName}_{normalizedCategory}";
                    bool saved = conditionsService.SaveConditions(conditions, combinedKey);
                    if (saved)
                    {
                        DebugLogger.Info($"[SaveConditionsToXml] Saved conditions for '{combinedKey}' (Filter: '{filter.Name}', Category: '{filter.Category}')");
                    }
                    else
                    {
                        DebugLogger.Warning($"[SaveConditionsToXml] Failed to save conditions for '{combinedKey}'");
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[SaveConditionsToXml] Error: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Read clearance settings from UI textboxes
        /// </summary>
        private ClearanceSettings ReadClearanceSettingsFromUI()
        {
            var settings = new ClearanceSettings();
            
            try
            {
                if (_clearancePanel != null)
                {
                    var allTextBoxes = _clearancePanel.Controls.OfType<WinForms.TextBox>();
                    
                    // Rectangular duct clearances
                    var normalTb = allTextBoxes.FirstOrDefault(tb => tb.Tag?.ToString() == "normal_clearance");
                    if (normalTb != null && double.TryParse(normalTb.Text, out double normalVal))
                    {
                        settings.RectangularNormal = normalVal;
                    }
                    
                    var insulatedTb = allTextBoxes.FirstOrDefault(tb => tb.Tag?.ToString() == "insulated_clearance");
                    if (insulatedTb != null && double.TryParse(insulatedTb.Text, out double insulatedVal))
                    {
                        settings.RectangularInsulated = insulatedVal;
                    }
                    
                    // Round duct clearances
                    var roundNormalTb = allTextBoxes.FirstOrDefault(tb => tb.Tag?.ToString() == "round_duct_normal_clearance");
                    if (roundNormalTb != null && double.TryParse(roundNormalTb.Text, out double roundNormalVal))
                    {
                        settings.RoundNormal = roundNormalVal;
                        DebugLogger.Info($"[ReadClearanceSettingsFromUI] Round Normal: TextBox.Text='{roundNormalTb.Text}' → Parsed={roundNormalVal}mm");
                    }
                    else
                    {
                        DebugLogger.Warning($"[ReadClearanceSettingsFromUI] Round Normal textbox not found or parse failed");
                    }
                    
                    var roundInsulatedTb = allTextBoxes.FirstOrDefault(tb => tb.Tag?.ToString() == "round_duct_insulated_clearance");
                    if (roundInsulatedTb != null && double.TryParse(roundInsulatedTb.Text, out double roundInsulatedVal))
                    {
                        settings.RoundInsulated = roundInsulatedVal;
                        DebugLogger.Info($"[ReadClearanceSettingsFromUI] Round Insulated: TextBox.Text='{roundInsulatedTb.Text}' → Parsed={roundInsulatedVal}mm");
                    }
                    else
                    {
                        DebugLogger.Warning($"[ReadClearanceSettingsFromUI] Round Insulated textbox not found or parse failed");
                    }
                }
                
                // Read damper clearances
                if (_damperPanel != null)
                {
                    var damperTextBoxes = _damperPanel.Controls.OfType<WinForms.TextBox>();
                    
                    var damperMepTb = damperTextBoxes.FirstOrDefault(tb => tb.Tag?.ToString() == "ductaccessories_mep_normal");
                    if (damperMepTb != null && double.TryParse(damperMepTb.Text, out double damperMepVal))
                    {
                        settings.DuctAccessoryMepNormal = damperMepVal;
                        DebugLogger.Info($"[ReadClearanceSettingsFromUI] Damper MEP: {damperMepVal}mm");
                    }
                    
                    var damperOtherTb = damperTextBoxes.FirstOrDefault(tb => tb.Tag?.ToString() == "ductaccessories_other_normal");
                    if (damperOtherTb != null && double.TryParse(damperOtherTb.Text, out double damperOtherVal))
                    {
                        settings.DuctAccessoryOtherNormal = damperOtherVal;
                        DebugLogger.Info($"[ReadClearanceSettingsFromUI] Damper Other: {damperOtherVal}mm");
                    }
                }
                
                // Read cable tray clearances
                if (_cableTrayPanel != null)
                {
                    var cableTrayTextBoxes = _cableTrayPanel.Controls.OfType<WinForms.TextBox>();
                    
                    var cableTrayTopTb = cableTrayTextBoxes.FirstOrDefault(tb => tb.Tag?.ToString() == "cabletray_top_clearance");
                    if (cableTrayTopTb != null && double.TryParse(cableTrayTopTb.Text, out double cableTrayTopVal))
                    {
                        settings.CableTrayTop = cableTrayTopVal;
                        DebugLogger.Info($"[ReadClearanceSettingsFromUI] Cable Tray Top: {cableTrayTopVal}mm");
                    }
                    
                    var cableTrayOtherTb = cableTrayTextBoxes.FirstOrDefault(tb => tb.Tag?.ToString() == "cabletray_other_clearance");
                    if (cableTrayOtherTb != null && double.TryParse(cableTrayOtherTb.Text, out double cableTrayOtherVal))
                    {
                        settings.CableTrayOther = cableTrayOtherVal;
                        DebugLogger.Info($"[ReadClearanceSettingsFromUI] Cable Tray Other: {cableTrayOtherVal}mm");
                    }
                }
                
                // Read pipe clearances from pipe panel
                if (_pipePanel != null)
                {
                    var pipeTextBoxes = _pipePanel.Controls.OfType<WinForms.TextBox>();
                    
                    var pipeNormalTb = pipeTextBoxes.FirstOrDefault(tb => tb.Tag?.ToString() == "pipes_normal_clearance");
                    if (pipeNormalTb != null && double.TryParse(pipeNormalTb.Text, out double pipeNormalVal))
                    {
                        settings.PipesNormal = pipeNormalVal;
                        DebugLogger.Info($"[ReadClearanceSettingsFromUI] Pipe Normal: {pipeNormalVal}mm");
                    }
                    else
                    {
                        DebugLogger.Warning($"[ReadClearanceSettingsFromUI] Pipe Normal textbox not found or parse failed - using default");
                        settings.PipesNormal = settings.RectangularNormal; // Default to duct normal
                    }
                    
                    var pipeInsTb = pipeTextBoxes.FirstOrDefault(tb => tb.Tag?.ToString() == "pipes_insulated_clearance");
                    if (pipeInsTb != null && double.TryParse(pipeInsTb.Text, out double pipeInsVal))
                    {
                        settings.PipesInsulated = pipeInsVal;
                        DebugLogger.Info($"[ReadClearanceSettingsFromUI] Pipe Insulated: {pipeInsVal}mm");
                    }
                    else
                    {
                        DebugLogger.Warning($"[ReadClearanceSettingsFromUI] Pipe Insulated textbox not found or parse failed - using default");
                        settings.PipesInsulated = settings.RectangularInsulated; // Default to duct insulated
                    }
                }
                else
                {
                    DebugLogger.Warning($"[ReadClearanceSettingsFromUI] _pipePanel is null - using duct values as fallback");
                    settings.PipesNormal = settings.RectangularNormal; // Default to duct normal
                    settings.PipesInsulated = settings.RectangularInsulated; // Default to duct insulated
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[ReadClearanceSettingsFromUI] Error: {ex.Message}");
            }
            
            return settings;
        }
        
        /// <summary>
        /// Read opening type preferences from UI radio buttons
        /// </summary>
        private OpeningTypePreferences ReadOpeningTypePreferencesFromUI()
        {
            var preferences = new OpeningTypePreferences();
            
            try
            {
                // Get duct opening type selection
                var ductOpeningType = GetDuctOpeningType();
                if (!string.IsNullOrEmpty(ductOpeningType))
                {
                    preferences.RoundDucts = ductOpeningType;
                }
                
                // Get pipe opening type selection (if pipe panel exists)
                var pipeOpeningType = GetPipeOpeningType();
                if (!string.IsNullOrEmpty(pipeOpeningType))
                {
                    preferences.Pipes = pipeOpeningType;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[ReadOpeningTypePreferencesFromUI] Error: {ex.Message}");
            }
            
            return preferences;
        }

        // ✅ REMOVED: NormalizeCategoryName - Now using MepCategoryConstants.GetXmlSuffix() for consistency

        /// <summary>
        /// Read mark prefix settings from UI textboxes
        /// ⚠️ UI STATE METHOD - Reads Project Prefix and Discipline Prefix from UI for mark parameter command
        /// </summary>
        private MarkPrefixSettings ReadMarkPrefixesFromUI()
        {
            var markPrefixes = new MarkPrefixSettings();
            
            try
            {
                // Read Project Prefix from UI
                if (_projectPrefixTextBox != null && !string.IsNullOrWhiteSpace(_projectPrefixTextBox.Text))
                {
                    markPrefixes.ProjectPrefix = _projectPrefixTextBox.Text.Trim();
                }
                else
                {
                    markPrefixes.ProjectPrefix = "SLEEVE_"; // Default fallback
                }

                // ✅ FIX: Use MarkPrefixService to get type-specific prefixes from ParameterServiceDialog
                // This ensures we get the current values the user has entered in the ParameterServiceDialog
                var currentPrefixes = MarkPrefixService.GetCurrentPrefixes();
                markPrefixes.DuctPrefix = currentPrefixes.DuctPrefix;
                markPrefixes.PipePrefix = currentPrefixes.PipePrefix;
                markPrefixes.CableTrayPrefix = currentPrefixes.CableTrayPrefix;
                markPrefixes.DamperPrefix = currentPrefixes.DamperPrefix;
                
                // ✅ NEW: Read "Re-mark all" checkbox state
                markPrefixes.RemarkAll = _remarkAllCheckBox?.Checked ?? false;

                DebugLogger.Info($"[ReadMarkPrefixesFromUI] Project Prefix: '{markPrefixes.ProjectPrefix}', Discipline Prefix: '{_disciplinePrefixTextBox?.Text ?? "null"}', Re-mark all: {markPrefixes.RemarkAll}");
                DebugLogger.Info($"[ReadMarkPrefixesFromUI] Final prefixes - Duct: '{markPrefixes.DuctPrefix}', Pipe: '{markPrefixes.PipePrefix}', CableTray: '{markPrefixes.CableTrayPrefix}', Damper: '{markPrefixes.DamperPrefix}'");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[ReadMarkPrefixesFromUI] Error: {ex.Message}");
                // Return defaults on error
                markPrefixes.ProjectPrefix = "SLEEVE_";
                markPrefixes.DuctPrefix = "DCT";
                markPrefixes.PipePrefix = "PLU";
                markPrefixes.CableTrayPrefix = "ELE";
                markPrefixes.DamperPrefix = "DAM";
            }
            
            return markPrefixes;
        }
        
        private void OnOkClick(object? sender, EventArgs e)
        {
            try
            {
                var __okStart = DateTime.Now;
                SafeFileLogger.SafeAppendText("performance.log", $"OK_START {__okStart:O}");
                try { SafeFileLogger.SafeAppendText("placement_event_trace.log", $"[{DateTime.Now:HH:mm:ss}] CLICK_OK: handler entered\n"); } catch { }
                
                // ✅ SAVE: Save current clearance values before OK button processing
                SaveCurrentClearanceValues();
                
                // Parameter Transfer button removed - functionality moved to Transfer All button

                // Get selected categories
                var selectedCategories = GetSelectedMepCategories();
                if (selectedCategories.Count == 0)
                {
                    MessageBox.Show("Please select at least one MEP category.", "No Categories Selected", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                // Mark prefix functionality moved to Parameter Service UI
                // Users can apply marks after placement via Parameter Service dialog
                
                // ⚠️ CRITICAL: Save CONDITIONS.xml before raising external event ⚠️
                // This implements proper architecture: Conditions saved to XML, not in static properties
                SaveConditionsToXml(selectedCategories);
                
                // Show placement progress on main UI
                StartPlacementProgress($"Placing sleeves for {selectedCategories.Count} categories...");
                _okButton.Enabled = false;
                _refreshButton.Enabled = false;
                DebugLogger.Info($"[OnOkClick] Starting sleeve placement for categories: {string.Join(", ", selectedCategories)}");

                // Initialize RevitTask if not already done
                if (!RevitTask.IsInitialized())
                {
                    var handler = new RevitTask.Handler();
                    var exEvent = ExternalEvent.Create(handler);
                    RevitTask.Init(exEvent);
                }

                // Pass categories and mark prefixes to external event handler
                // Get mark prefixes from ParameterServiceDialog via MarkPrefixService
                var markPrefixes = MarkPrefixService.GetCurrentPrefixes();
                var selectedFilterNames = GetSelectedFilterItems();
                var selectedFilterName = selectedFilterNames.Count > 0 ? selectedFilterNames[0] : "Default";
                _sleevePlacementHandler.SetContext(selectedCategories, markPrefixes, selectedFilterName);

                try { SafeFileLogger.SafeAppendText("placement_event_trace.log", $"[{DateTime.Now:HH:mm:ss}] CLICK_OK: raising external event\n"); } catch { }

                // Hide UI while placement runs to prevent re-entrancy and follow transaction guidelines
                this.Hide();

                // Raise external event (non-blocking)
                _sleevePlacementEvent.Raise();

                DebugLogger.Info($"[EmergencyMainDialog] External event raised for categories: {string.Join(", ", selectedCategories)}");
                try { SafeFileLogger.SafeAppendText("placement_event_trace.log", $"[{DateTime.Now:HH:mm:ss}] CLICK_OK: external event raised\n"); } catch { }

                var __okEnd = DateTime.Now;
                var __okMs = (long)(__okEnd - __okStart).TotalMilliseconds;
                SafeFileLogger.SafeAppendText("performance.log", $"OK_END {__okEnd:O} DURATION_MS {__okMs}");
            }
            catch (Exception ex)
            {
                CompletePlacementProgress($"Error starting sleeve placement: {ex.Message}");
                this.Show();
                _okButton.Enabled = true;
                _refreshButton.Enabled = true;
                DebugLogger.Error($"[EmergencyMainDialog] Exception in OnOkClick: {ex.Message}");
                try { SafeFileLogger.SafeAppendText("placement_event_trace.log", $"[{DateTime.Now:HH:mm:ss}] CLICK_OK: exception {ex.Message}\n"); } catch { }
                MessageBox.Show($"Error starting sleeve placement: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private bool ValidateConfiguration()
        {
            try
            {
                // Validate that filters are selected
                var selectedFilters = GetSelectedFilters();
                if (selectedFilters == null || selectedFilters.Count == 0)
                {
                    MessageBox.Show("Please select at least one filter before proceeding.", "No Filters Selected", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
                
                // Validate that reference and host elements are selected
                // TODO: Implement when _referenceElementsListBox and _hostElementsListBox are available
                // if (_referenceElementsListBox.SelectedItems.Count == 0)
                // {
                //     MessageBox.Show("Please select reference elements before proceeding.", "No Reference Elements", 
                //         MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //     return false;
                // }
                // 
                // if (_hostElementsListBox.SelectedItems.Count == 0)
                // {
                //     MessageBox.Show("Please select host elements before proceeding.", "No Host Elements", 
                //         MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //     return false;
                // }
                
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Configuration validation failed: {ex.Message}", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        
        private OrchestrationResult ExecuteSelectedFiltersWithProgress()
        {
            try
            {
                // CORRECT SOLUTION: External Event is just a transaction context bridge
                DebugLogger.Info("ExecuteSelectedFiltersWithProgress: Using External Event as transaction bridge");
                
                // Get selected MEP categories from UI
                var selectedCategories = GetSelectedMepCategories();
                
                if (selectedCategories == null || selectedCategories.Count == 0)
                {
                    _statusLabel.Text = "No MEP categories selected";
                    DebugLogger.Warning("ExecuteSelectedFiltersWithProgress: No MEP categories selected");
                    
                    return new OrchestrationResult
                    {
                        Success = false,
                        ErrorMessage = "No MEP categories selected"
                    };
                }
                
                // Set selected categories for External Event to process
                _sleevePlacementHandler.SetSelectedCategories(selectedCategories);
                
                DebugLogger.Info($"[EXTERNAL_EVENT] Passing categories to External Event: {string.Join(", ", selectedCategories)}");
                
                // CRITICAL: Don't show success immediately - External Event is asynchronous
                _statusLabel.Text = $"Raising external event for {selectedCategories.Count} categories...";
                
                _sleevePlacementEvent.Raise();
                
                _statusLabel.Text = $"External event raised - processing {selectedCategories.Count} categories...";
                
                // CRITICAL: Return pending status instead of success
                return new OrchestrationResult
                {
                    Success = true,
                    Message = $"External event raised to process categories: {string.Join(", ", selectedCategories)}"
                };
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"ExecuteSelectedFiltersWithProgress error: {ex.Message}");
                return new OrchestrationResult
                {
                    Success = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        private void ShowProgressDialog()
        {
            try
            {
                // Create a simple progress dialog to show what's happening
                var progressDialog = new WinForms.Form
                {
                    Text = "Opening Creation Progress",
                    Size = new System.Drawing.Size(400, 200),
                    StartPosition = WinForms.FormStartPosition.CenterParent,
                    FormBorderStyle = WinForms.FormBorderStyle.FixedDialog,
                    MaximizeBox = false,
                    MinimizeBox = false,
                    ShowInTaskbar = false
                };

                var progressLabel = new WinForms.Label
                {
                    Text = "Processing opening creation...",
                    Location = new System.Drawing.Point(20, 20),
                    Size = new System.Drawing.Size(350, 40),
                    Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold)
                };
                progressDialog.Controls.Add(progressLabel);

                var statusLabel = new WinForms.Label
                {
                    Text = "This may take several minutes. Check log files for detailed progress.",
                    Location = new System.Drawing.Point(20, 70),
                    Size = new System.Drawing.Size(350, 60),
                    Font = new System.Drawing.Font("Microsoft Sans Serif", 9F)
                };
                progressDialog.Controls.Add(statusLabel);

                var closeButton = new WinForms.Button
                {
                    Text = "Close",
                    Location = new System.Drawing.Point(300, 140),
                    Size = new System.Drawing.Size(70, 30),
                    DialogResult = WinForms.DialogResult.OK
                };
                progressDialog.Controls.Add(closeButton);

                // Show dialog modeless so processing can continue
                progressDialog.Show();
                
                // Auto-close after 3 seconds
                var timer = new System.Windows.Forms.Timer();
                timer.Interval = 3000; // 3 seconds
                timer.Tick += (s, e) => {
                    timer.Stop();
                    progressDialog.Close();
                };
                timer.Start();
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error showing progress dialog: {ex.Message}");
                // Fallback to simple message box
                MessageBox.Show("Processing opening creation...", "Progress", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void OnSaveClick(object? sender, EventArgs e)
        {
            try
            {
                DebugLogger.Info("=== STARTING Save Configuration ===");
                _statusLabel.Text = "Saving configuration...";
                
                // Check if we have a current profile
                var currentProfile = _appProfileService.GetCurrentProfile();
                DebugLogger.Info($"Current profile: {(currentProfile != null ? currentProfile.Name : "NULL")}");
                
                if (currentProfile == null)
                {
                    DebugLogger.Warning("No current profile - creating emergency profile");
                    // Create an emergency profile if none exists
                    var emergencyProfile = new UserProfile
                    {
                        Id = Guid.NewGuid(),
                        Name = "Emergency Profile",
                        Disciplines = new List<Discipline> { new Discipline("Mechanical", true, "Mechanical systems") },
                        Language = "English",
                        CreatedDate = DateTime.Now,
                        IsActive = true
                    };
                    
                    _appProfileService.SetCurrentProfile(emergencyProfile);
                    DebugLogger.Info($"Created emergency profile: {emergencyProfile.Name}");
                }
                
                // Save current profile with UI state and settings
                SaveCurrentConfiguration();
                
                // SIMPLE FIX: Save UI state directly to a simple file
                SaveUIStateDirectly();
                
                _statusLabel.Text = "Configuration saved successfully";
                DebugLogger.Info("=== Save Configuration COMPLETED ===");
            }
            catch (Exception ex)
            {
                _statusLabel.Text = $"Save failed: {ex.Message}";
                DebugLogger.Error($"Save configuration failed: {ex.Message}");
            }
        }
        private void SaveCurrentConfiguration()
        {
            try
            {
                // Get current profile
                var currentProfile = _appProfileService.GetCurrentProfile();
                if (currentProfile == null)
                {
                    DebugLogger.Warning("No current profile found - user must create a profile first");
                    // Don't create a default profile - let user create their own
                    return;
                }
                
                // Collect current UI state
                var configuration = CollectCurrentUIState();
                
                // Update profile with current configuration
                currentProfile.Configuration = configuration;
                currentProfile.LastModified = DateTime.Now;
                
                // Debug: Verify configuration was set
                DebugLogger.Info($"Profile {currentProfile.Name} configuration set with {configuration.SelectedReferenceFiles?.Count ?? 0} reference files");
                DebugLogger.Info($"Profile {currentProfile.Name} configuration object: {(currentProfile.Configuration != null ? "NOT NULL" : "NULL")}");
                
                // Save profile to disk
                DebugLogger.Info($"[SAVE_DEBUG] About to call SaveCurrentProfile for profile: {currentProfile.Name}");
                DebugLogger.Info($"[SAVE_DEBUG] Profile configuration is null: {currentProfile.Configuration == null}");
                if (currentProfile.Configuration?.OpeningSettings?.ClearanceSettings != null)
                {
                    DebugLogger.Info($"[SAVE_DEBUG] Clearance settings count: {currentProfile.Configuration.OpeningSettings.ClearanceSettings.Count}");
                    foreach (var kvp in currentProfile.Configuration.OpeningSettings.ClearanceSettings)
                    {
                        DebugLogger.Info($"[SAVE_DEBUG] Clearance setting: {kvp.Key} = {kvp.Value}");
                    }
                }
                
                _appProfileService.SaveCurrentProfile();
                
                DebugLogger.Info($"Configuration saved for profile: {currentProfile.Name}");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to save configuration: {ex.Message}");
                throw;
            }
        }
        private UserConfiguration CollectCurrentUIState()
        {
            var config = new UserConfiguration();
            
            try
            {
                // Collect selected linked files
                config.SelectedReferenceFiles = GetSelectedReferenceFiles();
                config.SelectedHostFiles = GetSelectedHostFiles();
                
                // Collect selected MEP categories
                config.SelectedMepCategories = GetSelectedMepCategories();
                
                // Collect selected host categories
                config.SelectedHostCategories = GetSelectedHostCategories();
                
                // Collect opening settings
                config.OpeningSettings = GetCurrentOpeningSettings();
                
                DebugLogger.Info($"=== UI STATE COLLECTION SUMMARY ===");
                DebugLogger.Info($"Reference files: {config.SelectedReferenceFiles?.Count ?? 0}");
                DebugLogger.Info($"Host files: {config.SelectedHostFiles?.Count ?? 0}");
                DebugLogger.Info($"MEP categories: {config.SelectedMepCategories?.Count ?? 0}");
                DebugLogger.Info($"Host categories: {config.SelectedHostCategories?.Count ?? 0}");
                if (config.SelectedReferenceFiles?.Count > 0)
                {
                    foreach (var file in config.SelectedReferenceFiles)
                        DebugLogger.Info($"  Selected reference file: {file}");
                }
                if (config.SelectedMepCategories?.Count > 0)
                {
                    foreach (var cat in config.SelectedMepCategories)
                        DebugLogger.Info($"  Selected MEP category: {cat}");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to collect UI state: {ex.Message}");
                throw;
            }
            
            return config;
        }
        
        private void LoadConfigurationFromProfile(UserProfile profile)
        {
            try
            {
                DebugLogger.Info($"=== STARTING LoadConfigurationFromProfile for {profile.Name} ===");
                
                if (profile.Configuration == null)
                {
                    // Try to load configuration from JSON file
                    LoadConfigurationFromJsonFile(profile);
                    
                    if (profile.Configuration == null)
                    {
                        DebugLogger.Info("No saved configuration found - using defaults");
                        return;
                    }
                }
                
                var config = profile.Configuration;
                DebugLogger.Info($"Configuration loaded: {config.SelectedReferenceFiles?.Count ?? 0} reference files, {config.SelectedMepCategories?.Count ?? 0} MEP categories");
                
                // Restore clearance settings if available
                if (config.OpeningSettings?.ClearanceSettings != null && config.OpeningSettings.ClearanceSettings.Count > 0)
                {
                    RestoreClearanceSettings(config.OpeningSettings.ClearanceSettings);
                    DebugLogger.Info($"Restored {config.OpeningSettings.ClearanceSettings.Count} clearance settings");
                }
                else
                {
                    DebugLogger.Info("No clearance settings found in profile - using defaults");
                }
                
                // TODO: Implement restoration methods for MEP categories, reference files, and host files
                DebugLogger.Info($"Profile configuration loaded: {config.SelectedMepCategories?.Count ?? 0} MEP categories, {config.SelectedReferenceFiles?.Count ?? 0} reference files, {config.SelectedHostFiles?.Count ?? 0} host files");
                
                DebugLogger.Info($"=== LoadConfigurationFromProfile COMPLETED for {profile.Name} ===");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to load configuration from profile {profile.Name}: {ex.Message}");
                // Don't throw - we can continue with defaults
            }
        }
        /// <summary>
        /// Restore clearance settings from saved configuration
        /// </summary>
        private void RestoreClearanceSettings(Dictionary<string, double> clearanceSettings)
        {
            try
            {
                DebugLogger.Info($"[CLEARANCE_RESTORE] Restoring {clearanceSettings.Count} clearance settings");
                
                // Restore clearance values in clearance panel
                if (_clearancePanel?.Controls.Count > 0)
                {
                    foreach (var control in _clearancePanel.Controls)
                    {
                        if (control is WinForms.TextBox textBox && textBox.Tag != null)
                        {
                            string genericKey = textBox.Tag.ToString() ?? "";
                            
                            // Try to find matching clearance setting by checking all possible category combinations
                            bool restored = false;
                            foreach (var kvp in clearanceSettings)
                            {
                                // Check if this clearance setting matches our generic key for any category
                                if (kvp.Key.EndsWith($"_{genericKey}") || kvp.Key == genericKey)
                                {
                                    textBox.Text = kvp.Value.ToString();
                                    DebugLogger.Info($"[CLEARANCE_RESTORE] Restored {genericKey} -> {kvp.Key} = {kvp.Value}mm");
                                    restored = true;
                                    break;
                                }
                            }
                            
                            if (!restored)
                            {
                                DebugLogger.Info($"[CLEARANCE_RESTORE] No saved value found for {genericKey}");
                            }
                        }
                    }
                }
                
                // Restore cable tray clearance values
                if (_cableTrayPanel?.Controls.Count > 0)
                {
                    foreach (var control in _cableTrayPanel.Controls)
                    {
                        if (control is WinForms.TextBox textBox && textBox.Tag != null)
                        {
                            string key = textBox.Tag.ToString() ?? "";
                            if (clearanceSettings.TryGetValue(key, out double value))
                            {
                                textBox.Text = value.ToString();
                                DebugLogger.Info($"[CLEARANCE_RESTORE] Restored cable tray {key} = {value}mm");
                            }
                        }
                    }
                }
                
                // Restore damper clearance values
                if (_damperPanel?.Controls.Count > 0)
                {
                    foreach (var control in _damperPanel.Controls)
                    {
                        if (control is WinForms.TextBox textBox && textBox.Tag != null)
                        {
                            string key = textBox.Tag.ToString() ?? "";
                            if (clearanceSettings.TryGetValue(key, out double value))
                            {
                                textBox.Text = value.ToString();
                                DebugLogger.Info($"[CLEARANCE_RESTORE] Restored damper {key} = {value}mm");
                            }
                        }
                    }
                }
                
                DebugLogger.Info($"[CLEARANCE_RESTORE] Clearance restoration completed");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[CLEARANCE_RESTORE] Failed to restore clearance settings: {ex.Message}");
            }
        }
        private void RestoreReferenceFileSelections(List<string>? selectedFiles)
        {
            if (selectedFiles == null || selectedFiles.Count == 0) 
            {
                DebugLogger.Info("No reference files to restore");
                return;
            }
            
            try
            {
                DebugLogger.Info($"RestoreReferenceFileSelections: Looking for {selectedFiles.Count} files");
                DebugLogger.Info($"TopLeftPanel exists: {_topLeftPanel != null}, Controls count: {_topLeftPanel?.Controls.Count ?? 0}");
                
                if (_topLeftPanel?.Controls.Count > 0)
                {
                    foreach (var control in _topLeftPanel.Controls)
                    {
                        if (control is WinForms.CheckedListBox listBox)
                        {
                            DebugLogger.Info($"Found CheckedListBox with {listBox.Items.Count} items");
                            
                            // Use BeginUpdate/EndUpdate for proper refresh
                            listBox.BeginUpdate();
                            
                            // CRITICAL: First uncheck ALL items, then check only the saved ones
                            for (int i = 0; i < listBox.Items.Count; i++)
                            {
                                listBox.SetItemChecked(i, false);
                            }
                            
                            // Debug: Log all items in the list box
                            DebugLogger.Info($"ListBox contains {listBox.Items.Count} items:");
                            for (int j = 0; j < listBox.Items.Count; j++)
                            {
                                DebugLogger.Info($"  Item {j}: {listBox.Items[j]}");
                            }
                            
                            // Debug: Log all selected files to restore
                            DebugLogger.Info($"Trying to restore {selectedFiles.Count} files:");
                            foreach (var file in selectedFiles)
                            {
                                DebugLogger.Info($"  Need to restore: {file}");
                            }
                            
                            // Now check only the saved selections
                            for (int i = 0; i < listBox.Items.Count; i++)
                            {
                                var item = listBox.Items[i]?.ToString();
                                if (item != null && selectedFiles.Contains(item))
                                {
                                    listBox.SetItemChecked(i, true);
                                    DebugLogger.Info($"Restored reference file selection: {item} (index {i})");
                                }
                                else if (item != null)
                                {
                                    DebugLogger.Info($"File not found for restoration: {item}");
                                }
                            }
                            
                            listBox.EndUpdate();
                            
                            // Force multiple refresh methods
                            listBox.Refresh();
                            listBox.Invalidate();
                            listBox.Update();
                        }
                    }
                }
                else
                {
                    DebugLogger.Warning("TopLeftPanel has no controls - cannot restore reference file selections");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to restore reference file selections: {ex.Message}");
            }
        }
        
        private void RestoreHostFileSelections(List<string>? selectedFiles)
        {
            if (selectedFiles == null || selectedFiles.Count == 0) 
            {
                DebugLogger.Info("[FILTER_HOST] No host files to restore - clearing all selections");
                // Clear all host file selections
                if (_bottomLeftPanel?.Controls.Count > 0)
                {
                    foreach (var control in _bottomLeftPanel.Controls)
                    {
                        if (control is WinForms.CheckedListBox listBox)
                        {
                            listBox.BeginUpdate();
                            for (int i = 0; i < listBox.Items.Count; i++)
                            {
                                listBox.SetItemChecked(i, false);
                            }
                            listBox.EndUpdate();
                        }
                    }
                }
                return;
            }
            
            try
            {
                DebugLogger.Info($"[FILTER_HOST] Restoring {selectedFiles.Count} host files: {string.Join(", ", selectedFiles)}");
                
                if (_bottomLeftPanel?.Controls.Count > 0)
                {
                    foreach (var control in _bottomLeftPanel.Controls)
                    {
                        if (control is WinForms.CheckedListBox listBox)
                        {
                            DebugLogger.Info($"[FILTER_HOST] Found CheckedListBox with {listBox.Items.Count} items");
                            
                            // Debug: Log all items in the list box
                            DebugLogger.Info($"[FILTER_HOST] ListBox contains {listBox.Items.Count} items:");
                            for (int j = 0; j < listBox.Items.Count; j++)
                            {
                                DebugLogger.Info($"[FILTER_HOST]   Item {j}: {listBox.Items[j]}");
                            }
                            
                            listBox.BeginUpdate();
                            
                            // CRITICAL: First uncheck ALL items
                            DebugLogger.Info("[FILTER_HOST] Clearing all host file selections first");
                            for (int i = 0; i < listBox.Items.Count; i++)
                            {
                                listBox.SetItemChecked(i, false);
                            }
                            
                            // Then check only the saved selections
                            DebugLogger.Info($"[FILTER_HOST] Restoring {selectedFiles.Count} saved selections:");
                            int restoredCount = 0;
                            for (int i = 0; i < listBox.Items.Count; i++)
                            {
                                var item = listBox.Items[i]?.ToString();
                                if (item != null && selectedFiles.Contains(item))
                                {
                                    listBox.SetItemChecked(i, true);
                                    DebugLogger.Info($"[FILTER_HOST] Restored host file selection: {item}");
                                    restoredCount++;
                                }
                            }
                            
                            DebugLogger.Info($"[FILTER_HOST] Successfully restored {restoredCount} out of {selectedFiles.Count} host file selections");
                            
                            listBox.EndUpdate();
                            listBox.Refresh();
                            listBox.Invalidate();
                            listBox.Update();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[FILTER_HOST] Failed to restore host file selections: {ex.Message}");
            }
        }
        
        private void RestoreMepCategorySelections(List<string>? selectedCategories)
        {
            if (selectedCategories == null || selectedCategories.Count == 0) 
            {
                DebugLogger.Info("[FILTER_UI] RestoreMepCategorySelections: No categories to restore");
                return;
            }
            
            try
            {
                DebugLogger.Info($"[FILTER_UI] RestoreMepCategorySelections: Restoring {selectedCategories.Count} categories: {string.Join(", ", selectedCategories)}");
                if (_topRightPanel?.Controls.Count > 0)
                {
                    foreach (var control in _topRightPanel.Controls)
                    {
                        if (control is WinForms.CheckedListBox listBox)
                        {
                            listBox.BeginUpdate();
                            
                            // First uncheck ALL items
                            for (int i = 0; i < listBox.Items.Count; i++)
                            {
                                listBox.SetItemChecked(i, false);
                            }
                            
                            // Then check only the saved selections
                            for (int i = 0; i < listBox.Items.Count; i++)
                            {
                                var item = listBox.Items[i]?.ToString();
                                if (item != null && selectedCategories.Contains(item))
                                {
                                    listBox.SetItemChecked(i, true);
                                    DebugLogger.Info($"Restored MEP category selection: {item}");
                                }
                            }
                            
                            listBox.EndUpdate();
                            listBox.Refresh();
                            listBox.Invalidate();
                            listBox.Update();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to restore MEP category selections: {ex.Message}");
            }
        }
        
        private void RestoreHostCategorySelections(List<string>? selectedCategories)
        {
            if (selectedCategories == null || selectedCategories.Count == 0) return;
            
            try
            {
                // ✅ CRITICAL FIX: Distinguish between horizontal and vertical listboxes
                // Horizontal: Walls, Structural Framing
                // Vertical: Floors, Ceilings
                var horizontalCategories = new[] { "Walls", "Structural Framing" };
                var verticalCategories = new[] { "Floors", "Ceilings" };
                
                if (_horizontalCategoriesListBox != null)
                {
                    _horizontalCategoriesListBox.BeginUpdate();
                    for (int i = 0; i < _horizontalCategoriesListBox.Items.Count; i++)
                    {
                        var item = _horizontalCategoriesListBox.Items[i]?.ToString();
                        bool shouldCheck = item != null && selectedCategories.Contains(item, StringComparer.OrdinalIgnoreCase) &&
                                          horizontalCategories.Contains(item, StringComparer.OrdinalIgnoreCase);
                        _horizontalCategoriesListBox.SetItemChecked(i, shouldCheck);
                        if (shouldCheck)
                            DebugLogger.Info($"[RESTORE] Restored horizontal host category: {item}");
                    }
                    _horizontalCategoriesListBox.EndUpdate();
                }
                
                if (_verticalCategoriesListBox != null)
                {
                    _verticalCategoriesListBox.BeginUpdate();
                    for (int i = 0; i < _verticalCategoriesListBox.Items.Count; i++)
                    {
                        var item = _verticalCategoriesListBox.Items[i]?.ToString();
                        bool shouldCheck = item != null && selectedCategories.Contains(item, StringComparer.OrdinalIgnoreCase) &&
                                          verticalCategories.Contains(item, StringComparer.OrdinalIgnoreCase);
                        _verticalCategoriesListBox.SetItemChecked(i, shouldCheck);
                        if (shouldCheck)
                            DebugLogger.Info($"[RESTORE] Restored vertical host category: {item}");
                    }
                    _verticalCategoriesListBox.EndUpdate();
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to restore host category selections: {ex.Message}");
            }
        }
        
        private void RestoreOpeningSettings(OpeningSettings? settings)
        {
            if (settings == null) return;
            
            try
            {
                // Note: Opening settings restoration would go here
                // For now, just log that we received settings
                DebugLogger.Info($"Opening settings available for restoration");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to restore opening settings: {ex.Message}");
            }
        }
        
        private void RestoreConfigurationToUI()
        {
            try
            {
                var currentProfile = _appProfileService.GetCurrentProfile();
                
                // Load configuration from text file if profile exists but configuration is null
                if (currentProfile != null && currentProfile.Configuration == null)
                {
                    LoadConfigurationFromJsonFile(currentProfile);
                }
                
                // SIMPLE FIX: Also load UI state directly
                if (_hostCategoriesInitialized)
                {
                    LoadUIStateDirectly();
                }
                else if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info("[UI-STATE] Host categories not initialized during configuration restore - waiting for repopulate");
                }
                
                if (currentProfile?.Configuration != null)
                {
                    DebugLogger.Info("=== STARTING RestoreConfigurationToUI ===");
                    
                    var config = currentProfile.Configuration;
                    DebugLogger.Info($"Restoring UI with: {config.SelectedReferenceFiles?.Count ?? 0} reference files, {config.SelectedMepCategories?.Count ?? 0} MEP categories");
                    
                    // Restore selected reference files (top-left panel)
                    RestoreReferenceFileSelections(config.SelectedReferenceFiles);
                    
                    // Restore selected host files  
                    RestoreHostFileSelections(config.SelectedHostFiles);
                    
                    // Restore MEP category selections (top-right panel)
                    RestoreMepCategorySelections(config.SelectedMepCategories);
                    
                    // ✅ CRITICAL FIX: Skip RestoreHostCategorySelections - UI state is restored by LoadUIStateDirectly()
                    // LoadUIStateDirectly() runs once host categories are populated (see RepopulateSectionsAfterLoad)
                    // This prevents config from overwriting UI state persistence
                    // RestoreHostCategorySelections(config.SelectedHostCategories);
                    
                    // Restore opening settings (right panel)
                    RestoreOpeningSettings(config.OpeningSettings);
                    
                    // CRITICAL: Force UI refresh to show restored selections
                    this.Refresh();
                    if (_topLeftPanel != null) _topLeftPanel.Refresh();
                    if (_topRightPanel != null) _topRightPanel.Refresh();
                    if (_bottomLeftPanel != null) _bottomLeftPanel.Refresh();
                    if (_bottomRightPanel != null) _bottomRightPanel.Refresh();
                    
                    DebugLogger.Info("=== RestoreConfigurationToUI COMPLETED ===");
                }
                else
                {
                    DebugLogger.Info("No configuration to restore to UI");
                }
                
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to restore configuration to UI: {ex.Message}");
            }
        }
        
        private void LoadConfigurationFromJsonFile(UserProfile profile)
        {
            try
            {
                var profileDir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings");
                var configFile = System.IO.Path.Combine(profileDir, $"config_{profile.Name}.txt");
                
                DebugLogger.Info($"Looking for configuration file: {configFile}");
                
                if (!System.IO.File.Exists(configFile))
                {
                    DebugLogger.Info($"Configuration file not found: {configFile}");
                    return;
                }
                
                var content = System.IO.File.ReadAllText(configFile);
                var config = new UserConfiguration();
                
                DebugLogger.Info($"Configuration file content:\n{content}");
                
                // Parse the new multi-line format
                var lines = content.Split(new string[] { Environment.NewLine, "\n", "\r" }, StringSplitOptions.None);
                string currentSection = "";
                
                DebugLogger.Info($"Parsing {lines.Length} lines from config file");
                
                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;
                    
                    // Check if this is a section header
                    if (line.Contains("=") && !line.StartsWith("  "))
                    {
                        var parts = line.Split(new char[] { '=' }, 2);
                        if (parts.Length == 2)
                        {
                            currentSection = parts[0].Trim();
                            DebugLogger.Info($"Found section header: {currentSection}");
                        }
                    }
                    else if (line.StartsWith("  ") && !string.IsNullOrEmpty(currentSection))
                    {
                        // This is a value under the current section
                        var value = line.Trim();
                        DebugLogger.Info($"Found value in {currentSection}: {value}");
                        
                        switch (currentSection)
                        {
                            case "SelectedReferenceFiles":
                                if (config.SelectedReferenceFiles == null)
                                    config.SelectedReferenceFiles = new List<string>();
                                config.SelectedReferenceFiles.Add(value);
                                break;
                            case "SelectedHostFiles":
                                if (config.SelectedHostFiles == null)
                                    config.SelectedHostFiles = new List<string>();
                                config.SelectedHostFiles.Add(value);
                                break;
                            case "SelectedMepCategories":
                                if (config.SelectedMepCategories == null)
                                    config.SelectedMepCategories = new List<string>();
                                config.SelectedMepCategories.Add(value);
                                break;
                            case "SelectedHostCategories":
                                if (config.SelectedHostCategories == null)
                                    config.SelectedHostCategories = new List<string>();
                                config.SelectedHostCategories.Add(value);
                                break;
                            case "ClearanceSettings":
                                // Parse clearance setting in format "key=value"
                                var clearanceParts = value.Split('=');
                                if (clearanceParts.Length == 2)
                                {
                                    if (config.OpeningSettings == null)
                                        config.OpeningSettings = new OpeningSettings();
                                    if (config.OpeningSettings.ClearanceSettings == null)
                                        config.OpeningSettings.ClearanceSettings = new Dictionary<string, double>();
                                    
                                    if (double.TryParse(clearanceParts[1], out double clearanceValue))
                                    {
                                        config.OpeningSettings.ClearanceSettings[clearanceParts[0]] = clearanceValue;
                                        DebugLogger.Info($"Loaded clearance setting: {clearanceParts[0]} = {clearanceValue}");
                                    }
                                }
                                break;
                            case "ClashZoneStorage":
                                // Parse clash zone storage metadata
                                if (value.StartsWith("LastUpdated="))
                                {
                                    if (config.ClashZoneStorage == null)
                                        config.ClashZoneStorage = new ClashZoneStorage();
                                    
                                    var dateStr = value.Substring("LastUpdated=".Length);
                                    if (DateTime.TryParse(dateStr, out DateTime lastUpdated))
                                    {
                                        config.ClashZoneStorage.LastUpdated = lastUpdated;
                                    }
                                }
                                else if (value.StartsWith("DocumentHash="))
                                {
                                    if (config.ClashZoneStorage == null)
                                        config.ClashZoneStorage = new ClashZoneStorage();
                                    
                                    config.ClashZoneStorage.DocumentHash = value.Substring("DocumentHash=".Length);
                                }
                                else if (value.StartsWith("ClashZonesCount="))
                                {
                                    if (config.ClashZoneStorage == null)
                                        config.ClashZoneStorage = new ClashZoneStorage();
                                    
                                    var countStr = value.Substring("ClashZonesCount=".Length);
                                    if (int.TryParse(countStr, out int count))
                                    {
                                        DebugLogger.Info($"Loaded clash zone storage with {count} zones");
                                    }
                                }
                                break;
                        }
                    }
                }
                
                // Log the loaded files
                if (config.SelectedReferenceFiles != null)
                {
                    DebugLogger.Info($"Loaded SelectedReferenceFiles: {config.SelectedReferenceFiles.Count} files");
                    foreach (var file in config.SelectedReferenceFiles)
                    {
                        DebugLogger.Info($"  Loaded reference file: {file}");
                    }
                }
                
                // Update profile with loaded configuration
                profile.Configuration = config;
                
                DebugLogger.Info($"Loaded configuration from text file: {config.SelectedReferenceFiles?.Count ?? 0} reference files, {config.SelectedMepCategories?.Count ?? 0} MEP categories");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to load configuration from text file: {ex.Message}");
            }
        }
        private void SaveUIStateDirectly()
        {
            try
            {
                if (!_hostCategoriesInitialized)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info("[UI-STATE-SAVE] Skipping save; host categories not initialized yet (preventing overwrite of vertical selections)");
                    return;
                }

                var uiState = new List<string>();
                
                // Save checked reference files
                if (_topLeftPanel?.Controls.Count > 0)
                {
                    foreach (var control in _topLeftPanel.Controls)
                    {
                        if (control is WinForms.CheckedListBox listBox)
                        {
                            for (int i = 0; i < listBox.Items.Count; i++)
                            {
                                if (listBox.GetItemChecked(i))
                                {
                                    uiState.Add($"REF:{listBox.Items[i]}");
                                }
                            }
                        }
                    }
                }
                
                // Save checked MEP categories
                if (_topRightPanel?.Controls.Count > 0)
                {
                    foreach (var control in _topRightPanel.Controls)
                    {
                        if (control is WinForms.CheckedListBox listBox)
                        {
                            for (int i = 0; i < listBox.Items.Count; i++)
                            {
                                if (listBox.GetItemChecked(i))
                                {
                                    uiState.Add($"MEP:{listBox.Items[i]}");
                                }
                            }
                        }
                    }
                }
                
                // Save checked host files
                if (_bottomLeftPanel?.Controls.Count > 0)
                {
                    foreach (var control in _bottomLeftPanel.Controls)
                    {
                        if (control is WinForms.CheckedListBox listBox)
                        {
                            for (int i = 0; i < listBox.Items.Count; i++)
                            {
                                if (listBox.GetItemChecked(i))
                                {
                                    uiState.Add($"HOST:{listBox.Items[i]}");
                                }
                            }
                        }
                    }
                }
                
                // Save checked host categories (distinguish horizontal and vertical)
                if (_horizontalCategoriesListBox != null)
                {
                    for (int i = 0; i < _horizontalCategoriesListBox.Items.Count; i++)
                    {
                        if (_horizontalCategoriesListBox.GetItemChecked(i))
                        {
                            uiState.Add($"HOSTCAT_HORIZONTAL:{_horizontalCategoriesListBox.Items[i]}");
                        }
                    }
                }
                
                if (_verticalCategoriesListBox != null)
                {
                    for (int i = 0; i < _verticalCategoriesListBox.Items.Count; i++)
                    {
                        if (_verticalCategoriesListBox.GetItemChecked(i))
                        {
                            var categoryValue = _verticalCategoriesListBox.Items[i]?.ToString() ?? string.Empty;
                            uiState.Add($"HOSTCAT_VERTICAL:{categoryValue}");
                            DebugLogger.Info($"[UI-STATE-SAVE] Saving vertical host category: {categoryValue}");
                        }
                    }
                    DebugLogger.Info($"[UI-STATE-SAVE] Vertical categories listbox has {_verticalCategoriesListBox.Items.Count} items, saved {uiState.Count(x => x.StartsWith("HOSTCAT_VERTICAL:"))} checked items");
                }
                else
                {
                    DebugLogger.Warning("[UI-STATE-SAVE] Vertical categories listbox is NULL - cannot save vertical category selections");
                }
                
                // Save clearance settings
                var clearanceSettings = GetClearanceSettings();
                foreach (var kvp in clearanceSettings)
                {
                    uiState.Add($"CLEARANCE:{kvp.Key}={kvp.Value}");
                }
                
                // Save clash zones from current profile
                var currentProfile = GetCurrentProfile();
                if (currentProfile?.Configuration?.ClashZoneStorage?.AllZones != null)
                {
                    uiState.Add($"CLASHZONES:Count={currentProfile.Configuration.ClashZoneStorage.AllZones.Count}");
                    uiState.Add($"CLASHZONES:LastUpdated={currentProfile.Configuration.ClashZoneStorage.LastUpdated:O}");
                    uiState.Add($"CLASHZONES:DocumentHash={currentProfile.Configuration.ClashZoneStorage.DocumentHash}");
                    foreach (var clashZone in currentProfile.Configuration.ClashZoneStorage.AllZones)
                    {
                        uiState.Add($"CLASHZONE:MEP={clashZone.MepElementId},Structural={clashZone.StructuralElementId},Resolved={clashZone.IsResolved}");
                    }
                }
                
                // Save to simple file
                var simpleFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
                    "JSE_MEP_Openings", "ui_state.txt");
                Directory.CreateDirectory(Path.GetDirectoryName(simpleFile) ?? "");
                File.WriteAllLines(simpleFile, uiState);
                
                DebugLogger.Info($"Saved {uiState.Count} UI state items to: {simpleFile}");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to save UI state directly: {ex.Message}");
            }
        }
        private void LoadUIStateDirectly()
        {
            try
            {
                var simpleFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
                    "JSE_MEP_Openings", "ui_state.txt");
                
                if (!File.Exists(simpleFile))
                    return;
                
                var uiState = File.ReadAllLines(simpleFile).ToList();
                DebugLogger.Info($"Loading {uiState.Count} UI state items from: {simpleFile}");
                
                // Restore reference files
                foreach (var item in uiState.Where(x => x.StartsWith("REF:")))
                {
                    var value = item.Substring(4);
                    RestoreCheckedItem(_topLeftPanel, value);
                }
                
                // Restore MEP categories
                foreach (var item in uiState.Where(x => x.StartsWith("MEP:")))
                {
                    var value = item.Substring(4);
                    RestoreCheckedItem(_topRightPanel, value);
                }
                
                // Restore host files
                foreach (var item in uiState.Where(x => x.StartsWith("HOST:")))
                {
                    var value = item.Substring(5);
                    RestoreCheckedItem(_bottomLeftPanel, value);
                }
                
                // Restore host categories (distinguish horizontal and vertical)
                foreach (var item in uiState.Where(x => x.StartsWith("HOSTCAT_HORIZONTAL:")))
                {
                    var value = item.Substring(20); // Remove "HOSTCAT_HORIZONTAL:" prefix
                    if (_horizontalCategoriesListBox != null)
                    {
                        for (int i = 0; i < _horizontalCategoriesListBox.Items.Count; i++)
                        {
                            if (_horizontalCategoriesListBox.Items[i]?.ToString() == value)
                            {
                                _horizontalCategoriesListBox.SetItemChecked(i, true);
                                DebugLogger.Info($"Restored horizontal host category selection: {value}");
                                break;
                            }
                        }
                    }
                }
                
                foreach (var item in uiState.Where(x => x.StartsWith("HOSTCAT_VERTICAL:")))
                {
                    var value = item.Substring(18); // Remove "HOSTCAT_VERTICAL:" prefix
                    if (_verticalCategoriesListBox != null)
                    {
                        bool found = false;
                        for (int i = 0; i < _verticalCategoriesListBox.Items.Count; i++)
                        {
                            var itemValue = _verticalCategoriesListBox.Items[i]?.ToString() ?? string.Empty;
                            if (itemValue == value)
                            {
                                _verticalCategoriesListBox.SetItemChecked(i, true);
                                DebugLogger.Info($"[UI-STATE-LOAD] ✅ Restored vertical host category selection: {value}");
                                found = true;
                                break;
                            }
                        }
                        if (!found)
                        {
                            DebugLogger.Warning($"[UI-STATE-LOAD] ⚠️ Vertical host category '{value}' not found in listbox (listbox has {_verticalCategoriesListBox.Items.Count} items)");
                        }
                    }
                    else
                    {
                        DebugLogger.Warning($"[UI-STATE-LOAD] ⚠️ Vertical categories listbox is NULL - cannot restore vertical category '{value}'");
                    }
                }
                
                // ✅ BACKWARD COMPATIBILITY: Also handle old "HOSTCAT:" format (restore to horizontal for compatibility)
                foreach (var item in uiState.Where(x => x.StartsWith("HOSTCAT:") && !x.StartsWith("HOSTCAT_HORIZONTAL:") && !x.StartsWith("HOSTCAT_VERTICAL:")))
                {
                    var value = item.Substring(8); // Remove "HOSTCAT:" prefix
                    if (_horizontalCategoriesListBox != null)
                    {
                        for (int i = 0; i < _horizontalCategoriesListBox.Items.Count; i++)
                        {
                            if (_horizontalCategoriesListBox.Items[i]?.ToString() == value)
                            {
                                _horizontalCategoriesListBox.SetItemChecked(i, true);
                                DebugLogger.Info($"Restored host category selection (legacy format): {value}");
                                break;
                            }
                        }
                    }
                }
                
                // Restore clearance settings
                foreach (var item in uiState.Where(x => x.StartsWith("CLEARANCE:")))
                {
                    var clearanceData = item.Substring(10); // Remove "CLEARANCE:" prefix
                    var parts = clearanceData.Split('=');
                    if (parts.Length == 2)
                    {
                        var key = parts[0];
                        if (double.TryParse(parts[1], out double value))
                        {
                            RestoreClearanceSetting(key, value);
                        }
                    }
                }
                
                // Restore clash zones
                RestoreClashZonesFromUIState(uiState);
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to load UI state directly: {ex.Message}");
            }
        }
        
        private void RestoreClashZonesFromUIState(List<string> uiState)
        {
            try
            {
                // 🔍 DIAGNOSTIC: Count sleeves BEFORE RestoreClashZonesFromUIState
                if (_document != null)
                {
                    var sleevesBeforeRestore = new FilteredElementCollector(_document)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                        .Count();
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 RestoreClashZonesFromUIState: Sleeves BEFORE: {sleevesBeforeRestore}\n");
                }
                
                DebugLogger.Info("[CLASH_RESTORE] Restoring clash zones from UI state");
                
                var currentProfile = GetCurrentProfile();
                if (currentProfile?.Configuration == null)
                {
                    DebugLogger.Info("[CLASH_RESTORE] No current profile configuration - skipping clash zone restoration");
                    return;
                }
                
                // Find clash zone metadata
                var countItem = uiState.FirstOrDefault(x => x.StartsWith("CLASHZONES:Count="));
                var lastUpdatedItem = uiState.FirstOrDefault(x => x.StartsWith("CLASHZONES:LastUpdated="));
                var documentHashItem = uiState.FirstOrDefault(x => x.StartsWith("CLASHZONES:DocumentHash="));
                
                if (countItem == null)
                {
                    DebugLogger.Info("[CLASH_RESTORE] No clash zones found in UI state");
                    return;
                }
                
                // Parse count
                var countStr = countItem.Substring("CLASHZONES:Count=".Length);
                if (!int.TryParse(countStr, out int count) || count == 0)
                {
                    DebugLogger.Info("[CLASH_RESTORE] No clash zones to restore");
                    return;
                }
                
                // Create clash zone storage
                var clashZoneStorage = new ClashZoneStorage
                {
                    ClashZones = new List<ClashZone>(),
                    LastUpdated = DateTime.Now,
                    DocumentHash = "Unknown"
                };
                
                // Parse metadata
                if (lastUpdatedItem != null)
                {
                    var dateStr = lastUpdatedItem.Substring("CLASHZONES:LastUpdated=".Length);
                    if (DateTime.TryParse(dateStr, out DateTime lastUpdated))
                    {
                        clashZoneStorage.LastUpdated = lastUpdated;
                    }
                }
                
                if (documentHashItem != null)
                {
                    clashZoneStorage.DocumentHash = documentHashItem.Substring("CLASHZONES:DocumentHash=".Length);
                }
                
                // Parse individual clash zones
                foreach (var item in uiState.Where(x => x.StartsWith("CLASHZONE:")))
                {
                    var clashData = item.Substring("CLASHZONE:".Length);
                    var parts = clashData.Split(',');
                    
                    if (parts.Length >= 3)
                    {
                        var mepPart = parts[0].Split('=');
                        var structuralPart = parts[1].Split('=');
                        var resolvedPart = parts[2].Split('=');
                        
                        if (mepPart.Length == 2 && structuralPart.Length == 2 && resolvedPart.Length == 2)
                        {
                            if (int.TryParse(mepPart[1], out int mepId) && 
                                int.TryParse(structuralPart[1], out int structuralId) &&
                                bool.TryParse(resolvedPart[1], out bool isResolved))
                            {
                                var clashZone = new ClashZone
                                {
                                    MepElementId = new Autodesk.Revit.DB.ElementId(mepId),
                                    StructuralElementId = new Autodesk.Revit.DB.ElementId(structuralId),
                                    IsResolved = isResolved,
                                    DetectedAt = clashZoneStorage.LastUpdated
                                };
                                clashZoneStorage.ClashZones.Add(clashZone);
                            }
                        }
                    }
                }
                
                // Save to profile configuration
                currentProfile.Configuration.ClashZoneStorage = clashZoneStorage;
                DebugLogger.Info($"[CLASH_RESTORE] Restored {clashZoneStorage.ClashZones.Count} clash zones to profile configuration");
                
                // 🔍 DIAGNOSTIC: Count sleeves AFTER RestoreClashZonesFromUIState
                if (_document != null)
                {
                    var sleevesAfterRestore = new FilteredElementCollector(_document)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                        .Count();
                    DebugLogger.Info($"[{DateTime.Now}] 🔍 RestoreClashZonesFromUIState: Sleeves AFTER: {sleevesAfterRestore}\n");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[CLASH_RESTORE] Failed to restore clash zones: {ex.Message}");
                
                // 🔍 DIAGNOSTIC: Count sleeves AFTER RestoreClashZonesFromUIState ERROR
                try
                {
                    if (_document != null)
                    {
                        var sleevesAfterError = new FilteredElementCollector(_document)
                            .OfClass(typeof(FamilyInstance))
                            .Cast<FamilyInstance>()
                            .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                            .Count();
                        DebugLogger.Info($"[{DateTime.Now}] 🔍 RestoreClashZonesFromUIState: Sleeves AFTER ERROR: {sleevesAfterError}\n");
                    }
                }
                catch { }
            }
        }
        
        private void RestoreClearanceSetting(string key, double value)
        {
            try
            {
                DebugLogger.Info($"[CLEARANCE_RESTORE] Restoring clearance setting: {key} = {value}mm");
                
                // Find the appropriate text box and set its value
                if (_clearancePanel?.Controls.Count > 0)
                {
                    foreach (var control in _clearancePanel.Controls)
                    {
                        if (control is WinForms.TextBox textBox && textBox.Tag != null)
                        {
                            string genericKey = textBox.Tag.ToString() ?? "";
                            
                            // Check if this clearance setting matches our generic key
                            if (key.EndsWith($"_{genericKey}") || key == genericKey)
                            {
                                textBox.Text = value.ToString();
                                DebugLogger.Info($"[CLEARANCE_RESTORE] Restored {genericKey} -> {key} = {value}mm");
                                return;
                            }
                        }
                    }
                }
                
                DebugLogger.Info($"[CLEARANCE_RESTORE] No matching text box found for {key}");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[CLEARANCE_RESTORE] Failed to restore clearance setting {key}: {ex.Message}");
            }
        }
        /// <summary>
        /// Restores MEP categories from a filter to the UI controls
        /// </summary>
        private void RestoreMepCategoriesFromFilter(List<string> selectedCategories)
        {
            if (selectedCategories == null || selectedCategories.Count == 0) return;
            
            try
            {
                DebugLogger.Info($"[FILTER_UI] Restoring MEP categories to UI: {string.Join(", ", selectedCategories)}");
                
                if (_topRightPanel?.Controls.Count > 0)
                {
                    foreach (var control in _topRightPanel.Controls)
                    {
                        if (control is WinForms.CheckedListBox listBox)
                        {
                            listBox.BeginUpdate();
                            
                            // First, uncheck all items
                            for (int i = 0; i < listBox.Items.Count; i++)
                            {
                                listBox.SetItemChecked(i, false);
                            }
                            
                            // Then, check the selected categories
                            foreach (var category in selectedCategories)
                            {
                                for (int i = 0; i < listBox.Items.Count; i++)
                                {
                                    if (listBox.Items[i]?.ToString() == category)
                                    {
                                        listBox.SetItemChecked(i, true);
                                        DebugLogger.Info($"[FILTER_UI] Checked MEP category: {category}");
                                        break;
                                    }
                                }
                            }
                            
                            listBox.EndUpdate();
                        }
                    }
                }
                
                DebugLogger.Info($"[FILTER_UI] MEP categories restoration completed");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[FILTER_UI] Failed to restore MEP categories: {ex.Message}");
            }
        }
        
        private void RestoreCheckedItem(WinForms.Panel panel, string value)
        {
            if (panel?.Controls.Count > 0)
            {
                foreach (var control in panel.Controls)
                {
                    if (control is WinForms.CheckedListBox listBox)
                    {
                        for (int i = 0; i < listBox.Items.Count; i++)
                        {
                            if (listBox.Items[i]?.ToString() == value)
                            {
                                listBox.SetItemChecked(i, true);
                                DebugLogger.Info($"Restored checked item: {value}");
                                break;
                            }
                        }
                    }
                }
            }
        }
        
        private void ParseOldTextFormat(string content, UserConfiguration config)
        {
            var lines = content.Split(new string[] { Environment.NewLine, "\n", "\r" }, StringSplitOptions.None);
                
                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line) || !line.Contains("="))
                        continue;
                    
                    var parts = line.Split(new char[] { '=' }, 2);
                    if (parts.Length != 2)
                        continue;
                    
                    var key = parts[0].Trim();
                    var value = parts[1].Trim();
                    
                    switch (key)
                    {
                        case "SelectedReferenceFiles":
                            config.SelectedReferenceFiles = string.IsNullOrEmpty(value) ? new List<string>() : value.Split('|').ToList();
                            break;
                        case "SelectedHostFiles":
                            config.SelectedHostFiles = string.IsNullOrEmpty(value) ? new List<string>() : value.Split('|').ToList();
                            break;
                        case "SelectedMepCategories":
                            config.SelectedMepCategories = string.IsNullOrEmpty(value) ? new List<string>() : value.Split('|').ToList();
                            break;
                        case "SelectedHostCategories":
                            config.SelectedHostCategories = string.IsNullOrEmpty(value) ? new List<string>() : value.Split('|').ToList();
                            break;
                    }
            }
        }
        
        private List<string> GetSelectedReferenceFiles()
        {
            var selectedFiles = new List<string>();
            
            // Get selected files from top-left panel (Reference Elements)
            if (_topLeftPanel?.Controls.Count > 0)
            {
                foreach (var control in _topLeftPanel.Controls)
                {
                    if (control is WinForms.CheckedListBox listBox)
                    {
                        for (int i = 0; i < listBox.Items.Count; i++)
                        {
                            if (listBox.GetItemChecked(i))
                            {
                                // ✅ FIX: Extract just filename - remove any extra info like "[NOT LOADED]", .rvt extension, or old format with element count
                                string itemText = listBox.Items[i].ToString() ?? "";
                                // Remove "[NOT LOADED]" suffix if present
                                itemText = itemText.Replace(" [NOT LOADED]", "");
                                // Remove any old format with element count: "filename (X elements)" or ": location Shared"
                                int parenIndex = itemText.IndexOf('(');
                                if (parenIndex > 0) itemText = itemText.Substring(0, parenIndex).Trim();
                                int colonIndex = itemText.IndexOf(':');
                                if (colonIndex > 0) itemText = itemText.Substring(0, colonIndex).Trim();
                                // Remove .rvt extension if present
                                itemText = System.IO.Path.GetFileNameWithoutExtension(itemText);
                                selectedFiles.Add(itemText);
                            }
                        }
                    }
                }
            }
            
            return selectedFiles;
        }
        
        private List<string> GetSelectedHostFiles()
        {
            var selectedFiles = new List<string>();
            
            // Get selected files from bottom-left panel (Host Elements)
            if (_bottomLeftPanel?.Controls.Count > 0)
            {
                foreach (var control in _bottomLeftPanel.Controls)
                {
                    if (control is WinForms.CheckedListBox listBox)
                    {
                        for (int i = 0; i < listBox.Items.Count; i++)
                        {
                            if (listBox.GetItemChecked(i))
                            {
                                // ✅ FIX: Extract just filename - remove any extra info like "[NOT LOADED]", .rvt extension, or old format with element count
                                string itemText = listBox.Items[i].ToString() ?? "";
                                // Remove "[NOT LOADED]" suffix if present
                                itemText = itemText.Replace(" [NOT LOADED]", "");
                                // Remove any old format with element count: "filename (X elements)" or ": location Shared"
                                int parenIndex = itemText.IndexOf('(');
                                if (parenIndex > 0) itemText = itemText.Substring(0, parenIndex).Trim();
                                int colonIndex = itemText.IndexOf(':');
                                if (colonIndex > 0) itemText = itemText.Substring(0, colonIndex).Trim();
                                // Remove .rvt extension if present
                                itemText = System.IO.Path.GetFileNameWithoutExtension(itemText);
                                selectedFiles.Add(itemText);
                            }
                        }
                    }
                }
            }
            
            return selectedFiles;
        }
        
        private List<string> GetSelectedMepCategories()
        {
            try
        {
            var selectedCategories = new List<string>();
            
            // Get selected categories from top-right panel (MEP Categories)
            if (_topRightPanel?.Controls.Count > 0)
            {
                foreach (var control in _topRightPanel.Controls)
                {
                    if (control is WinForms.CheckedListBox listBox)
                    {
                        for (int i = 0; i < listBox.Items.Count; i++)
                        {
                            if (listBox.GetItemChecked(i))
                            {
                                selectedCategories.Add(listBox.Items[i].ToString() ?? "");
                            }
                        }
                    }
                }
            }
            
                DebugLogger.Info($"[FILTER_UI] GetSelectedMepCategories found {selectedCategories.Count} categories: {string.Join(", ", selectedCategories)}");
            return selectedCategories;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[FILTER_UI] Error getting selected MEP categories: {ex.Message}");
                return new List<string>();
            }
        }
        
        private List<string> GetSelectedHostCategories()
        {
            var selectedCategories = new List<string>();
            
            // Get selected categories from bottom-right panel (Host Categories)
            if (_bottomRightPanel?.Controls.Count > 0)
            {
                foreach (var control in _bottomRightPanel.Controls)
                {
                    if (control is WinForms.CheckedListBox listBox)
                    {
                        for (int i = 0; i < listBox.Items.Count; i++)
                        {
                            if (listBox.GetItemChecked(i))
                            {
                                selectedCategories.Add(listBox.Items[i].ToString() ?? "");
                            }
                        }
                    }
                }
            }
            
            return selectedCategories;
        }

        private List<ElementId> GetSelectedOpenings()
        {
            var openingIds = new List<ElementId>();
            
            try
            {
                if (_document == null) return openingIds;
                
                var doc = _document;
                
                // Get opening categories
                var openingCategories = new List<BuiltInCategory>
                {
                    BuiltInCategory.OST_GenericModel, // Generic openings
                    BuiltInCategory.OST_GenericAnnotation, // Alternative for openings
                    BuiltInCategory.OST_StructuralFraming // Alternative for structural openings
                };
                
                var filter = new ElementMulticategoryFilter(openingCategories);
                var collector = new FilteredElementCollector(doc)
                    .WherePasses(filter)
                    .WhereElementIsNotElementType();
                
                foreach (Element element in collector)
                {
                    if (IsOpeningElement(element))
                    {
                        openingIds.Add(element.Id);
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error getting selected openings: {ex.Message}");
            }
            
            return openingIds;
        }

        private bool IsOpeningElement(Element element)
        {
            try
            {
                // Check if element is an opening based on category and family name
                var category = element.Category?.Name;
                var familyName = element.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)?.AsString();
                
                // Check for opening-related categories
                if (category != null && (
                    category.Contains("Opening") ||
                    category.Contains("Generic Model")))
                {
                    return true;
                }
                
                // Check for opening-related family names
                if (familyName != null && (
                    familyName.Contains("Opening") ||
                    familyName.Contains("Sleeve") ||
                    familyName.Contains("Penetration")))
                {
                    return true;
                }
                
                return false;
            }
            catch
            {
                return false;
            }
        }
        
        private OpeningSettings GetCurrentOpeningSettings()
        {
            var settings = new OpeningSettings();
            
            try
            {
                // Get MEP type selection
                if (_mepTypeCombo?.SelectedItem != null)
                {
                    settings.SelectedMepType = _mepTypeCombo.SelectedItem.ToString() ?? "Pipe";
                }
                
                // Get clearance settings
                settings.ClearanceSettings = GetClearanceSettings();
                
                DebugLogger.Info($"Collected opening settings: MEP Type={settings.SelectedMepType}");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to collect opening settings: {ex.Message}");
            }
            
            return settings;
        }
        private Dictionary<string, double> GetClearanceSettings(string? specificCategory = null)
        {
            var clearances = new Dictionary<string, double>();
            
            try
            {
                DebugLogger.Info($"[CLEARANCE_DEBUG] === GetClearanceSettings START ===");
                
                // If specific category is provided, only collect clearances for that category
                // Otherwise, collect for all selected categories (for backward compatibility)
                List<string> targetCategories;
                if (!string.IsNullOrEmpty(specificCategory))
                {
                    targetCategories = new List<string> { specificCategory };
                    DebugLogger.Info($"[CLEARANCE_DEBUG] Collecting clearances for specific category: '{specificCategory}'");
                }
                else
                {
                    targetCategories = GetSelectedMepCategories();
                    DebugLogger.Info($"[CLEARANCE_DEBUG] Collecting clearances for selected categories: {string.Join(", ", targetCategories)}");
                }
                
                if (targetCategories.Count == 0)
                {
                    DebugLogger.Warning($"[CLEARANCE_DEBUG] No target categories - returning empty clearances");
                    return clearances;
                }
                
                DebugLogger.Info($"[CLEARANCE_DEBUG] _clearancePanel exists: {_clearancePanel != null}");
                DebugLogger.Info($"[CLEARANCE_DEBUG] _clearancePanel.Controls.Count: {_clearancePanel?.Controls.Count ?? 0}");
                
                // Get clearance values from clearance panel - only for target categories
                if (_clearancePanel?.Controls.Count > 0)
                {
                    foreach (var control in _clearancePanel.Controls)
                    {
                        if (control is WinForms.TextBox textBox && textBox.Tag != null)
                        {
                            DebugLogger.Info($"[CLEARANCE_DEBUG] Found TextBox: Tag='{textBox.Tag}', Text='{textBox.Text}', Visible={textBox.Visible}");
                            
                            if (double.TryParse(textBox.Text, out double value))
                            {
                                string genericKey = textBox.Tag.ToString() ?? "";
                                
                                // Generate clearance keys for target categories only
                                foreach (var category in targetCategories)
                                {
                                    string specificKey = ConvertToSpecificClearanceKey(genericKey, category);
                                clearances[specificKey] = value;
                                
                                    DebugLogger.Info($"[CLEARANCE_DEBUG] Clearance setting: {genericKey} -> {specificKey} = {value}mm (for category: {category})");
                                }
                            }
                            else
                            {
                                DebugLogger.Warning($"[CLEARANCE_DEBUG] Failed to parse TextBox value: '{textBox.Text}' for Tag: '{textBox.Tag}'");
                            }
                        }
                    }
                }
                else
                {
                    DebugLogger.Warning($"[CLEARANCE_DEBUG] _clearancePanel is null or has no controls!");
                }
                
                // Get cable tray specific clearance values - only if Cable Trays is in target categories
                if (targetCategories.Contains("Cable Trays") && _cableTrayPanel?.Controls.Count > 0)
                {
                    DebugLogger.Info($"[CLEARANCE_DEBUG] Collecting cable tray clearances (Cable Trays is in target categories)");
                    foreach (var control in _cableTrayPanel.Controls)
                    {
                        if (control is WinForms.TextBox textBox && textBox.Tag != null)
                        {
                            if (double.TryParse(textBox.Text, out double value))
                            {
                                string key = textBox.Tag.ToString() ?? "";
                                clearances[key] = value;
                                
                                DebugLogger.Info($"[CLEARANCE_DEBUG] Cable tray clearance setting: {key} = {value}mm");
                            }
                        }
                    }
                }
                else
                {
                    DebugLogger.Info($"[CLEARANCE_DEBUG] Skipping cable tray clearances - Cable Trays not in target categories or panel empty");
                }

                // Get damper specific clearance values - only if Duct Accessories is in target categories
                if (targetCategories.Contains("Duct Accessories") && _damperPanel?.Controls.Count > 0)
                {
                    DebugLogger.Info($"[CLEARANCE_DEBUG] Collecting damper clearances (Duct Accessories is in target categories)");
                    foreach (var control in _damperPanel.Controls)
                    {
                        if (control is WinForms.TextBox textBox && textBox.Tag != null)
                        {
                            if (double.TryParse(textBox.Text, out double value))
                            {
                                string key = textBox.Tag.ToString() ?? "";
                                clearances[key] = value;
                                
                                DebugLogger.Info($"[CLEARANCE_DEBUG] Damper clearance setting: {key} = {value}mm");
                            }
                        }
                    }
                }
                else
                {
                    DebugLogger.Info($"[CLEARANCE_DEBUG] Skipping damper clearances - Duct Accessories not in target categories or panel empty");
                }

                // Get pipe specific clearance values - only if Pipes is in target categories
                if (targetCategories.Contains("Pipes") && _pipePanel?.Controls.Count > 0)
                {
                    DebugLogger.Info($"[CLEARANCE_DEBUG] Collecting pipe clearances (Pipes is in target categories)");
                    foreach (var control in _pipePanel.Controls)
                    {
                        if (control is WinForms.TextBox textBox && textBox.Tag != null)
                        {
                            if (double.TryParse(textBox.Text, out double value))
                            {
                                string key = textBox.Tag.ToString() ?? "";
                                clearances[key] = value;
                                
                                DebugLogger.Info($"[CLEARANCE_DEBUG] Pipe clearance setting: {key} = {value}mm");
                            }
                        }
                    }
                }
                else
                {
                    DebugLogger.Info($"[CLEARANCE_DEBUG] Skipping pipe clearances - Pipes not in target categories or panel empty");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to collect clearance settings: {ex.Message}");
            }
            
            // DEBUG: Log final clearance dictionary
            DebugLogger.Info($"[CLEARANCE_DEBUG] Final clearance dictionary ({clearances.Count} entries):");
            foreach (var kvp in clearances)
            {
                DebugLogger.Info($"  {kvp.Key} = {kvp.Value}mm");
            }
            
            return clearances;
        }
        /// <summary>
        /// Convert generic clearance key to category-specific key
        /// </summary>
        private string ConvertToSpecificClearanceKey(string genericKey, string category)
        {
            // Handle special cases first
            if (category.Equals("Duct Accessories", StringComparison.OrdinalIgnoreCase))
            {
                // Duct accessories have specific keys for MEP connector side and other sides
                return genericKey switch
                {
                    "normal_clearance" => "ductaccessories_other_normal",
                    "insulated_clearance" => "ductaccessories_other_insulated",
                    "ductaccessories_mep_normal" => "ductaccessories_mep_normal",
                    "ductaccessories_mep_insulated" => "ductaccessories_mep_insulated",
                    "ductaccessories_other_normal" => "ductaccessories_other_normal",
                    "ductaccessories_other_insulated" => "ductaccessories_other_insulated",
                    _ => genericKey
                };
            }
            
            if (category.Equals("Cable Trays", StringComparison.OrdinalIgnoreCase))
            {
                // Cable trays use specific keys for top and other sides
                return genericKey switch
                {
                    "normal_clearance" => "cabletray_top_normal", // Default to top clearance
                    "insulated_clearance" => "cabletray_top_insulated", // Default to top clearance
                    _ => genericKey
                };
            }
            
            if (category.Equals("Pipes", StringComparison.OrdinalIgnoreCase))
            {
                // Pipes use specific keys for normal and insulated clearances
                return genericKey switch
                {
                    "normal_clearance" => "pipes_normal_clearance",
                    "insulated_clearance" => "pipes_insulated_clearance",
                    "pipes_normal_clearance" => "pipes_normal_clearance",
                    "pipes_insulated_clearance" => "pipes_insulated_clearance",
                    _ => genericKey
                };
            }
            
            // Standard MEP categories (Ducts)
            string categoryKey = category.ToLower().Replace(" ", "_");
            
            return genericKey switch
            {
                "normal_clearance" => $"{categoryKey}_normal_clearance",
                "insulated_clearance" => $"{categoryKey}_insulated_clearance",
                _ => genericKey
            };
        }

        /// <summary>
        /// Get current MEP category from UI selection
        /// </summary>
        private string GetCurrentMepCategory()
        {
            // Get category from MEP type combo selection
            string selectedMepType = _mepTypeCombo?.SelectedItem?.ToString() ?? string.Empty;
            
            // Map MEP type to category
            return selectedMepType switch
            {
                "Ducts" => "Ducts",
                "Duct Accessories" => "Duct Accessories", 
                "Pipes" => "Pipes",
                "Cable Trays" => "Cable Trays",
                "Cable Tray" => "Cable Trays",
                "Conduit" => "Conduits",
                _ => "Default"
            };
        }


        /// <summary>
        /// Get selected filter items from the filters panel (left column)
        /// </summary>
        private List<string> GetSelectedFilterItems()
        {
            var selectedItems = new List<string>();

            try
            {
                // Get selected items from the filters panel (left column)
                if (_filtersPanel?.Controls.Count > 0)
                {
                    foreach (var control in _filtersPanel.Controls)
                    {
                        if (control is WinForms.ListBox listBox)
                        {
                            foreach (var selectedItem in listBox.SelectedItems)
                            {
                                if (selectedItem != null)
                                {
                                    selectedItems.Add(selectedItem.ToString() ?? "");
                                }
                            }
                        }
                    }
                }

                DebugLogger.Info($"GetSelectedFilterItems: Found {selectedItems.Count} selected filter items: {string.Join(", ", selectedItems)}");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to get selected filter items: {ex.Message}");
            }

            return selectedItems;
        }

        /// <summary>
        /// Get selected filters from the filters panel
        /// </summary>
        public List<OpeningFilter> GetSelectedFilters()
        {
            var selectedFilters = new List<OpeningFilter>();
            
            try
            {
                // Get the actual selected filter names from the UI
                var selectedFilterNames = GetSelectedFilterItems();
                DebugLogger.Info($"GetSelectedFilters: UI selected filter names: {string.Join(", ", selectedFilterNames)}");
                
                // Get selected MEP categories from the UI
                var selectedMepCategories = GetSelectedMepCategories();
                DebugLogger.Info($"GetSelectedFilters: UI selected MEP categories: {string.Join(", ", selectedMepCategories)}");
                
                // Create filters using the actual user-selected filter names
                foreach (var filterName in selectedFilterNames)
                {
                foreach (var categoryName in selectedMepCategories)
                {
                    Models.MepCategory category;
                    
                        // Map UI category names to enum
                    switch (categoryName.ToLower())
                    {
                        case "ducts":
                            category = Models.MepCategory.Ducts;
                            break;
                        case "ductaccessories":
                        case "duct accessories":
                            category = Models.MepCategory.DuctAccessories;
                            break;
                        case "pipes":
                            category = Models.MepCategory.Pipes;
                            break;
                        case "cabletrays":
                        case "cable trays":
                            category = Models.MepCategory.CableTrays;
                            break;
                        default:
                            DebugLogger.Warning($"GetSelectedFilters: Unknown category '{categoryName}' - skipping");
                            continue;
                    }
                    
                        // Use the actual user-selected filter name as the discipline name
                        var filter = OpeningFilter.CreateDefault(category, filterName);
                    selectedFilters.Add(filter);
                        DebugLogger.Info($"GetSelectedFilters: Created filter for {categoryName} -> {filterName}");
                    }
                }

                DebugLogger.Info($"GetSelectedFilters: Returning {selectedFilters.Count} filters based on UI selections");
                foreach (var filter in selectedFilters)
                {
                    DebugLogger.Info($"  - {filter.GetDescription()} (Category: {filter.Category}, Name: {filter.Name})");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to get selected filters: {ex.Message}");
            }
            
            return selectedFilters;
        }

        private void OnCancelClick(object? sender, EventArgs e)
        {
            _statusLabel.Text = "Cancel clicked";
        }

        private void OnCloseClick(object? sender, EventArgs e)
        {
            this.Close();
        }

        private void OnParameterTransferClick(object? sender, EventArgs e)
        {
            try
            {
                _statusLabel.Text = "Opening Parameter Transfer dialog...";
                
                // Open the Parameter Transfer dialog
                using (var parameterTransferDialog = new ParameterTransferDialog(_document))
                {
                    if (parameterTransferDialog.ShowDialog() == WinForms.DialogResult.OK)
                    {
                        var configuration = parameterTransferDialog.GetConfiguration();
                        
                        // ENHANCEMENT: Save the parameter transfer configuration for future use
                        try
                        {
                            var parameterTransferService = new ParameterTransferService();
                            if (parameterTransferService.SaveCurrentParameterTransferConfiguration(configuration))
                            {
                                DebugLogger.Info($"[PARAMETER_TRANSFER] Saved current parameter transfer configuration with {configuration.Mappings.Count} mappings");
                            }
                        }
                        catch (Exception saveEx)
                        {
                            DebugLogger.Warning($"[PARAMETER_TRANSFER] Failed to save parameter transfer configuration: {saveEx.Message}");
                        }
                        
                        // Get selected openings (you can implement this based on your selection logic)
                        var selectedOpeningIds = GetSelectedOpenings();
                        
                        if (selectedOpeningIds.Count == 0)
                        {
                            WinForms.MessageBox.Show("No openings found in the project.", "No Openings", 
                                WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                            _statusLabel.Text = "No openings found for parameter transfer.";
                            return;
                        }
                        
                        // Execute parameter transfer
                        var transferService = new ParameterTransferService();
                        var result = transferService.ExecuteTransferConfiguration(_document, selectedOpeningIds, configuration);
                        
                        // Show result
                        if (result.Success)
                        {
                            var messageText = $"Parameter transfer completed successfully!\n\n" +
                                            $"Transferred: {result.TransferredCount}\n" +
                                            $"Failed: {result.FailedCount}";
                            
                            if (result.Warnings.Count > 0)
                            {
                                messageText += $"\nWarnings: {result.Warnings.Count}";
                            }
                            
                            WinForms.MessageBox.Show(messageText, "Success", 
                                WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                            _statusLabel.Text = $"Parameter transfer completed: {result.TransferredCount} successful, {result.FailedCount} failed.";
                        }
                        else
                        {
                            WinForms.MessageBox.Show($"Parameter transfer failed: {result.Message}", "Error", 
                                WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
                            _statusLabel.Text = "Parameter transfer failed.";
                        }
                    }
                    else
                    {
                        _statusLabel.Text = "Parameter transfer cancelled.";
                    }
                }
            }
            catch (Exception ex)
            {
                WinForms.MessageBox.Show($"Error opening parameter transfer dialog: {ex.Message}", "Error", 
                    WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
                _statusLabel.Text = "Error opening parameter transfer dialog.";
                DebugLogger.Error($"Error in OnParameterTransferClick: {ex.Message}");
            }
        }

        /// <summary>
        /// Launches the Parameter Service dialog with integration to update main UI dropdowns
        /// </summary>
        public void LaunchParameterService()
        {
            try
            {
                DebugLogger.Info("=== LAUNCHING PARAMETER SERVICE V2 WITH MAIN UI INTEGRATION ===");
                
                // ✅ V2: Use ParameterServiceDialogV2 (old version moved to Backup folder)
                using (var parameterServiceDialog = new ParameterServiceDialogV2(_document, _uiDocument))
                {
                    var result = parameterServiceDialog.ShowDialog();
                    
                    if (result == WinForms.DialogResult.OK)
                    {
                        DebugLogger.Info("Parameter service completed successfully.");
                        _statusLabel.Text = "Parameter service completed!";
                    }
                    else
                    {
                        DebugLogger.Info("Parameter service cancelled");
                        _statusLabel.Text = "Parameter service cancelled";
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error launching parameter service: {ex.Message}");
                WinForms.MessageBox.Show($"Error launching parameter service: {ex.Message}", "Error", 
                    WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
                _statusLabel.Text = "Error launching parameter service";
            }
        }

        private void OnRefreshClick(object? sender, EventArgs e)
        {
            // IMMEDIATE LOGGING BEFORE ANYTHING ELSE
            System.Diagnostics.Debug.WriteLine($"[ON_REFRESH_CLICK] === REFRESH BUTTON CLICKED AT {DateTime.Now} ===");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("logger_debug.txt"), $"[{DateTime.Now}] === ON_REFRESH_CLICK STARTED ===\n");

            // ✅ TIMING: Start timing measurement
            var refreshStopwatch = System.Diagnostics.Stopwatch.StartNew();
            var startTime = DateTime.Now;
            
            try
            {
                
                DebugLogger.Info("=== REFRESH BUTTON CLICKED ===");
                System.Diagnostics.Debug.WriteLine($"[ON_REFRESH_CLICK] === START TIME: {startTime:HH:mm:ss.fff} ===");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("refresh_timing.log"), 
                    $"[{startTime:HH:mm:ss.fff}] === REFRESH CLICK STARTED ===\n");

                // Use RefreshService instead of inline method
                var document = GetCurrentDocument();
                if (document != null)
                {
                    // ✅ CRITICAL FIX: Capture Section Box BEFORE any other logic
                    // This ensures the database has the FRESH section box bounds for the current view
                    if (document.ActiveView is View3D view3D && view3D.IsSectionBoxActive)
                    {
                        try
                        {
                            using (var dbContext = new JSE_RevitAddin_MEP_OPENINGS.Data.SleeveDbContext(document))
                            {
                                var sectionBoxService = new JSE_RevitAddin_MEP_OPENINGS.Services.SectionBoxService();
                                sectionBoxService.CaptureAndStore(view3D, dbContext.Connection);
                                DebugLogger.Info($"[SECTION-BOX] PRE-REFRESH: Updated DB Section Box from View: {view3D.Name}");
                                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("logger_debug.txt"), $"[{DateTime.Now}] [SECTION-BOX] PRE-REFRESH: Updated DB Section Box from View: {view3D.Name}\n");
                            }
                        }
                        catch (Exception sbEx)
                        {
                             DebugLogger.Warning($"[SECTION-BOX] PRE-REFRESH: Failed to update section box: {sbEx.Message}");
                             JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("logger_debug.txt"), $"[{DateTime.Now}] [SECTION-BOX] PRE-REFRESH: Failed to update section box: {sbEx.Message}\n");
                        }
                    }
                    // Get actual UI selections
                    var selectedFilterItems = GetSelectedFilterItems();
                    var selectedMepCategories = GetSelectedMepCategories();
                    var selectedReferenceFiles = GetSelectedReferenceFiles();
                    var selectedHostFiles = GetSelectedHostFiles();
                    var clearanceSettings = GetClearanceSettings();
                    
                    // ✅ CRITICAL DIAGNOSTIC: Log UI selections with full details including linked file names
                    DebugLogger.Info($"[ON_REFRESH_CLICK] UI Selections - Filters: {selectedFilterItems?.Count ?? 0} [{string.Join(", ", selectedFilterItems ?? new List<string>())}]");
                    DebugLogger.Info($"[ON_REFRESH_CLICK] UI Selections - MEP Categories: {selectedMepCategories?.Count ?? 0} [{string.Join(", ", selectedMepCategories ?? new List<string>())}]");
                    DebugLogger.Info($"[ON_REFRESH_CLICK] UI Selections - Reference Files: {selectedReferenceFiles?.Count ?? 0} [{string.Join(", ", selectedReferenceFiles ?? new List<string>())}]");
                    DebugLogger.Info($"[ON_REFRESH_CLICK] UI Selections - Host Files: {selectedHostFiles?.Count ?? 0} [{string.Join(", ", selectedHostFiles ?? new List<string>())}]");
                    
                    System.Diagnostics.Debug.WriteLine($"[ON_REFRESH_CLICK] UI Selections - Filters: {selectedFilterItems?.Count ?? 0} [{string.Join(", ", selectedFilterItems ?? new List<string>())}]");
                    System.Diagnostics.Debug.WriteLine($"[ON_REFRESH_CLICK] UI Selections - MEP Categories: {selectedMepCategories?.Count ?? 0} [{string.Join(", ", selectedMepCategories ?? new List<string>())}]");
                    System.Diagnostics.Debug.WriteLine($"[ON_REFRESH_CLICK] UI Selections - Reference Files: {selectedReferenceFiles?.Count ?? 0} [{string.Join(", ", selectedReferenceFiles ?? new List<string>())}]");
                    System.Diagnostics.Debug.WriteLine($"[ON_REFRESH_CLICK] UI Selections - Host Files: {selectedHostFiles?.Count ?? 0} [{string.Join(", ", selectedHostFiles ?? new List<string>())}]");
                    
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("logger_debug.txt"), 
                        $"[{DateTime.Now}] [ON_REFRESH_CLICK] UI Selections - Filters: {selectedFilterItems?.Count ?? 0} [{string.Join(", ", selectedFilterItems ?? new List<string>())}]\n" +
                        $"[{DateTime.Now}] [ON_REFRESH_CLICK] UI Selections - MEP Categories: {selectedMepCategories?.Count ?? 0} [{string.Join(", ", selectedMepCategories ?? new List<string>())}]\n" +
                        $"[{DateTime.Now}] [ON_REFRESH_CLICK] UI Selections - Reference Files: {selectedReferenceFiles?.Count ?? 0} [{string.Join(", ", selectedReferenceFiles ?? new List<string>())}]\n" +
                        $"[{DateTime.Now}] [ON_REFRESH_CLICK] UI Selections - Host Files: {selectedHostFiles?.Count ?? 0} [{string.Join(", ", selectedHostFiles ?? new List<string>())}]\n");

                    // Guard: No filters selected → prompt and STOP (do nothing else)
                    if (selectedFilterItems == null || selectedFilterItems.Count == 0)
                    {
                        System.Windows.Forms.MessageBox.Show(
                            "Please select at least one filter before running Refresh.",
                            "No Filter Selected",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Warning);
                        _statusLabel.Text = "Refresh cancelled - No filter selected";
                        _progressBar.Visible = false;
                        _refreshButton.Enabled = true;
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("logger_debug.txt"), $"[{DateTime.Now}] [ON_REFRESH_CLICK] Cancelled - No filter selected\n");
                        return;
                    }
                    
                    // ✅ SAVE FILTER CHECK: Prompt user to save filter if it's not saved
                    var selectedFilterName = selectedFilterItems.FirstOrDefault();
                    if (!string.IsNullOrWhiteSpace(selectedFilterName) && !_filterManagementService.IsFilterSaved(selectedFilterName))
                    {
                        // Bring main dialog to front
                        this.BringToFront();
                        this.Activate();
                        this.TopMost = true;
                        this.TopMost = false;
                        
                        var result = System.Windows.Forms.MessageBox.Show(
                            $"Filter '{selectedFilterName}' has not been saved.\n\nPlease save the filter before running Refresh.\n\nWould you like to save it now?",
                            "Filter Not Saved",
                            System.Windows.Forms.MessageBoxButtons.YesNo,
                            System.Windows.Forms.MessageBoxIcon.Warning,
                            System.Windows.Forms.MessageBoxDefaultButton.Button1);
                        
                        if (result == System.Windows.Forms.DialogResult.Yes)
                        {
                            // Find the filter listbox and save the filter
                            WinForms.ListBox filterListBox = null;
                            if (_filtersPanel?.Controls.Count > 0)
                            {
                                foreach (var control in _filtersPanel.Controls)
                                {
                                    if (control is WinForms.ListBox lb)
                                    {
                                        filterListBox = lb;
                                        break;
                                    }
                                }
                            }
                            
                            if (filterListBox != null)
                            {
                                // Select the filter in listbox if not already selected
                                if (filterListBox.SelectedItem?.ToString() != selectedFilterName)
                                {
                                    for (int i = 0; i < filterListBox.Items.Count; i++)
                                    {
                                        if (filterListBox.Items[i]?.ToString() == selectedFilterName)
                                        {
                                            filterListBox.SelectedIndex = i;
                                            break;
                                        }
                                    }
                                }
                                
                                // Save the filter
                                SaveFilterWithUIState(filterListBox);
                                _statusLabel.Text = $"Filter '{selectedFilterName}' saved";
                            }
                            else
                            {
                                System.Windows.Forms.MessageBox.Show(
                                    "Could not find filter list. Please save the filter manually using the Save (↓) button.",
                                    "Save Filter",
                                    System.Windows.Forms.MessageBoxButtons.OK,
                                    System.Windows.Forms.MessageBoxIcon.Information);
                                _statusLabel.Text = "Please save filter manually";
                                _refreshButton.Enabled = true;
                                return;
                            }
                        }
                        else
                        {
                            // User chose not to save - cancel refresh
                            _statusLabel.Text = "Refresh cancelled - Filter not saved";
                            _refreshButton.Enabled = true;
                            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("logger_debug.txt"), $"[{DateTime.Now}] [ON_REFRESH_CLICK] Cancelled - Filter not saved\n");
                            return;
                        }
                    }

                    // ✅ FEATURE FLAG: Use factory to create appropriate refresh service (legacy or refactored)
                    DebugLogger.Info("[ON_REFRESH_CLICK] Creating refresh service...");
                    System.Diagnostics.Debug.WriteLine("[ON_REFRESH_CLICK] Creating refresh service...");
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("logger_debug.txt"), 
                        $"[{DateTime.Now}] [ON_REFRESH_CLICK] Creating refresh service...\n");
                    
                    var refreshService = Services.RefreshServiceFactory.Create(document, _uiDocument, _appProfileService);
                    refreshService.SetUIReferences(_statusLabel, _progressBar, _refreshButton);
                    
                    // 🔥 CRITICAL FIX: Load existing clash zone data to preserve cluster information
                    // This prevents the Refresh process from overwriting existing cluster data
                    // Note: Refactored service loads data internally, but we call this for legacy compatibility
                    DebugLogger.Info("[ON_REFRESH_CLICK] Loading existing clash zone data...");
                    System.Diagnostics.Debug.WriteLine("[ON_REFRESH_CLICK] Loading existing clash zone data...");
                    refreshService.LoadExistingClashZoneData();
                    
                    DebugLogger.Info("[ON_REFRESH_CLICK] Calling ExecuteRefresh...");
                    System.Diagnostics.Debug.WriteLine("[ON_REFRESH_CLICK] Calling ExecuteRefresh...");
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("logger_debug.txt"), 
                        $"[{DateTime.Now}] [ON_REFRESH_CLICK] Calling ExecuteRefresh...\n");
                    
                    // ✅ ExecuteRefresh returns void (via IRefreshService interface)
                    // The wrapper logs the result internally (Result.Succeeded/Cancelled/Failed)
                    refreshService.ExecuteRefresh(selectedFilterItems, selectedMepCategories, selectedReferenceFiles, selectedHostFiles, clearanceSettings);
                    
                    DebugLogger.Info("[ON_REFRESH_CLICK] ExecuteRefresh completed");
                    System.Diagnostics.Debug.WriteLine("[ON_REFRESH_CLICK] ExecuteRefresh completed");
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("logger_debug.txt"), 
                        $"[{DateTime.Now}] [ON_REFRESH_CLICK] ExecuteRefresh completed\n");
                    
                    // ✅ TIMING: Stop timing and log elapsed time
                    refreshStopwatch.Stop();
                    var endTime = DateTime.Now;
                    var elapsedMs = refreshStopwatch.ElapsedMilliseconds;
                    var elapsedSeconds = elapsedMs / 1000.0;
                    
                    DebugLogger.Info($"=== REFRESH COMPLETED in {elapsedSeconds:F2} seconds ({elapsedMs}ms) ===");
                    System.Diagnostics.Debug.WriteLine($"[ON_REFRESH_CLICK] === REFRESH COMPLETED in {elapsedSeconds:F2} seconds === END TIME: {endTime:HH:mm:ss.fff} ===");
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("refresh_timing.log"), 
                        $"[{endTime:HH:mm:ss.fff}] === REFRESH COMPLETED ===\n" +
                        $"Duration: {elapsedSeconds:F2} seconds ({elapsedMs}ms)\n" +
                        $"Start: {startTime:HH:mm:ss.fff} → End: {endTime:HH:mm:ss.fff}\n\n");
                    
                    // ⚠️ CRITICAL FIX: Skip parameter dropdown updates during refresh to prevent sleeve deletion
                    // Parameter dropdowns don't need to be updated during refresh operations
                    System.Diagnostics.Debug.WriteLine("[ON_REFRESH_CLICK] SKIPPING parameter dropdown updates during refresh to prevent sleeve deletion");
                    // DISABLED: Excessive logging - JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("logger_debug.txt"), $"[{DateTime.Now}] SKIPPING parameter dropdown updates during refresh to prevent sleeve deletion\n");
                    
                    // Update parameter dropdowns after successful refresh
                    // System.Diagnostics.Debug.WriteLine("[ON_REFRESH_CLICK] About to update parameter dropdowns");
                    // JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("logger_debug.txt"), $"[{DateTime.Now}] About to update parameter dropdowns\n");
                    
                    // UpdateParameterDropdownsFromMepCategories(document);

                    // Also update host parameters for the Host to Opening tab
                    // System.Diagnostics.Debug.WriteLine("[ON_REFRESH_CLICK] About to update host parameters");
                    // PopulateHostParameters();

                    // System.Diagnostics.Debug.WriteLine("[ON_REFRESH_CLICK] Parameter dropdowns updated successfully");
                    // JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("logger_debug.txt"), $"[{DateTime.Now}] Parameter dropdowns updated successfully\n");
                    
                    DebugLogger.Info("[OK_BUTTON_DEBUG] About to check OK button enabling logic");
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("logger_debug.txt"), $"[{DateTime.Now}] [OK_BUTTON_DEBUG] About to check OK button enabling logic\n");

                    // Gate: Enable OK only if there are unresolved clash zones after refresh
                    // ✅ PHASE 2: DATABASE-FIRST - Check database for unresolved zones (primary source of truth)
                    // Fallback to Global XML only if database has no data (backward compatibility)
                    try
                    {
                        DebugLogger.Info("[OK_BUTTON_DEBUG] Checking database for unresolved zones (database-first)");
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("logger_debug.txt"), $"[{DateTime.Now}] [OK_BUTTON_DEBUG] Checking database for unresolved zones\n");
                        
                        int unresolvedCount = 0;
                        bool usedDatabase = false;
                        
                        if (document != null)
                        {
                            // ✅ PHASE 2: Try database first
                            try
                            {
                                using (var dbContext = new Data.SleeveDbContext(document))
                                {
                                    var repository = new Data.Repositories.ClashZoneRepository(dbContext);
                                    
                                    // ✅ CRITICAL FIX: Reuse selectedFilterName from outer scope (already declared at line 6862)
                                    // Use GetClashZonesByFilter with selected filter name to ensure we only check zones for the selected filter
                                    if (selectedMepCategories != null && selectedMepCategories.Count > 0)
                                    {
                                        foreach (var category in selectedMepCategories)
                                        {
                                            if (string.IsNullOrWhiteSpace(category))
                                                continue;
                                            
                                            try
                                            {
                                                // ✅ CRITICAL FIX: Use GetClashZonesByFilter with selected filter name (not GetClashZonesByCategory)
                                                // This ensures we only check zones for the selected filter, not all filters
                                                var allZones = repository.GetClashZonesByFilter(selectedFilterName, category, unresolvedOnly: false);
                                                if (allZones != null && allZones.Count > 0)
                                                {
                                                    // ✅ CRITICAL FIX: A zone needs placement if:
                                                    // 1. BOTH IsResolved=false AND IsClusterResolved=false (no sleeve at all), OR
                                                    // 2. ReadyForPlacementFlag=true (sleeve was deleted, ready to place again)
                                                    int categoryUnresolved = allZones.Count(z => 
                                                        (!z.IsResolved && !z.IsClusterResolved) || 
                                                        z.ReadyForPlacement == true);
                                                    unresolvedCount += categoryUnresolved;
                                                    usedDatabase = true;
                                                    
                                                    if (categoryUnresolved > 0)
                                                    {
                                                        int noSleeveCount = allZones.Count(z => !z.IsResolved && !z.IsClusterResolved);
                                                        int readyForPlacementCount = allZones.Count(z => z.ReadyForPlacement == true);
                                                        DebugLogger.Info($"[OK_BUTTON_DEBUG] Database Filter='{selectedFilterName}', Category='{category}': {categoryUnresolved} unresolved out of {allZones.Count} total zones (NoSleeve={noSleeveCount}, ReadyForPlacement={readyForPlacementCount})");
                                                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("logger_debug.txt"), $"[{DateTime.Now}] [OK_BUTTON_DEBUG] Database Filter='{selectedFilterName}', Category='{category}': {categoryUnresolved}/{allZones.Count} unresolved (NoSleeve={noSleeveCount}, ReadyForPlacement={readyForPlacementCount})\n");
                                                    }
                                                }
                                            }
                                            catch (Exception dbEx)
                                            {
                                                DebugLogger.Warning($"[OK_BUTTON_DEBUG] Error checking database for filter='{selectedFilterName}', category='{category}': {dbEx.Message}");
                                                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("logger_debug.txt"), $"[{DateTime.Now}] [OK_BUTTON_DEBUG] ERROR checking database for Filter='{selectedFilterName}', Category='{category}': {dbEx.Message}\n");
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception dbEx)
                            {
                                DebugLogger.Warning($"[OK_BUTTON_DEBUG] Database check failed, falling back to Global XML: {dbEx.Message}");
                                usedDatabase = false;
                            }
                            
                            // ✅ FALLBACK: Use Global XML if database has no data (backward compatibility)
                            if (!usedDatabase || unresolvedCount == 0)
                            {
                                DebugLogger.Info("[OK_BUTTON_DEBUG] Database has no data. Global XML fallback removed (Legacy XML deleted).");
                                // Legacy XML logic removed entirely.
                            }
                        }
                        
                        _okButton.Enabled = unresolvedCount > 0;
                        DebugLogger.Info($"[OK_BUTTON_DEBUG] ✅ OK button enabled: {_okButton.Enabled} (unresolved: {unresolvedCount} from {(usedDatabase ? "database" : "Global XML fallback")})");
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("logger_debug.txt"), $"[{DateTime.Now}] [OK_BUTTON_DEBUG] ✅ OK button enabled: {_okButton.Enabled} (unresolved: {unresolvedCount} from {(usedDatabase ? "database" : "Global XML")})\n");
                    }
                    catch (Exception ex)
                    {
                        // DISABLED: Excessive logging - JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("logger_debug.txt"), $"[{DateTime.Now}] [OK_BUTTON_DEBUG] ❌ ERROR enabling OK button: {ex.Message}\n");
                        // DISABLED: Excessive logging - JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("logger_debug.txt"), $"[{DateTime.Now}] [OK_BUTTON_DEBUG] Stack: {ex.StackTrace}\n");
                        _okButton.Enabled = false; 
                    }
                }
                else
                {
                    // ✅ CRITICAL FIX: Log when document is null (prevents silent failure)
                    DebugLogger.Warning("[ON_REFRESH_CLICK] ❌ Document is null - refresh cannot proceed");
                    System.Diagnostics.Debug.WriteLine("[ON_REFRESH_CLICK] ❌ Document is null - refresh cannot proceed");
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("logger_debug.txt"), 
                        $"[{DateTime.Now}] [ON_REFRESH_CLICK] ❌ Document is null - refresh cannot proceed\n");
                    
                    _statusLabel.Text = "No active document";
                    _progressBar.Visible = false;
                    _refreshButton.Enabled = true;
                    
                    System.Windows.Forms.MessageBox.Show(
                        "No active Revit document found. Please open a Revit project file before running Refresh.",
                        "No Document",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                }

                System.Diagnostics.Debug.WriteLine("[ON_REFRESH_CLICK] RefreshService.ExecuteRefresh() method completed successfully");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("logger_debug.txt"), $"[{DateTime.Now}] RefreshService.ExecuteRefresh() method completed successfully\n");
                    }
                    catch (Exception ex)
                    {
                        // ✅ TIMING: Stop timing even on error
                        var endTime = DateTime.Now;
                        var elapsedMs = refreshStopwatch.ElapsedMilliseconds;
                        var elapsedSeconds = elapsedMs / 1000.0;
                        
                        System.Diagnostics.Debug.WriteLine($"[ON_REFRESH_CLICK] ERROR: {ex.Message} (completed in {elapsedSeconds:F2}s)");
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("logger_debug.txt"), $"[{DateTime.Now}] ERROR in OnRefreshClick: {ex.Message}\n");
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("refresh_timing.log"), 
                            $"[{endTime:HH:mm:ss.fff}] === REFRESH FAILED ===\n" +
                            $"Duration: {elapsedSeconds:F2} seconds ({elapsedMs}ms)\n" +
                            $"Error: {ex.Message}\n\n");

                        _statusLabel.Text = $"Error during refresh: {ex.Message}";
                        _progressBar.Visible = false;
                        _refreshButton.Enabled = true;
                        DebugLogger.Error($"Error in OnRefreshClick (took {elapsedSeconds:F2}s): {ex.Message}");
                    }
        }
        
        private void PerformBasicRefresh()
        {
            try
            {
                _progressBar.Value = 25;
                _statusLabel.Text = "Updating element lists...";
                
                _progressBar.Value = 50;
                _statusLabel.Text = "Refreshing clearance settings...";
                
                var clearances = GetClearanceSettings();
                DebugLogger.Info($"Refreshed {clearances.Count} clearance settings");
                
                _progressBar.Value = 75;
                _statusLabel.Text = "Finalizing refresh...";
                
                _progressBar.Value = 100;
                _statusLabel.Text = "Basic refresh complete";
                
                var timer = new System.Windows.Forms.Timer();
                timer.Interval = 1500;
                timer.Tick += (s, args) => {
                    _progressBar.Visible = false;
                    timer.Stop();
                    timer.Dispose();
                };
                timer.Start();
                
                _refreshButton.Enabled = true;
            }
            catch (Exception ex)
            {
                _statusLabel.Text = $"Error during basic refresh: {ex.Message}";
                _progressBar.Visible = false;
                _refreshButton.Enabled = true;
                DebugLogger.Error($"Error in PerformBasicRefresh: {ex.Message}");
            }
        }

        /// <summary>
        /// Updates parameter dropdowns using the existing GetParametersForCategory method
        /// NOTE: This method is no longer used as parameter service has been moved to separate dialog.
        /// </summary>
        private void PopulateParameterDropdowns()
        {
            // Parameter service moved to separate dialog - this method is no longer used
            DebugLogger.Info("[PARAMETER_SERVICE] Parameter service moved to separate dialog - method no longer used");
        }

        /// <summary>
        /// Populates host parameter dropdowns using HostParameterService
        /// NOTE: This method is no longer used as parameter service has been moved to separate dialog.
        /// </summary>
        public void PopulateHostParameters()
        {
            // Parameter service moved to separate dialog - this method is no longer used
            DebugLogger.Info("[HOST_SERVICE] Parameter service moved to separate dialog - method no longer used");
        }


        /// <summary>
        /// Gets MEP parameters for the current tab context
        /// NOTE: This method is no longer used as parameter service has been moved to separate dialog.
        /// </summary>
        private List<string> GetMepParametersForTab(WinForms.Panel servicePanel)
        {
            // Parameter service moved to separate dialog - this method is no longer used
            DebugLogger.Info("[MEP_PARAMETERS] Parameter service moved to separate dialog - method no longer used");
            return new List<string>();
        }

        /// <summary>
        /// Gets default MEP parameters that should be pre-populated in parameter rows
        /// </summary>
        private List<string> GetDefaultMepParameters()
        {
            return new List<string>
            {
                "Reference Level",
                "Width", 
                "Height",
                "Diameter",
                "System Type",
                "Service Type",
                "Outside Diameter",
                "Inside Diameter",
                "Mark",
                "Type Mark"
            };
        }

        /// <summary>
        /// Gets default opening parameters that should be pre-populated in parameter rows
        /// </summary>
        private List<string> GetDefaultOpeningParameters()
        {
            return new List<string>
            {
                "Mark",
                "Type Mark", 
                "Assembly Code",
                "Assembly Description",
                "Height",
                "Width",
                "Diameter",
                "Clearance",
                "Opening Type",
                "Size",
            };
        }
        /// <summary>
        /// Creates default parameter rows with common MEP and opening parameters
        /// </summary>
        private void CreateDefaultParameterRows(WinForms.TabPage tabPage, List<string> mepParameters, List<string> openingParameters)
        {
            try
            {
                DebugLogger.Info($"[PARAMETER_DEFAULT] Creating default parameter rows for tab '{tabPage.Text}'");
                
                // Find the service panel
                var servicePanel = tabPage.Controls.OfType<WinForms.Panel>().FirstOrDefault();
                if (servicePanel == null)
                {
                    DebugLogger.Warning($"[PARAMETER_DEFAULT] No service panel found in tab '{tabPage.Text}'");
                    return;
                }

                // Get default parameters
                var defaultMepParams = GetDefaultMepParameters();
                var defaultOpeningParams = GetDefaultOpeningParameters();
                
                // Filter to only include parameters that exist in the available lists
                var availableMepDefaults = defaultMepParams.Where(p => mepParameters.Contains(p)).ToList();
                var availableOpeningDefaults = defaultOpeningParams.Where(p => openingParameters.Contains(p)).ToList();
                
                DebugLogger.Info($"[PARAMETER_DEFAULT] Found {availableMepDefaults.Count} available MEP defaults and {availableOpeningDefaults.Count} available opening defaults");
                
                // Create rows for each available default parameter
                int yPosition = 10;
                int rowHeight = 25;
                
                for (int i = 0; i < Math.Max(availableMepDefaults.Count, availableOpeningDefaults.Count); i++)
                {
                    // Create row panel
                    var rowPanel = new WinForms.Panel
                    {
                        Location = new System.Drawing.Point(5, yPosition),
                        Size = new System.Drawing.Size(servicePanel.Width - 10, rowHeight),
                        BorderStyle = WinForms.BorderStyle.FixedSingle
                    };
                    
                    // Create MEP parameter ComboBox (left side)
                    if (i < availableMepDefaults.Count)
                    {
                        var mepComboBox = new WinForms.ComboBox
                        {
                            Name = $"mepParamCombo_{tabPage.Text}_{i}",
                            Location = new System.Drawing.Point(5, 2),
                            Size = new System.Drawing.Size(150, 20),
                            DropDownStyle = WinForms.ComboBoxStyle.DropDownList,
                            Tag = "mep"
                        };
                        mepComboBox.Items.AddRange(mepParameters.ToArray());
                        mepComboBox.SelectedItem = availableMepDefaults[i]; // Set default selection
                        rowPanel.Controls.Add(mepComboBox);
                        DebugLogger.Info($"[PARAMETER_DEFAULT] Created MEP ComboBox with default: {availableMepDefaults[i]}");
                    }
                    
                    // Create Opening parameter ComboBox (right side)
                    if (i < availableOpeningDefaults.Count)
                    {
                        var openingComboBox = new WinForms.ComboBox
                        {
                            Name = $"openingParamCombo_{tabPage.Text}_{i}",
                            Location = new System.Drawing.Point(160, 2),
                            Size = new System.Drawing.Size(150, 20),
                            DropDownStyle = WinForms.ComboBoxStyle.DropDownList,
                            Tag = "opening"
                        };
                        openingComboBox.Items.AddRange(openingParameters.ToArray());
                        openingComboBox.SelectedItem = availableOpeningDefaults[i]; // Set default selection
                        rowPanel.Controls.Add(openingComboBox);
                        DebugLogger.Info($"[PARAMETER_DEFAULT] Created Opening ComboBox with default: {availableOpeningDefaults[i]}");
                    }
                    
                    // Add remove button
                    var removeButton = new WinForms.Button
                    {
                        Text = "×",
                        Location = new System.Drawing.Point(rowPanel.Width - 25, 1),
                        Size = new System.Drawing.Size(20, 20),
                        Tag = "remove"
                    };
                    removeButton.Click += (s, e) => {
                        servicePanel.Controls.Remove(rowPanel);
                        rowPanel.Dispose();
                    };
                    rowPanel.Controls.Add(removeButton);
                    
                    servicePanel.Controls.Add(rowPanel);
                    yPosition += rowHeight + 5;
                }
                
                DebugLogger.Info($"[PARAMETER_DEFAULT] Created {Math.Max(availableMepDefaults.Count, availableOpeningDefaults.Count)} default parameter rows");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[PARAMETER_DEFAULT] Error creating default parameter rows: {ex.Message}");
            }
        }
        /// <summary>
        /// Updates parameter dropdowns with parameters from MEP categories (not clash zones)
        /// </summary>
        public void UpdateParameterDropdownsFromMepCategories(Document document)
        {
            try
            {
                DebugLogger.Info("[PARAMETER_DEBUG] Starting parameter dropdown update using ParameterExtractionService");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("refresh_debug.log"), $"[{DateTime.Now}] [PARAMETER_DEBUG] Starting parameter dropdown update using ParameterExtractionService\n");
                
                // ⚠️ CRITICAL FIX: Skip parameter extraction during initialization to prevent sleeve deletion
                if (_isInitializing)
                {
                    DebugLogger.Info("🔍 SKIPPING UpdateParameterDropdownsFromMepCategories during initialization - using cached parameters");
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("refresh_debug.log"), $"[{DateTime.Now}] 🔍 SKIPPING UpdateParameterDropdownsFromMepCategories during initialization - using cached parameters\n");
                    return;
                }
                
                var mepParameters = new HashSet<string>();
                var openingParameters = new HashSet<string>();
                
                // Get currently selected MEP category from UI
                var selectedMepCategory = GetSelectedMepCategory();
                DebugLogger.Info($"[PARAMETER_DEBUG] Selected MEP category: {selectedMepCategory}");
                
                // Get currently selected reference files from UI
                var selectedReferenceFiles = GetSelectedReferenceFiles();
                DebugLogger.Info($"[PARAMETER_DEBUG] Selected reference files: {string.Join(", ", selectedReferenceFiles)}");
                
                // If no reference files selected, fall back to current document
                if (selectedReferenceFiles.Count == 0)
                {
                    DebugLogger.Info("[PARAMETER_DEBUG] No reference files selected, using current document");
                    selectedReferenceFiles.Add("Current Document");
                }
                
                // Use ParameterExtractionService to collect parameters from selected linked files
                var parameterExtractionService = new ParameterExtractionService();
                
                foreach (var referenceFile in selectedReferenceFiles)
                {
                    DebugLogger.Info($"[PARAMETER_DEBUG] Processing reference file: {referenceFile}");
                    
                    Document targetDocument = document;
                    
                    // If it's a linked file, get the linked document
                    if (referenceFile != "Current Document")
                    {
                        var linkedFileService = _linkedFileService;
                        if (linkedFileService != null)
                        {
                            var linkedFiles = linkedFileService.GetLinkedFiles(document);
                            
                            // Extract just the filename part (remove element count and parentheses)
                            var cleanFileName = referenceFile.Split('(')[0].Trim();
                            DebugLogger.Info($"[PARAMETER_DEBUG] Looking for linked file: '{cleanFileName}' in {linkedFiles.Count} available files");
                            
                            var linkedFile = linkedFiles.FirstOrDefault(lf => lf.FileName == cleanFileName);
                            if (linkedFile?.LinkInstance?.GetLinkDocument() != null)
                            {
                                targetDocument = linkedFile.LinkInstance.GetLinkDocument();
                                DebugLogger.Info($"[PARAMETER_DEBUG] Using linked document: {linkedFile.FileName}");
                            }
                            else
                            {
                                DebugLogger.Warning($"[PARAMETER_DEBUG] Could not find linked document for: '{cleanFileName}'. Available files: {string.Join(", ", linkedFiles.Select(lf => lf.FileName))}");
                                continue;
                            }
                        }
                        else
                        {
                            DebugLogger.Warning("[PARAMETER_DEBUG] LinkedFileService not available");
                            continue;
                        }
                    }
                    
                    // Use ParameterExtractionService to get parameters for the selected MEP category
                    var mepCategories = new List<Services.MepCategory> { selectedMepCategory };
                    var parameterInfos = parameterExtractionService.GetParametersForMepCategories(targetDocument, mepCategories);
                    
                    DebugLogger.Info($"[PARAMETER_DEBUG] Found {parameterInfos.Count} parameters from {selectedMepCategory} in {referenceFile}");
                    
                    // Extract parameter names for dropdown
                    foreach (var paramInfo in parameterInfos)
                    {
                        if (!string.IsNullOrEmpty(paramInfo.Name) && 
                            !paramInfo.Name.StartsWith("Internal") &&
                            !paramInfo.Name.StartsWith("Revit") &&
                            !paramInfo.Name.StartsWith("Assembly"))
                        {
                            mepParameters.Add(paramInfo.Name);
                        }
                    }
                }
                
                // Collect parameters from opening families (always from current document)
                var openingFamilies = GetOpeningFamilies(document);
                DebugLogger.Info($"[PARAMETER_DEBUG] Found {openingFamilies.Count} opening families");

                foreach (var family in openingFamilies)
                {
                    DebugLogger.Info($"[PARAMETER_DEBUG] Processing opening family: {family.Name}");

                    foreach (Parameter param in family.Parameters)
                    {
                        // Include all parameters except those with empty names or truly internal parameters
                        // Removed overly restrictive filtering - include all valid parameters
                        if (!string.IsNullOrEmpty(param.Definition.Name))
                        {
                            openingParameters.Add(param.Definition.Name);
                            DebugLogger.Info($"[PARAMETER_DEBUG] Added parameter: {param.Definition.Name}");
                        }
                    }
                }
                
                DebugLogger.Info($"[PARAMETER_DEBUG] Found {mepParameters.Count} MEP parameters and {openingParameters.Count} opening parameters");

                // Parameter service moved to separate dialog - this functionality is now handled in ParameterServiceDialog
                DebugLogger.Info($"[PARAMETER_DEBUG] Parameter service moved to separate dialog - no longer updating parameter service dropdowns here");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("refresh_debug.log"), $"[{DateTime.Now}] [PARAMETER_DEBUG] Parameter service moved to separate dialog - no longer updating parameter service dropdowns here\n");

                // Store parameters globally for later use
                _allCollectedParameters["MEP Elements"] = mepParameters.Select(p => new Models.ParameterInfo 
                { 
                    Name = p, 
                    Type = "Text", 
                    IsReadOnly = false 
                }).ToList();

                _allCollectedParameters["Opening Families"] = openingParameters.Select(p => new Models.ParameterInfo
                {
                    Name = p,
                    Type = "Text",
                    IsReadOnly = false
                }).ToList();

                DebugLogger.Info("[PARAMETER_DEBUG] Parameter dropdown update completed successfully");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[PARAMETER_DEBUG] Error in UpdateParameterDropdownsFromMepCategories: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Filters clash zones by MEP element category
        /// </summary>
        private List<ClashZone> FilterClashZonesByCategory(List<ClashZone> clashZones, string category, Document document)
        {
            var filteredZones = new List<ClashZone>();
            
            try
            {
                DebugLogger.Info($"[CLASH_DEBUG] Filtering {clashZones.Count} clash zones for category: {category}");
                
                foreach (var clashZone in clashZones)
                {
                    try
                    {
                        // Get MEP element from document or linked documents
                        var mepElement = GetElementFromDocumentOrLinked(document, clashZone.MepElementId);
                        
                        if (mepElement != null)
                        {
                            // Determine element category
                            var elementCategory = GetElementCategory(mepElement);
                            
                            // Check if element belongs to the requested category
                            if (IsElementInCategory(elementCategory, category))
                            {
                                filteredZones.Add(clashZone);
                                DebugLogger.Info($"[CLASH_DEBUG] Clash zone {clashZone.Id} matches category '{category}' - element category: {elementCategory}");
                            }
                        }
                        else
                        {
                            DebugLogger.Warning($"[CLASH_DEBUG] MEP element {clashZone.MepElementId} not found for clash zone {clashZone.Id}");
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
                // First try host document
                var element = document.GetElement(elementId);
                if (element != null)
                {
                    return element;
                }
                
                // If not found, search linked documents
                var linkInstances = new FilteredElementCollector(document)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>();
                
                foreach (var link in linkInstances)
                {
                    var linkDoc = link.GetLinkDocument();
                    if (linkDoc != null)
                    {
                        element = linkDoc.GetElement(elementId);
                        if (element != null)
                        {
                            return element;
                        }
                    }
                }
                
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
                    case "ducts":
                        // Match: Duct Curves, Duct Fitting, Duct Terminal (but not Duct Accessory)
                        return (elementCatLower.Contains("duct") && elementCatLower.Contains("curve")) ||
                               (elementCatLower.Contains("duct") && elementCatLower.Contains("fitting")) ||
                               (elementCatLower.Contains("duct") && elementCatLower.Contains("terminal"));
                               
                    case "duct accessories":
                        // Match: Duct Accessory
                        return elementCatLower.Contains("duct") && elementCatLower.Contains("accessory");
                        
                    case "pipes":
                        // Match: Pipe Curves, Pipe Fitting, Pipe Terminal
                        return (elementCatLower.Contains("pipe") && elementCatLower.Contains("curve")) ||
                               (elementCatLower.Contains("pipe") && elementCatLower.Contains("fitting")) ||
                               (elementCatLower.Contains("pipe") && elementCatLower.Contains("terminal"));
                        
                    case "cable trays":
                        // Match: Cable Tray, Cable Tray Fitting
                        return (elementCatLower.Contains("cable") && elementCatLower.Contains("tray"));
                               
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
        /// Gets opening families from the current document
        /// </summary>
        private List<FamilySymbol> GetOpeningFamilies(Document document)
        {
            var openingFamilies = new List<FamilySymbol>();

            try
            {
                // Specific opening family names to filter by
                var targetFamilyNames = new List<string>
                {
                    "RectangularOpeningOnWall",
                    "RectangularOpeningOnSlab",
                    "CircularOpeningOnWall",
                    "CircularOpeningOnSlab"
                };

                // FamilySymbol is an ElementType; do NOT filter with WhereElementIsNotElementType
                var collector = new FilteredElementCollector(document)
                    .OfClass(typeof(FamilySymbol));

                int inspected = 0;
                foreach (Element element in collector)
                {
                    inspected++;
                    if (element is FamilySymbol familySymbol)
                    {
                        var familyName = familySymbol.Family?.Name ?? "";
                        var symbolName = familySymbol.Name ?? "";

                        // Check if this family matches our target families
                        bool isTargetFamily = targetFamilyNames.Any(targetName =>
                            familyName.Contains(targetName) ||
                            symbolName.Contains(targetName) ||
                            $"{familyName} {symbolName}".Contains(targetName));

                        if (isTargetFamily)
                        {
                            openingFamilies.Add(familySymbol);
                        }
                    }
                }

                DebugLogger.Info($"[PARAMETER_DEBUG] GetOpeningFamilies inspected {inspected} FamilySymbols, matched {openingFamilies.Count} families from specific opening families: {string.Join(", ", targetFamilyNames)}");

                // Enhanced debugging: Log some sample family names to help diagnose
                if (inspected > 0 && openingFamilies.Count == 0)
                {
                    DebugLogger.Info($"[PARAMETER_DEBUG] No specific opening families found. Sample family names in document:");
                    var sampleCollector = new FilteredElementCollector(document)
                        .OfClass(typeof(FamilySymbol))
                        .Take(10); // Just get first 10 for debugging

                    foreach (Element element in sampleCollector)
                    {
                        if (element is FamilySymbol familySymbol)
                        {
                            var familyName = $"{familySymbol.Family?.Name} {familySymbol.Name}";
                            var categoryName = familySymbol.Category?.Name ?? "Unknown";
                            DebugLogger.Info($"[PARAMETER_DEBUG] Sample family: '{familyName}' (Category: {categoryName})");
                        }
                    }
                }

                // Fallback: If no opening families found, return empty list
                // The UI will use the hardcoded GetOpeningSleeveParameters() method instead
                if (openingFamilies.Count == 0)
                {
                    DebugLogger.Info($"[PARAMETER_DEBUG] No specific opening families found - UI will use hardcoded opening parameters");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[PARAMETER_DEBUG] Error getting opening families: {ex.Message}");
            }

            return openingFamilies;
        }

        private void PerformClashDetectionAndSaveToFilter(List<OpeningFilter> selectedFilters)
        {
            try
            {
                _progressBar.Value = 10;
                _statusLabel.Text = "Starting clash detection...";

                // Get current intersections
                var currentIntersections = GetCurrentIntersections();
                DebugLogger.Info($"Refresh: Found {currentIntersections.Count} current intersections");

                if (currentIntersections.Count == 0)
                {
                    DebugLogger.Warning("Refresh: No intersections found - no clash zones to save");
                    _statusLabel.Text = "No intersections found";
                    return;
                }

                // Save clash zones to the first enabled filter
                var targetFilter = selectedFilters.FirstOrDefault(f => f.IsEnabled);
                if (targetFilter != null)
                {
                    // Create clash zone storage
                    var clashZoneStorage = new ClashZoneStorage
                    {
                        ClashZones = new List<ClashZone>(),
                        LastUpdated = DateTime.Now,
                        DocumentHash = _document?.PathName ?? "Unknown"
                    };

                    // Convert intersections to clash zones
                    foreach (var intersection in currentIntersections)
                    {
                        var mepElement = intersection.Item1;
                        var structuralElement = intersection.Item2;
                        var intersectionBBox = intersection.Item3;
                        var intersectionPoint = intersection.Item4;

                        var clashZone = new ClashZone
                        {
                            MepElementId = mepElement.Id,
                            StructuralElementId = structuralElement.Id,
                            IntersectionPoint = intersectionPoint,
                            ClashBoundingBox = intersectionBBox,
                            DetectedAt = DateTime.Now,
                            IsResolved = false
                        };
                        clashZoneStorage.ClashZones.Add(clashZone);
                    }

                    targetFilter.ClashZoneStorage = clashZoneStorage;
                    targetFilter.LastModified = DateTime.Now;

                    DebugLogger.Info($"Refresh: Saved {clashZoneStorage.ClashZones.Count} clash zones to filter '{targetFilter.Name}'");
                    _statusLabel.Text = $"Saved {clashZoneStorage.ClashZones.Count} clash zones to filter";
                }
                else
                {
                    DebugLogger.Warning("Refresh: No enabled filter found to save clash zones");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"PerformClashDetectionAndSaveToFilter failed: {ex.Message}");
                throw;
            }
        }
        /// <summary>
        /// Gets the current profile from the application
        /// </summary>
        private UserProfile? GetCurrentProfile()
        {
            try
            {
                // Get the current profile from the application profile service
                var profile = _appProfileService?.GetCurrentProfile();
                DebugLogger.Info($"GetCurrentProfile: Retrieved profile: {(profile?.Name ?? "NULL")}");
                return profile;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error getting current profile: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// Gets current intersections using MepIntersectionService
        /// </summary>
        private List<(Element, Element, BoundingBoxXYZ, XYZ)> GetCurrentIntersections()
        {
            try
            {
                var document = GetCurrentDocument();
                if (document == null)
                {
                    DebugLogger.Warning("GetCurrentIntersections: No document available");
                    return new List<(Element, Element, BoundingBoxXYZ, XYZ)>();
                }

                var view3D = document.ActiveView as View3D;
                if (view3D == null)
                {
                    DebugLogger.Warning("No 3D view active. Cannot detect intersections.");
                    return new List<(Element, Element, BoundingBoxXYZ, XYZ)>();
                }

                DebugLogger.Info("GetCurrentIntersections: Starting intersection detection using TestMepIntersection service...");

                // Use the proven TestMepIntersection service with current UI selections
                var selectedCategories = GetSelectedMepCategories();
                var selectedHostCategories = Services.FilterUiStateProvider.GetSelectedHostCategories?.Invoke() ?? new List<string>();
                var intersectionService = new IntersectionDetectionService(msg => DebugLogger.Info(msg));
                var intersections = intersectionService.FindIntersections(document, view3D, selectedCategories, null, null, selectedHostCategories);

                DebugLogger.Info($"GetCurrentIntersections: Found {intersections.Count} total intersections using TestMepIntersection service");
                return intersections;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"GetCurrentIntersections failed: {ex.Message}");
                return new List<(Element, Element, BoundingBoxXYZ, XYZ)>();
            }
        }
        /// <summary>
        /// Gets the current Revit document
        /// </summary>
        private Document? GetCurrentDocument()
        {
            try
            {
                // Use the document passed to the constructor
                if (_document != null)
                {
                    DebugLogger.Info($"GetCurrentDocument: Using document from constructor - {_document.Title}");
                    return _document;
                }
                
                // Try to get from UIDocument if available
                var uiDoc = GetCurrentUIDocument();
                if (uiDoc != null)
                {
                    DebugLogger.Info($"GetCurrentDocument: Using document from UIDocument - {uiDoc.Document.Title}");
                    return uiDoc.Document;
                }
                
                DebugLogger.Warning("GetCurrentDocument: No document available");
                return null;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"GetCurrentDocument failed: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// Gets the current UIDocument from Revit context
        /// </summary>
        private UIDocument? GetCurrentUIDocument()
        {
            try
            {
                if (_uiDocument != null)
                {
                    DebugLogger.Info($"GetCurrentUIDocument: Using UIDocument from constructor");
                    return _uiDocument;
                }
                
                DebugLogger.Warning("GetCurrentUIDocument: No UIDocument available");
                return null;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"GetCurrentUIDocument failed: {ex.Message}");
                return null;
            }
        }

        // Missing UI-related methods that were removed during refactoring
        private void OnConfigureClick(object sender, EventArgs e)
        {
            try
            {
                DebugLogger.Info("Configure button clicked - opening SettingsDialog");
                
                // Load current settings
                var settingsService = new SettingsService();
                var currentSettings = settingsService.LoadSettings();
                
                // Show settings dialog
                var settingsDialog = new SettingsDialog(currentSettings);
                var result = settingsDialog.ShowDialog(this);
                
                if (result == WinForms.DialogResult.OK)
                {
                    DebugLogger.Info("Settings dialog closed with OK - settings saved");
                    _statusLabel.Text = "Configuration updated successfully";
                }
                else
                {
                    DebugLogger.Info("Settings dialog closed with Cancel");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"OnConfigureClick failed: {ex.Message}");
                _statusLabel.Text = $"Failed to open settings: {ex.Message}";
            }
        }

        private void PositionPanels()
        {
            try
        {
            DebugLogger.Info("=== STARTING PositionPanels (3-Column Layout) ===");
            DebugLogger.Info($"Form ClientSize: {this.ClientSize.Width}x{this.ClientSize.Height}");

                // Calculate available height (excluding header)
                var headerHeight = _headerPanel?.Height ?? 80;
                var availableHeight = this.ClientSize.Height - headerHeight;
                DebugLogger.Info($"_headerPanel.Bottom: {headerHeight}");

                // Define panel widths
                var filtersWidth = 185;  // ✅ Reduced from 200 to 185 pixels (15px reduction) to make room for file listing section
                var splitterWidth = 3;
                var rightPanelWidth = 460;
                var leftPanelWidth = this.ClientSize.Width - filtersWidth - splitterWidth - rightPanelWidth - splitterWidth;

                DebugLogger.Info($"Calculated height: {availableHeight}");

                // Position filters panel (left column)
            if (_filtersPanel != null)
            {
                    _filtersPanel.Location = new System.Drawing.Point(0, headerHeight);
                    _filtersPanel.Size = new System.Drawing.Size(filtersWidth, availableHeight);
                    _filtersPanel.Visible = true;  // Ensure visibility
                    _filtersPanel.BringToFront();  // Bring to front
                DebugLogger.Info($"_filtersPanel positioned: Location={_filtersPanel.Location}, Size={_filtersPanel.Size}, Visible={_filtersPanel.Visible}");
            }

                // Position filters splitter
            if (_filtersSplitter != null)
            {
                    _filtersSplitter.Location = new System.Drawing.Point(filtersWidth, headerHeight);
                    _filtersSplitter.Size = new System.Drawing.Size(splitterWidth, availableHeight);
                DebugLogger.Info($"_filtersSplitter positioned: Location={_filtersSplitter.Location}, Size={_filtersSplitter.Size}");
            }

                // Position right panel (rightmost column)
                if (_rightPanel != null)
                {
                    _rightPanel.Location = new System.Drawing.Point(this.ClientSize.Width - rightPanelWidth, headerHeight);
                    _rightPanel.Size = new System.Drawing.Size(rightPanelWidth, availableHeight);
            DebugLogger.Info($"_rightPanel positioned: Location={_rightPanel.Location}, Size={_rightPanel.Size}");
                }

                // Position main splitter (between left and right panels)
                if (_mainSplitter != null)
                {
                    _mainSplitter.Location = new System.Drawing.Point(this.ClientSize.Width - rightPanelWidth - splitterWidth, headerHeight);
                    _mainSplitter.Height = availableHeight;
                    DebugLogger.Info($"_mainSplitter positioned: Location={_mainSplitter.Location}, Height={_mainSplitter.Height}");
                }

                // Position left panel (center column)
                if (_leftPanel != null)
                {
                    _leftPanel.Location = new System.Drawing.Point(filtersWidth + splitterWidth, headerHeight);
                    _leftPanel.Size = new System.Drawing.Size(leftPanelWidth, availableHeight);
                    DebugLogger.Info($"_leftPanel positioned: Location={_leftPanel.Location}, Size={_leftPanel.Size}");
                }

                DebugLogger.Info("=== PositionPanels (3-Column) COMPLETED ===");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"PositionPanels failed: {ex.Message}");
            }
        }

        // ========================================================================
        // ⚠️  CRITICAL UI LAYOUT - DO NOT MODIFY WITHOUT USER CONSENT  ⚠️
        // ========================================================================
        // This method provides manual positioning for the 2x2 grid layout
        // DO NOT CHANGE: This manual positioning is essential for proper layout
        // ========================================================================
        private void BalanceLeftLayout()
        {
            try
            {
                // Basic left panel layout balancing
                if (_topLeftPanel != null && _bottomLeftPanel != null && _leftPanel != null)
                {
                    int availableHeight = _leftPanel.Height - _horizontalSplitter.Height;
                    _topLeftPanel.Height = availableHeight / 2;
                    _bottomLeftPanel.Height = availableHeight / 2;
                    _bottomLeftPanel.Top = _topLeftPanel.Bottom + _horizontalSplitter.Height;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"BalanceLeftLayout failed: {ex.Message}");
            }
        }

        private void SaveFilterWithUIState(ListBox filterListBox)
        {
            // ✅ DEBUG LOG: Create dedicated log file for filter UI state persistence
            string filterDebugLogPath = SafeFileLogger.GetLogFilePath("filter_ui_state_debug.log");
            var logBuilder = new System.Text.StringBuilder();
            
            // ✅ CRITICAL: Log the file path so user knows where to find it
            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ========== FILTER UI STATE DEBUG LOG ==========");
            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Log file location: {filterDebugLogPath}");
            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Log directory exists: {System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(filterDebugLogPath))}");
            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ==============================================");
            
            // ✅ ALWAYS VISIBLE: Show in multiple places (even in deployment mode for debugging)
            // 1. DebugLogger (if enabled)
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[FILTER_UI_DEBUG] Log file location: {filterDebugLogPath}");
            }
            
            // 2. System.Diagnostics (always works, shows in Visual Studio Output window)
            System.Diagnostics.Debug.WriteLine($"[FILTER_UI_DEBUG] ========================================");
            System.Diagnostics.Debug.WriteLine($"[FILTER_UI_DEBUG] FILTER UI STATE DEBUG LOG FILE");
            System.Diagnostics.Debug.WriteLine($"[FILTER_UI_DEBUG] Location: {filterDebugLogPath}");
            System.Diagnostics.Debug.WriteLine($"[FILTER_UI_DEBUG] ========================================");
            
            // 3. Console (if console is available)
            try
            {
                Console.WriteLine($"[FILTER_UI_DEBUG] Log file location: {filterDebugLogPath}");
            }
            catch { }
            
            // 4. Write to log file immediately (so it's always there)
            try
            {
                System.IO.File.WriteAllText(filterDebugLogPath, logBuilder.ToString());
            }
            catch (Exception immediateWriteEx)
            {
                System.Diagnostics.Debug.WriteLine($"[FILTER_UI_DEBUG] ⚠️ Could not write log file immediately: {immediateWriteEx.Message}");
            }
            
            // 5. Show in UI status label (if available)
            try
            {
                if (_statusLabel != null)
                {
                    _statusLabel.Text = $"Debug log: {System.IO.Path.GetFileName(filterDebugLogPath)}";
                    _statusLabel.ForeColor = System.Drawing.Color.Blue;
                }
            }
            catch { }
            
            try
            {
                logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ========== SAVE FILTER WITH UI STATE STARTED ==========");
                logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] SaveFilterWithUIState called");
                DebugLogger.Info("SaveFilterWithUIState called");
                
                // ✅ CRITICAL FIX: Update UI state BEFORE saving (not after!)
                if (filterListBox.SelectedItem != null)
                {
                    var selectedFilterName = filterListBox.SelectedItem.ToString();
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Selected Filter Name: '{selectedFilterName}'");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] FilterListBox SelectedIndex: {filterListBox.SelectedIndex}");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] FilterListBox Items Count: {filterListBox.Items.Count}");
                    
                    if (_document == null)
                    {
                        DebugLogger.Error("[FILTER-SAVE] ❌ Document is null - cannot determine filter directory");
                        throw new InvalidOperationException("Document is null - cannot save filter");
                    }
                    
                    // ✅ OOP: Use ProjectPathService to get and ensure directory exists
                    ProjectPathService.EnsureFiltersDirectory(_document);
                    var filterDir = ProjectPathService.GetFiltersDirectory(_document);
                    var filePath = System.IO.Path.Combine(filterDir, $"{selectedFilterName}.xml");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Filter Directory: '{filterDir}'");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Filter File Path: '{filePath}'");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Filter File Exists: {System.IO.File.Exists(filePath)}");
                    
                    // ✅ DEBUG: Log current UI state BEFORE any operations
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ========== CURRENT UI STATE (BEFORE SAVE) ==========");
                    try
                    {
                        var currentHostCats = FilterUiStateProvider.GetSelectedHostCategories?.Invoke() ?? new List<string>();
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] SelectedHostCategories Count: {currentHostCats.Count}");
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] SelectedHostCategories: [{string.Join(", ", currentHostCats)}]");
                        
                        var currentMepCats = GetSelectedMepCategories();
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] SelectedMepCategories Count: {currentMepCats?.Count ?? 0}");
                        if (currentMepCats != null && currentMepCats.Count > 0)
                            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] SelectedMepCategories: [{string.Join(", ", currentMepCats)}]");
                        
                        var currentRefFiles = GetSelectedReferenceFiles();
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] SelectedReferenceFiles Count: {currentRefFiles?.Count ?? 0}");
                        
                        var currentHostFiles = GetSelectedHostFiles();
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] SelectedHostFiles Count: {currentHostFiles?.Count ?? 0}");
                    }
                    catch (Exception uiStateEx)
                    {
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ⚠️ Error reading UI state: {uiStateEx.Message}");
                    }
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ========== END CURRENT UI STATE ==========");
                    
                    OpeningFilter currentFilter;
                    
                    // ✅ LOGIC FIX: Only load filter if file exists, otherwise create new from UI state
                    if (System.IO.File.Exists(filePath))
                    {
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Filter file EXISTS - Loading existing filter");
                        DebugLogger.Info($"[FILTER-SAVE] Filter file exists, loading: {filePath}");
                        currentFilter = _filterManagementService.LoadFilterAuto(selectedFilterName);
                        if (currentFilter == null)
                        {
                            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ⚠️ LoadFilterAuto returned NULL - Creating new filter from UI state");
                            DebugLogger.Warning($"[FILTER-SAVE] LoadFilterAuto returned null, creating new filter from UI state");
                            currentFilter = _filterManagementService.CreateFilterFromCurrentUIState(selectedFilterName);
                            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ✅ Created new filter from UI state");
                            
                            // ✅ CRITICAL FIX: Update new filter with current UI state (including OpeningSettings)
                            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Calling UpdateFilterWithCurrentUIState for NEW filter...");
                            UpdateFilterWithCurrentUIState(currentFilter);
                            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ✅ UpdateFilterWithCurrentUIState completed for NEW filter");
                            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] New filter HostCats after update: {currentFilter?.SelectedHostCategories?.Count ?? 0}");
                            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] New filter OpeningSettings after update: {(currentFilter?.OpeningSettings != null ? "EXISTS" : "NULL")}");
                        }
                        else
                        {
                            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ✅ Loaded existing filter - HostCats before update: {currentFilter.SelectedHostCategories?.Count ?? 0}");
                            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ✅ Loaded existing filter - OpeningSettings before update: {(currentFilter.OpeningSettings != null ? "EXISTS" : "NULL")}");
                            
                            // ✅ FIX: Update the existing filter with current UI state BEFORE saving
                            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Calling UpdateFilterWithCurrentUIState...");
                            UpdateFilterWithCurrentUIState(currentFilter);
                            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ✅ UpdateFilterWithCurrentUIState completed");
                            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] HostCats after update: {currentFilter.SelectedHostCategories?.Count ?? 0}");
                            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] HostCats after update: [{string.Join(", ", currentFilter.SelectedHostCategories ?? new List<string>())}]");
                            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] OpeningSettings after update: {(currentFilter.OpeningSettings != null ? "EXISTS" : "NULL")}");
                            DebugLogger.Info($"Updated existing filter '{selectedFilterName}' with current UI state BEFORE save");
                        }
                    }
                    else
                    {
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Filter file does NOT exist - Creating new filter from UI state");
                        DebugLogger.Info($"[FILTER-SAVE] Filter file does NOT exist, creating new filter from UI state: {filePath}");
                        // ✅ CREATE NEW: Create filter from current UI state (new filter)
                        currentFilter = _filterManagementService.CreateFilterFromCurrentUIState(selectedFilterName);
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ✅ Created new filter from UI state");
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] New filter HostCats BEFORE update: {currentFilter?.SelectedHostCategories?.Count ?? 0}");
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] New filter OpeningSettings BEFORE update: {(currentFilter?.OpeningSettings != null ? "EXISTS" : "NULL")}");
                        
                        // ✅ CRITICAL FIX: Update new filter with current UI state (including OpeningSettings and AdoptToDocument)
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Calling UpdateFilterWithCurrentUIState for NEW filter (file doesn't exist)...");
                        UpdateFilterWithCurrentUIState(currentFilter);
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ✅ UpdateFilterWithCurrentUIState completed for NEW filter");
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] New filter HostCats AFTER update: {currentFilter?.SelectedHostCategories?.Count ?? 0}");
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] New filter HostCats AFTER update: [{string.Join(", ", currentFilter?.SelectedHostCategories ?? new List<string>())}]");
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] New filter OpeningSettings AFTER update: {(currentFilter?.OpeningSettings != null ? "EXISTS" : "NULL")}");
                        if (currentFilter?.OpeningSettings != null)
                        {
                            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] New filter OpeningSettings.AdoptToDocument: {currentFilter.OpeningSettings.AdoptToDocument}");
                        }
                        DebugLogger.Info($"Created new filter '{selectedFilterName}' from current UI state and updated with UI state");
                    }
                    
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ========== FILTER OBJECT STATE (BEFORE XML SAVE) ==========");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Filter Name: '{currentFilter?.Name}'");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Filter is null: {currentFilter == null}");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Directory exists: {System.IO.Directory.Exists(filterDir)}");
                    
                    if (currentFilter == null)
                    {
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ❌ ERROR: Filter is null after creation/load - cannot save");
                        DebugLogger.Error($"[FILTER-SAVE] ❌ Filter is null after creation/load - cannot save");
                        throw new InvalidOperationException($"Filter '{selectedFilterName}' is null - cannot save");
                    }
                    
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Filter SelectedHostCategories Count: {currentFilter.SelectedHostCategories?.Count ?? 0}");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Filter SelectedHostCategories: [{string.Join(", ", currentFilter.SelectedHostCategories ?? new List<string>())}]");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Filter OpeningSettings: {(currentFilter.OpeningSettings != null ? "EXISTS" : "NULL")}");
                    if (currentFilter.OpeningSettings != null)
                    {
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] OpeningSettings SelectedMepType: '{currentFilter.OpeningSettings.SelectedMepType}'");
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] OpeningSettings OpeningType: '{currentFilter.OpeningSettings.OpeningType}'");
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] OpeningSettings SleeveParameterPrefix: '{currentFilter.OpeningSettings.SleeveParameterPrefix}'");
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] OpeningSettings ClearanceSettings: {(currentFilter.OpeningSettings.ClearanceSettings != null ? "EXISTS" : "NULL")}");
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] OpeningSettings AdoptToDocument: {currentFilter.OpeningSettings.AdoptToDocument}");
                    }
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ========== END FILTER OBJECT STATE ==========");
                    
                    // ✅ STEP 1: Save to XML
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ========== STEP 1: SAVING TO XML ==========");
                    try
                    {
                        DebugLogger.Info($"[SAVE-DEBUG] SaveFilterToXmlFile: Name='{currentFilter?.Name}', HostCats={currentFilter?.SelectedHostCategories?.Count ?? 0}, OpenSettings={(currentFilter?.OpeningSettings != null ? "Yes" : "NULL")}");
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Calling SaveFilterToXmlFile...");
                        _filterManagementService.SaveFilterToXmlFile(currentFilter, filePath);
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ✅ SaveFilterToXmlFile completed");
                        DebugLogger.Info($"✅ Saved filter '{selectedFilterName}' with UI state to: {filePath}");
                        
                        // ✅ VERIFY: Check if file was actually created
                        if (System.IO.File.Exists(filePath))
                        {
                            var fileInfo = new System.IO.FileInfo(filePath);
                            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ✅ XML file verified: {fileInfo.Length} bytes");
                            DebugLogger.Info($"✅ Verified: Filter XML file exists ({fileInfo.Length} bytes) at: {filePath}");
                        }
                        else
                        {
                            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ⚠️ WARNING: XML file was NOT created!");
                            DebugLogger.Warning($"⚠️ WARNING: Filter XML file was NOT created at: {filePath}");
                        }
                    }
                    catch (Exception saveEx)
                    {
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ❌ XML SAVE FAILED: {saveEx.Message}");
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ❌ Exception Type: {saveEx.GetType().Name}");
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ❌ Stack Trace: {saveEx.StackTrace}");
                        if (saveEx.InnerException != null)
                            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ❌ Inner Exception: {saveEx.InnerException.Message}");
                        DebugLogger.Error($"❌ Failed to save filter '{selectedFilterName}' to '{filePath}': {saveEx.Message}");
                        DebugLogger.Error($"❌ Exception type: {saveEx.GetType().Name}");
                        DebugLogger.Error($"❌ Stack trace: {saveEx.StackTrace}");
                        if (saveEx.InnerException != null)
                        {
                            DebugLogger.Error($"❌ Inner exception: {saveEx.InnerException.Message}");
                        }
                        throw;
                    }
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ========== END STEP 1: XML SAVE ==========");
                    
                    // ✅ STEP 2: Save UI state to database (SelectedHostCategories + OpeningSettings)
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ========== STEP 2: SAVING TO DATABASE ==========");
                    var category = _filterManagementService.GetDisplayCategory(currentFilter);
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Category: '{category}'");
                    DebugLogger.Info($"[SAVE-DEBUG] Category: '{category}'");
                    
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Calling RegisterFilterInDatabase...");
                    _filterManagementService.RegisterFilterInDatabase(selectedFilterName, category);
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ✅ RegisterFilterInDatabase completed");
                    DebugLogger.Info($"Registered filter '{selectedFilterName}' in database for category '{category}'");
                    
                    // ✅ CRITICAL: Persist UI state to DB (linked files, host categories, opening settings)
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ========== STEP 2A: SAVING UI STATE TO DATABASE ==========");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Filter: '{selectedFilterName}'");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Category: '{category}'");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] HostCategories to save: [{string.Join(", ", currentFilter.SelectedHostCategories ?? new List<string>())}]");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] HostCategories Count: {currentFilter.SelectedHostCategories?.Count ?? 0}");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] OpeningSettings: {(currentFilter.OpeningSettings != null ? "EXISTS" : "NULL")}");
                    
                    try
                    {
                        DebugLogger.Info($"[SAVE-DEBUG] SaveFilterUIState: Filter='{selectedFilterName}', Category='{category}', HostCats={currentFilter.SelectedHostCategories?.Count ?? 0}, OpenSettings={(currentFilter.OpeningSettings != null ? "Yes" : "NULL")}");
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Calling UseFilterRepository -> SaveFilterUIState...");
                        _filterManagementService.UseFilterRepository(repo =>
                        {
                            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Inside UseFilterRepository callback - calling SaveFilterUIState");
                            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] SelectedMepCategoryNames to save: [{string.Join(", ", currentFilter.SelectedMepCategoryNames ?? new List<string>())}]");
                            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] SelectedReferenceFiles to save: [{string.Join(", ", currentFilter.SelectedReferenceFiles ?? new List<string>())}]");
                            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] SelectedHostFiles to save: [{string.Join(", ", currentFilter.SelectedHostFiles ?? new List<string>())}]");
                            repo.SaveFilterUIState(
                                selectedFilterName,
                                category,
                                currentFilter.SelectedHostCategories ?? new List<string>(),
                                currentFilter.OpeningSettings,
                                currentFilter.SelectedMepCategoryNames,
                                currentFilter.SelectedReferenceFiles,
                                currentFilter.SelectedHostFiles
                            );
                            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ✅ SaveFilterUIState call completed");
                        });
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ✅ UseFilterRepository completed");
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ✅ UI STATE SAVED TO DATABASE SUCCESSFULLY");
                        DebugLogger.Info($"✅ Saved UI state to DB for filter '{selectedFilterName}' (HostCategories: {currentFilter.SelectedHostCategories?.Count ?? 0})");
                    }
                    catch (Exception uiStateEx)
                    {
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ❌ DATABASE SAVE FAILED: {uiStateEx.Message}");
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ❌ Exception Type: {uiStateEx.GetType().Name}");
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ❌ Stack Trace: {uiStateEx.StackTrace}");
                        if (uiStateEx.InnerException != null)
                            logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ❌ Inner Exception: {uiStateEx.InnerException.Message}");
                        DebugLogger.Error($"⚠️ Failed to save UI state to DB: {uiStateEx.Message}");
                    }
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ========== END STEP 2A: UI STATE DATABASE SAVE ==========");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ========== END STEP 2: DATABASE SAVE ==========");
                    
                    // ✅ STEP 3: Update in-memory filter object
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ========== STEP 3: UPDATE IN-MEMORY FILTER ==========");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Calling UpdateFilterInMemory...");
                    _filterManagementService.UpdateFilterInMemory(selectedFilterName, currentFilter);
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ✅ UpdateFilterInMemory completed");
                    DebugLogger.Info($"Updated in-memory filter object for '{selectedFilterName}' with saved UI state");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ========== END STEP 3: IN-MEMORY UPDATE ==========");
                }
                else
                {
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ⚠️ WARNING: filterListBox.SelectedItem is NULL - cannot save");
                }
                
                logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ========== SAVE FILTER WITH UI STATE COMPLETED ==========");
                
                // ✅ AUTO-HIDING PROMPT: Show "Filter is saved" message and hide after 0.7 seconds
                ShowFilterSavedPrompt();
            }
            catch (Exception ex)
            {
                logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ❌ EXCEPTION IN SaveFilterWithUIState: {ex.Message}");
                logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ❌ Exception Type: {ex.GetType().Name}");
                logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ❌ Stack Trace: {ex.StackTrace}");
                if (ex.InnerException != null)
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ❌ Inner Exception: {ex.InnerException.Message}");
                DebugLogger.Error($"SaveFilterWithUIState failed: {ex.Message}");
            }
            finally
            {
                // ✅ Write all logs to file
                try
                {
                    System.IO.File.AppendAllText(filterDebugLogPath, logBuilder.ToString());
                    DebugLogger.Info($"[FILTER_UI_DEBUG] ✅ Successfully wrote log to: {filterDebugLogPath}");
                    System.Diagnostics.Debug.WriteLine($"[FILTER_UI_DEBUG] ✅ Successfully wrote log to: {filterDebugLogPath}");
                }
                catch (Exception logEx)
                {
                    DebugLogger.Warning($"Failed to write to filter debug log: {logEx.Message}");
                    DebugLogger.Warning($"Log file path was: {filterDebugLogPath}");
                    System.Diagnostics.Debug.WriteLine($"[FILTER_UI_DEBUG] ❌ FAILED to write log: {logEx.Message}");
                    System.Diagnostics.Debug.WriteLine($"[FILTER_UI_DEBUG] Log file path was: {filterDebugLogPath}");
                }
            }
        }
        /// <summary>
        /// Shows an auto-hiding "Filter is saved" prompt that disappears after 0.7 seconds
        /// ✅ FIXED: Prompt now appears as a top-level window to ensure visibility
        /// </summary>
        private void ShowFilterSavedPrompt()
        {
            try
            {
                // ✅ FIXED: Create a top-level form instead of a label inside the dialog
                // This ensures the prompt appears on top of all windows, including the main dialog
                var promptForm = new WinForms.Form
                {
                    Text = "",
                    FormBorderStyle = WinForms.FormBorderStyle.None,
                    StartPosition = WinForms.FormStartPosition.Manual,
                    Size = new System.Drawing.Size(150, 40),
                    BackColor = System.Drawing.Color.White,
                    TopMost = true, // ✅ CRITICAL: Keep on top of all windows
                    ShowInTaskbar = false,
                    ControlBox = false,
                    MaximizeBox = false,
                    MinimizeBox = false
                };
                
                // Create the label inside the form
                var promptLabel = new WinForms.Label
                {
                    Text = "Filter is saved",
                    Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold),
                    ForeColor = System.Drawing.Color.Green,
                    BackColor = System.Drawing.Color.White,
                    Dock = WinForms.DockStyle.Fill,
                    TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                    Padding = new WinForms.Padding(5, 3, 5, 3)
                };
                
                promptForm.Controls.Add(promptLabel);
                
                // Position the form at the center of the main dialog (or screen if dialog not available)
                if (this.WindowState != WinForms.FormWindowState.Minimized && this.Visible)
                {
                    // Position relative to main dialog center
                    promptForm.Location = new System.Drawing.Point(
                        this.Left + (this.Width / 2) - 75,
                        this.Top + (this.Height / 2) - 20
                    );
                }
                else
                {
                    // Fallback: Center on screen
                    var screen = WinForms.Screen.PrimaryScreen.WorkingArea;
                    promptForm.Location = new System.Drawing.Point(
                        (screen.Width - promptForm.Width) / 2,
                        (screen.Height - promptForm.Height) / 2
                    );
                }
                
                // Bring main dialog to front first, then show prompt
                this.Activate();
                this.BringToFront();
                promptForm.Show();
                promptForm.BringToFront();
                    
                    // Create a Timer to hide the prompt after 0.7 seconds
                    var timer = new System.Windows.Forms.Timer
                    {
                        Interval = 700 // 0.7 seconds
                    };
                    
                    timer.Tick += (s, e) =>
                    {
                        timer.Stop();
                        timer.Dispose();
                        
                    if (promptForm != null && !promptForm.IsDisposed)
                        {
                        promptForm.Close();
                        promptForm.Dispose();
                        }
                    };
                    
                    timer.Start();
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"ShowFilterSavedPrompt failed: {ex.Message}");
                // Fallback: Update status label directly and bring dialog to front
                try
                {
                    this.Activate();
                    this.BringToFront();
                _statusLabel.Text = "Filter is saved";
                    _statusLabel.ForeColor = System.Drawing.Color.Green;
                }
                catch { }
            }
        }

        private void RestoreFilterStateToUI(ListBox filterListBox)
        {
            try
            {
                DebugLogger.Info("RestoreFilterStateToUI called");
                
                if (filterListBox.SelectedItem != null)
                {
                    var selectedFilterName = filterListBox.SelectedItem.ToString();
                    var loadedFilter = _filterManagementService.LoadFilterAuto(selectedFilterName);
                    
                    if (loadedFilter != null)
                    {
                        // Restore UI state from the loaded filter
                        RestoreUIStateFromFilter(loadedFilter);
                        DebugLogger.Info($"Restored UI state from filter: {selectedFilterName}");
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"RestoreFilterStateToUI failed: {ex.Message}");
            }
        }

        private void UpdateFilterWithCurrentUIState(OpeningFilter? filter)
        {
            // ✅ DEBUG LOG: Add to filter UI state debug log
            string filterDebugLogPath = SafeFileLogger.GetLogFilePath("filter_ui_state_debug.log");
            var logBuilder = new System.Text.StringBuilder();
            
            try
            {
                logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ========== UpdateFilterWithCurrentUIState STARTED ==========");
                DebugLogger.Info("UpdateFilterWithCurrentUIState called");
                
                // Update filter with current UI selections
                if (filter != null)
                {
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Filter Name: '{filter.Name}'");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Filter is null: {filter == null}");
                    
                    // Get current selected MEP categories
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Getting SelectedMepCategories...");
                    var selectedMepCategories = GetSelectedMepCategories();
                    filter.SelectedMepCategoryNames = selectedMepCategories;
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] SelectedMepCategories Count: {selectedMepCategories?.Count ?? 0}");
                    if (selectedMepCategories != null && selectedMepCategories.Count > 0)
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] SelectedMepCategories: [{string.Join(", ", selectedMepCategories)}]");
                    DebugLogger.Info($"Updated filter with {selectedMepCategories?.Count ?? 0} MEP categories");
                    
                    // Get current selected reference files
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Getting SelectedReferenceFiles...");
                    var selectedReferenceFiles = GetSelectedReferenceFiles();
                    filter.SelectedReferenceFiles = selectedReferenceFiles;
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] SelectedReferenceFiles Count: {selectedReferenceFiles?.Count ?? 0}");
                    DebugLogger.Info($"Updated filter with {selectedReferenceFiles?.Count ?? 0} reference files");
                    
                    // Get current selected host files
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Getting SelectedHostFiles...");
                    var selectedHostFiles = GetSelectedHostFiles();
                    filter.SelectedHostFiles = selectedHostFiles;
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] SelectedHostFiles Count: {selectedHostFiles?.Count ?? 0}");
                    DebugLogger.Info($"Updated filter with {selectedHostFiles?.Count ?? 0} host files");
                    
                    // ✅ FIX: Get current selected host categories (host categories) and save to filter
                    // This includes BOTH horizontal (Walls, Structural Framing) AND vertical (Floors, Ceilings) categories
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Getting SelectedHostCategories from FilterUiStateProvider...");
                    var selectedHostCategories = FilterUiStateProvider.GetSelectedHostCategories?.Invoke() ?? new List<string>();
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] FilterUiStateProvider.GetSelectedHostCategories returned: {selectedHostCategories.Count} categories");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] SelectedHostCategories BEFORE assignment: [{string.Join(", ", selectedHostCategories)}]");
                    
                    filter.SelectedHostCategories = selectedHostCategories;
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ✅ Assigned SelectedHostCategories to filter");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Filter.SelectedHostCategories AFTER assignment: [{string.Join(", ", filter.SelectedHostCategories ?? new List<string>())}]");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Filter.SelectedHostCategories Count AFTER assignment: {filter.SelectedHostCategories?.Count ?? 0}");
                    
                    // ✅ ENHANCED LOGGING: Show breakdown of horizontal vs vertical categories
                    var horizontal = selectedHostCategories.Where(h => h.Equals("Walls", StringComparison.OrdinalIgnoreCase) || 
                                                                        h.Equals("Structural Framing", StringComparison.OrdinalIgnoreCase)).ToList();
                    var vertical = selectedHostCategories.Where(v => v.Equals("Floors", StringComparison.OrdinalIgnoreCase) || 
                                                                       v.Equals("Ceilings", StringComparison.OrdinalIgnoreCase)).ToList();
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Horizontal categories: {horizontal.Count} - [{string.Join(", ", horizontal)}]");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Vertical categories: {vertical.Count} - [{string.Join(", ", vertical)}]");
                    DebugLogger.Info($"Updated filter with {selectedHostCategories?.Count ?? 0} host categories:");
                    DebugLogger.Info($"  Horizontal: {horizontal.Count} ({string.Join(", ", horizontal)})");
                    DebugLogger.Info($"  Vertical: {vertical.Count} ({string.Join(", ", vertical)})");
                    DebugLogger.Info($"  All: {string.Join(", ", selectedHostCategories)}");
                    
                    // Get current clearance settings and store in OpeningSettings
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Getting OpeningSettings from UI...");
                    var clearanceSettings = GetClearanceSettings();
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ClearanceSettings retrieved: {(clearanceSettings != null ? "EXISTS" : "NULL")}");
                    
                    if (filter.OpeningSettings == null)
                    {
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] OpeningSettings is NULL - Creating new OpeningSettings object");
                        filter.OpeningSettings = new OpeningSettings();
                    }
                    else
                    {
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] OpeningSettings EXISTS - Updating existing object");
                    }
                    
                    filter.OpeningSettings.ClearanceSettings = clearanceSettings ?? new Dictionary<string, double>();
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ✅ Assigned ClearanceSettings to filter.OpeningSettings");
                    
                    // ✅ SAVE: "Adopt to modified document" checkbox value to OpeningSettings
                    if (_enableThreePointValidationCheckBox != null)
                    {
                        filter.OpeningSettings.AdoptToDocument = _enableThreePointValidationCheckBox.Checked;
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ✅ Assigned AdoptToDocument = {_enableThreePointValidationCheckBox.Checked} to filter.OpeningSettings");
                        DebugLogger.Info($"Updated filter with AdoptToDocument = {_enableThreePointValidationCheckBox.Checked}");
                    }
                    else
                    {
                        logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ⚠️ WARNING: _enableThreePointValidationCheckBox is NULL - cannot save AdoptToDocument");
                    }
                    
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Filter.OpeningSettings after update: {(filter.OpeningSettings != null ? "EXISTS" : "NULL")}");
                    DebugLogger.Info($"Updated filter with clearance settings");
                    
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ========== UpdateFilterWithCurrentUIState COMPLETED ==========");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Final Filter.SelectedHostCategories Count: {filter.SelectedHostCategories?.Count ?? 0}");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Final Filter.SelectedHostCategories: [{string.Join(", ", filter.SelectedHostCategories ?? new List<string>())}]");
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Final Filter.OpeningSettings: {(filter.OpeningSettings != null ? "EXISTS" : "NULL")}");
                }
                else
                {
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ⚠️ WARNING: Filter is NULL - cannot update UI state");
                }
            }
            catch (Exception ex)
            {
                logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ❌ EXCEPTION IN UpdateFilterWithCurrentUIState: {ex.Message}");
                logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ❌ Exception Type: {ex.GetType().Name}");
                logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ❌ Stack Trace: {ex.StackTrace}");
                if (ex.InnerException != null)
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ❌ Inner Exception: {ex.InnerException.Message}");
                DebugLogger.Error($"UpdateFilterWithCurrentUIState failed: {ex.Message}");
            }
            finally
            {
                // ✅ Write all logs to file
                try
                {
                    System.IO.File.AppendAllText(filterDebugLogPath, logBuilder.ToString());
                    DebugLogger.Info($"[FILTER_UI_DEBUG] ✅ Successfully wrote UpdateFilterWithCurrentUIState log to: {filterDebugLogPath}");
                }
                catch (Exception logEx)
                {
                    DebugLogger.Warning($"Failed to write to filter debug log: {logEx.Message}");
                    DebugLogger.Warning($"Log file path was: {filterDebugLogPath}");
                }
            }
        }

        private void RestoreUIStateFromFilter(OpeningFilter filter)
        {
            try
            {
                DebugLogger.Info("RestoreUIStateFromFilter called");
                DebugLogger.Info($"[RESTORE-DEBUG] Filter is null: {filter == null}");
                if (filter != null)
                {
                    DebugLogger.Info($"[RESTORE-DEBUG] Filter.SelectedMepCategoryNames: {(filter.SelectedMepCategoryNames == null ? "NULL" : filter.SelectedMepCategoryNames.Count + " items: [" + string.Join(", ", filter.SelectedMepCategoryNames) + "]")}");
                    DebugLogger.Info($"[RESTORE-DEBUG] Filter.SelectedReferenceFiles: {(filter.SelectedReferenceFiles == null ? "NULL" : filter.SelectedReferenceFiles.Count + " items: [" + string.Join(", ", filter.SelectedReferenceFiles) + "]")}");
                    DebugLogger.Info($"[RESTORE-DEBUG] Filter.SelectedHostFiles: {(filter.SelectedHostFiles == null ? "NULL" : filter.SelectedHostFiles.Count + " items: [" + string.Join(", ", filter.SelectedHostFiles) + "]")}");
                    DebugLogger.Info($"[RESTORE-DEBUG] Filter.SelectedHostCategories: {(filter.SelectedHostCategories == null ? "NULL" : filter.SelectedHostCategories.Count + " items: [" + string.Join(", ", filter.SelectedHostCategories) + "]")}");
                }
                
                // ✅ CRITICAL FIX: DO NOT check _userHasMadeManualChanges here!
                // That flag is for preventing restoration DURING UI editing, but we should ALWAYS restore from filter.
                // When user selects a filter explicitly, we WANT to restore the saved UI state.
                // The flag prevents restoration only if user is actively making changes within the restoration process,
                // not between operations.

                if (filter != null)
                {
                    // Restore MEP category selections
                    if (filter.SelectedMepCategoryNames != null && filter.SelectedMepCategoryNames.Any())
                    {
                        DebugLogger.Info($"[RESTORE-DEBUG] Restoring MEP categories: {string.Join(", ", filter.SelectedMepCategoryNames)}");
                        RestoreMepCategorySelections(filter.SelectedMepCategoryNames);
                        DebugLogger.Info($"Restored {filter.SelectedMepCategoryNames.Count} MEP category selections");
                    }
                    else
                    {
                        DebugLogger.Info($"[RESTORE-DEBUG] ⚠️ NO MEP categories to restore");
                    }
                    
                    // Restore reference file selections
                    if (filter.SelectedReferenceFiles != null && filter.SelectedReferenceFiles.Any())
                    {
                        DebugLogger.Info($"[RESTORE-DEBUG] Restoring reference files: {string.Join(", ", filter.SelectedReferenceFiles)}");
                        RestoreReferenceFileSelections(filter.SelectedReferenceFiles);
                        DebugLogger.Info($"Restored {filter.SelectedReferenceFiles.Count} reference file selections");
                    }
                    else
                    {
                        DebugLogger.Info($"[RESTORE-DEBUG] ⚠️ NO reference files to restore");
                    }
                    
                    // Restore host file selections
                    if (filter.SelectedHostFiles != null && filter.SelectedHostFiles.Any())
                    {
                        DebugLogger.Info($"[RESTORE-DEBUG] Restoring host files: {string.Join(", ", filter.SelectedHostFiles)}");
                        RestoreHostFileSelections(filter.SelectedHostFiles);
                        DebugLogger.Info($"Restored {filter.SelectedHostFiles.Count} host file selections");
                    }
                    else
                    {
                        DebugLogger.Info($"[RESTORE-DEBUG] ⚠️ NO host files to restore");
                    }
                    
                    // ✅ FIX: Restore host categories (host categories) selections
                    // This includes BOTH horizontal (Walls, Structural Framing) AND vertical (Floors, Ceilings) categories
                    if (filter.SelectedHostCategories != null && filter.SelectedHostCategories.Any())
                    {
                        // ✅ ENHANCED LOGGING: Show breakdown before restoration
                        var horizontal = filter.SelectedHostCategories.Where(h => h.Equals("Walls", StringComparison.OrdinalIgnoreCase) || 
                                                                                   h.Equals("Structural Framing", StringComparison.OrdinalIgnoreCase)).ToList();
                        var vertical = filter.SelectedHostCategories.Where(v => v.Equals("Floors", StringComparison.OrdinalIgnoreCase) || 
                                                                                  v.Equals("Ceilings", StringComparison.OrdinalIgnoreCase)).ToList();
                        DebugLogger.Info($"Restoring {filter.SelectedHostCategories.Count} host element type selections:");
                        DebugLogger.Info($"  Horizontal: {horizontal.Count} ({string.Join(", ", horizontal)})");
                        DebugLogger.Info($"  Vertical: {vertical.Count} ({string.Join(", ", vertical)})");
                        
                        RestoreHostCategorySelections(filter.SelectedHostCategories);
                        DebugLogger.Info($"✅ Restored all host categories: {string.Join(", ", filter.SelectedHostCategories)}");
                    }
                    else
                    {
                        DebugLogger.Info($"[RESTORE-DEBUG] ⚠️ NO host categories to restore");
                    }
                    
                    // Restore clearance settings
                    if (filter.OpeningSettings?.ClearanceSettings != null && filter.OpeningSettings.ClearanceSettings.Any())
                    {
                        RestoreClearanceSettings(filter.OpeningSettings.ClearanceSettings);
                        DebugLogger.Info("Restored clearance settings");
                    }
                    
                    // ✅ RESTORE: "Adopt to modified document" checkbox value from OpeningSettings
                    if (filter.OpeningSettings != null && _enableThreePointValidationCheckBox != null)
                    {
                        _enableThreePointValidationCheckBox.Checked = filter.OpeningSettings.AdoptToDocument;
                        DebugLogger.Info($"Restored AdoptToDocument = {filter.OpeningSettings.AdoptToDocument}");
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"RestoreUIStateFromFilter failed: {ex.Message}");
            }
        }


        private bool IsFileAvailableForReference(LinkedFileType fileType)
        {
            try
            {
                // Check if file type is suitable for reference
                return fileType == LinkedFileType.Electrical || 
                       fileType == LinkedFileType.Mechanical || 
                       fileType == LinkedFileType.Plumbing ||
                       fileType == LinkedFileType.FireProtection;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"IsFileAvailableForReference failed: {ex.Message}");
                return false;
            }
        }

        private bool IsFileAvailableForHost(LinkedFileType fileType)
        {
            try
            {
                // Only show structural files (AR, ARC, ST, STR) in the main "Host Elements" section
                // MEP files (Electrical, Mechanical, Plumbing, FireProtection) should go to "Other Files" section
                return Services.LinkedFileDetectionService.IsHostOpeningFile(fileType);
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"IsFileAvailableForHost failed: {ex.Message}");
                return false;
            }
        }

        private Services.MepCategory GetSelectedMepCategory()
        {
            try
            {
                // Get currently selected MEP category from UI
                // This is a simplified implementation - should be enhanced based on actual UI
                return Services.MepCategory.Ducts; // Default fallback
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"GetSelectedMepCategory failed: {ex.Message}");
                return Services.MepCategory.Ducts; // Default fallback
            }
        }

        private void UpdateParameterServiceDropdowns(Dictionary<string, List<Models.ParameterInfo>> categoryParameters)
        {
            try
            {
                DebugLogger.Info($"UpdateParameterServiceDropdowns called with {categoryParameters.Count} categories");

                // Use the working PopulateParameterDropdowns method instead of placeholder
                PopulateParameterDropdowns();

                DebugLogger.Info("[PARAMETER_SERVICE] UpdateParameterServiceDropdowns completed successfully");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"UpdateParameterServiceDropdowns failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Transfers parameter value from opening parameter dropdown to MEP parameter dropdown
        /// This implements the parameter transfer functionality that matches the dropdowns in parameter service
        /// </summary>
        private void TransferParameterValue(WinForms.ComboBox mepParamCombo, WinForms.ComboBox openingParamCombo)
        {
            try
            {
                DebugLogger.Info("[PARAMETER_TRANSFER] Starting parameter value transfer");

                // Get the selected opening parameter value
                var selectedOpeningParam = openingParamCombo?.SelectedItem?.ToString();
                if (string.IsNullOrEmpty(selectedOpeningParam) || selectedOpeningParam == "<Select Opening Parameter>")
                {
                    DebugLogger.Warning("[PARAMETER_TRANSFER] No valid opening parameter selected");
                    _statusLabel.Text = "Please select an opening parameter to transfer";
                    return;
                }

                // Check if the MEP parameter dropdown contains this parameter
                if (mepParamCombo?.Items.Contains(selectedOpeningParam) == true)
                {
                    // Set the MEP parameter dropdown to the selected opening parameter
                    mepParamCombo.SelectedItem = selectedOpeningParam;
                    DebugLogger.Info($"[PARAMETER_TRANSFER] Successfully transferred parameter: '{selectedOpeningParam}' from opening to MEP dropdown");
                    _statusLabel.Text = $"Parameter '{selectedOpeningParam}' transferred to MEP parameter";
                }
                else
                {
                    // Parameter not found in MEP dropdown - add it if possible
                    if (mepParamCombo != null && !mepParamCombo.Items.Contains(selectedOpeningParam))
                    {
                        mepParamCombo.Items.Add(selectedOpeningParam);
                        mepParamCombo.SelectedItem = selectedOpeningParam;
                        DebugLogger.Info($"[PARAMETER_TRANSFER] Added and selected parameter: '{selectedOpeningParam}' in MEP dropdown");
                        _statusLabel.Text = $"Parameter '{selectedOpeningParam}' added and transferred to MEP parameter";
                    }
                    else
                    {
                        DebugLogger.Warning($"[PARAMETER_TRANSFER] Parameter '{selectedOpeningParam}' not found in MEP parameter list");
                        _statusLabel.Text = $"Parameter '{selectedOpeningParam}' not available in MEP parameters";
                    }
                }

                DebugLogger.Info("[PARAMETER_TRANSFER] Parameter transfer completed");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[PARAMETER_TRANSFER] Error during parameter transfer: {ex.Message}");
                _statusLabel.Text = $"Error transferring parameter: {ex.Message}";
            }
        }
    }
}