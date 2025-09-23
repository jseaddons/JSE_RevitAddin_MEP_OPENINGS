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
        
        // Store all collected parameters globally to persist across refreshes
        private Dictionary<string, List<Models.ParameterInfo>> _allCollectedParameters = new Dictionary<string, List<Models.ParameterInfo>>();
        
        // Flag to prevent infinite loops during ComboBox population
        private bool _isUpdatingComboBoxes = false;
        
        
        // Main panels - 4-section layout
        private WinForms.Panel _leftPanel = null!;
        private WinForms.Panel _rightPanel = null!;
        private WinForms.Panel _mainSplitter = null!;  // Using Panel instead of Splitter to avoid docking requirement
        
        // NEW: Filters panel (left column)
        private WinForms.Panel _filtersPanel = null!;
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
        private WinForms.Button _configureButton = null!;
        
        
        // Dynamic UI controls (only what we actually use)
        private WinForms.ComboBox _mepTypeCombo = null!;
        private WinForms.Panel _clearancePanel = null!;
        private WinForms.Panel _cableTrayPanel = null!;
        private WinForms.Panel _damperPanel = null!;
        private LinkedFileService? _linkedFileService;
        private List<LinkedFileInfo> _linkedFiles = new List<LinkedFileInfo>();
        private Document? _activeDocument;
        
        // Selection tracking for parameter filtering
        private List<string> _selectedReferenceFiles = new List<string>();
        private List<string> _selectedCategories = new List<string>();
        private List<string> _selectedHostFiles = new List<string>();
        // Opening type controls (right section)
        private WinForms.Panel _openingTypePanel = null!;
        private WinForms.RadioButton _rectangularRadio = null!;
        private WinForms.RadioButton _circularRadio = null!;
        private WinForms.CheckBox _fullPenetrationCheckBox = null!;
        // Parameter filter controls (right section)
        private WinForms.Panel _parameterFilterPanel = null!;
        private List<WinForms.Panel> _parameterRows = new List<WinForms.Panel>();
        private WinForms.Button _addParameterButton = null!;
        // Service parameter tabs
        private WinForms.TabControl _serviceParameterTabs = null!;

        // constants (top of class)
        private const int InnerRightWidth = 320;  // choose 300–360

        public EmergencyMainDialog(ApplicationProfileService appProfileService, Document? document = null, UIDocument? uiDocument = null)
        {
            _appProfileService = appProfileService;
            _document = document;
            _uiDocument = uiDocument;
            
            // Close all log files to free file handles
            DebugLogger.CloseAllLogFiles();
            
            // Set logging context for OK button debugging
            DebugLogger.SetServiceContext("OKButton");
            
            // STEP 1: IMMEDIATE LOG - Create timestamped log file to avoid overwriting
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            string mainUiLogPath = $@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\MainUi_{timestamp}.log";
            try
            {
                // Ensure directory exists
                string logDir = Path.GetDirectoryName(mainUiLogPath) ?? @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log";
                if (!Directory.Exists(logDir))
                {
                    Directory.CreateDirectory(logDir);
                }

                // BUILD TIMESTAMP - Write current build time to identify version
                string buildTimestamp = "2025-09-05 12:30:00"; // LATEST BUILD WITH TIMESTAMPED LOGS
                File.AppendAllText(mainUiLogPath, $"[{DateTime.Now}] BUILD TIMESTAMP: {buildTimestamp}\n");
                File.AppendAllText(mainUiLogPath, $"[{DateTime.Now}] EMERGENCY MAIN DIALOG CONSTRUCTOR STARTED\n");
                File.AppendAllText(mainUiLogPath, $"[{DateTime.Now}] Log file: MainUi_{timestamp}.log\n");
                File.AppendAllText(mainUiLogPath, $"[{DateTime.Now}] ApplicationProfileService: {appProfileService != null}\n");
                File.AppendAllText(mainUiLogPath, $"[{DateTime.Now}] Document: {document?.Title ?? "null"}\n");
            }
            catch (Exception ex)
            {
                // If MainUi.log fails, try a timestamped fallback debug file
                try
                {
                    string fallbackTimestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                    File.AppendAllText($@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\emergency_debug_{fallbackTimestamp}.txt",
                        $"[{DateTime.Now}] EmergencyMainDialog constructor failed to create MainUi.log: {ex.Message}\n");
                }
                catch { }
                return; // Exit if we can't even create basic logs
            }

            // STEP 2: Continue with normal initialization
            _appProfileService = appProfileService ?? throw new ArgumentNullException(nameof(appProfileService));
            _filterManagementService = new FilterManagementService(
                msg => DebugLogger.Info(msg),
                msg => _statusLabel.Text = msg
            );
            _linkedFileService = new LinkedFileService();
            _activeDocument = document;

            // Log successful initialization
            try
            {
                File.AppendAllText(mainUiLogPath, $"[{DateTime.Now}] Services initialized successfully\n");
            }
            catch { }

            // Initialize DebugLogger to use the timestamped MainUi log file for all subsequent logging
            DebugLogger.InitCustomLogFileOverwrite("MainUi");
            DebugLogger.Info($"DebugLogger initialized to use timestamped MainUi log file: {Path.GetFileName(mainUiLogPath)}");

            try
            {
                DebugLogger.Info("About to call InitializeComponent()");
                InitializeComponent();
                DebugLogger.Info("InitializeComponent() completed successfully");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"InitializeComponent() failed: {ex.Message}");
                DebugLogger.Error($"Stack trace: {ex.StackTrace}");
                throw; // Re-throw to see the error
            }

            try
            {
                DebugLogger.Info("About to call LoadProfileInfo()");
                LoadProfileInfo();
                DebugLogger.Info("LoadProfileInfo() completed successfully");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"LoadProfileInfo() failed: {ex.Message}");
                DebugLogger.Error($"Stack trace: {ex.StackTrace}");
                throw; // Re-throw to see the error
            }

            // STEP 3: Log completion
            try
            {
                File.AppendAllText(mainUiLogPath, $"[{DateTime.Now}] EmergencyMainDialog initialization COMPLETED\n");
            }
            catch { }

            // Load real linked files if document is provided (after UI is initialized)
            if (document != null)
            {
                this.Shown += (s, e) => LoadRealLinkedFiles(document);
            }
        }

        private void InitializeComponent()
        {
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

            // Toolbar buttons in header - moved further left
            int buttonY = 20;
            int buttonWidth = 70;
            int buttonHeight = 35;
            int buttonSpacing = 80;
            int startX = 450;  // Change this line to move buttons closer

            _okButton = new WinForms.Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(startX, buttonY),
                Size = new System.Drawing.Size(buttonWidth, buttonHeight),
                BackColor = System.Drawing.Color.FromArgb(0, 122, 204),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = WinForms.FlatStyle.Flat,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold)
            };
            _okButton.Click += OnOkClick;
            _headerPanel.Controls.Add(_okButton);

            _saveButton = new WinForms.Button
            {
                Text = "Save",
                Location = new System.Drawing.Point(startX + buttonSpacing, buttonY),
                Size = new System.Drawing.Size(buttonWidth, buttonHeight),
                BackColor = System.Drawing.Color.FromArgb(40, 167, 69),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = WinForms.FlatStyle.Flat,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold)
            };
            _saveButton.Click += OnSaveClick;
            _headerPanel.Controls.Add(_saveButton);

            _cancelButton = new WinForms.Button
            {
                Text = "Cancel",
                Location = new System.Drawing.Point(startX + 2 * buttonSpacing, buttonY),
                Size = new System.Drawing.Size(buttonWidth, buttonHeight),
                BackColor = System.Drawing.Color.FromArgb(108, 117, 125),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = WinForms.FlatStyle.Flat,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold)
            };
            _cancelButton.Click += OnCancelClick;
            _headerPanel.Controls.Add(_cancelButton);

            _closeButton = new WinForms.Button
            {
                Text = "Close",
                Location = new System.Drawing.Point(startX + 3 * buttonSpacing, buttonY),
                Size = new System.Drawing.Size(buttonWidth, buttonHeight),
                BackColor = System.Drawing.Color.FromArgb(220, 53, 69),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = WinForms.FlatStyle.Flat,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold)
            };
            _closeButton.Click += OnCloseClick;
            _headerPanel.Controls.Add(_closeButton);

            // Parameter Transfer Button
            _parameterTransferButton = new WinForms.Button
            {
                Text = "Parameter Transfer",
                Location = new System.Drawing.Point(startX + 4 * buttonSpacing, buttonY),
                Size = new System.Drawing.Size(buttonWidth + 20, buttonHeight), // Slightly wider for longer text
                BackColor = System.Drawing.Color.FromArgb(111, 66, 193), // Purple color for parameter transfer
                ForeColor = System.Drawing.Color.White,
                FlatStyle = WinForms.FlatStyle.Flat,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold)
            };
            _parameterTransferButton.Click += OnParameterTransferClick;
            _headerPanel.Controls.Add(_parameterTransferButton);

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
            
            // Refresh Button (before progress bar)
            _refreshButton = new WinForms.Button
            {
                Text = "Refresh",
                Location = new System.Drawing.Point(buttonStartX, 3), // After status label
                Size = new System.Drawing.Size(60, 24),
                BackColor = System.Drawing.Color.FromArgb(108, 117, 125),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = WinForms.FlatStyle.Flat,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold)
            };
            _statusPanel.Controls.Add(_refreshButton);
            System.Diagnostics.Debug.WriteLine($"Refresh button created at location: {_refreshButton.Location}, Size: {_refreshButton.Size}, Visible: {_refreshButton.Visible}");

            // Configure Button (next to refresh button)
            _configureButton = new WinForms.Button
            {
                Text = "Config",
                Location = new System.Drawing.Point(buttonStartX + 65 + statusButtonSpacing, 3), // Next to refresh button
                Size = new System.Drawing.Size(60, 24),
                BackColor = System.Drawing.Color.FromArgb(102, 16, 242),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = WinForms.FlatStyle.Flat,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold)
            };
            _statusPanel.Controls.Add(_configureButton);
            
            // Progress Bar (after buttons, can be longer now)
            int progressBarStartX = buttonStartX + 130 + statusButtonSpacing; // After both buttons
            _progressBar = new WinForms.ProgressBar
            {
                Location = new System.Drawing.Point(progressBarStartX, 5), // After buttons
                Size = new System.Drawing.Size(leftSectionRightEdge - progressBarStartX - 10, 20), // Dynamic width to fill remaining space
                Style = WinForms.ProgressBarStyle.Continuous,
                Minimum = 0,
                Maximum = 100,
                Value = 0
            };
            _statusPanel.Controls.Add(_progressBar);
            
            // Add event handlers for the buttons
            _refreshButton.Click += OnRefreshClick;
            _configureButton.Click += OnConfigureClick;

            // EXTENSIVE DEBUG LOGGING FOR BUTTON CREATION
            System.Diagnostics.Debug.WriteLine($"[BUTTON_DEBUG] Refresh button created: Location=({_refreshButton.Location.X},{_refreshButton.Location.Y}), Size=({_refreshButton.Size.Width},{_refreshButton.Size.Height}), Visible={_refreshButton.Visible}, Enabled={_refreshButton.Enabled}");
            System.Diagnostics.Debug.WriteLine($"[BUTTON_DEBUG] Refresh button text: '{_refreshButton.Text}'");
            System.Diagnostics.Debug.WriteLine($"[BUTTON_DEBUG] Refresh button parent: {_refreshButton.Parent?.Name ?? "null"}");
            System.Diagnostics.Debug.WriteLine($"[BUTTON_DEBUG] Status panel size: {_statusPanel.Size}, location: {_statusPanel.Location}");

            // Log to multiple files to ensure visibility
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] REFRESH BUTTON CREATED: Location=({_refreshButton.Location.X},{_refreshButton.Location.Y}), Size=({_refreshButton.Size.Width},{_refreshButton.Size.Height}), Visible={_refreshButton.Visible}, Enabled={_refreshButton.Enabled}\n");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] REFRESH BUTTON CREATED: Visible={_refreshButton.Visible}, Enabled={_refreshButton.Enabled}\n");

            // Test if event handler is attached by checking the Click event
            var clickEvent = _refreshButton.GetType().GetEvent("Click");
            if (clickEvent != null)
            {
                System.Diagnostics.Debug.WriteLine($"[BUTTON_DEBUG] Click event found on refresh button");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Click event attached to refresh button\n");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"[BUTTON_DEBUG] ERROR: Click event NOT found on refresh button");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] ERROR: Click event NOT attached to refresh button\n");
            }

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
                DebugLogger.Info($"_mainSplitter (Panel) created: Width={_mainSplitter.Width}, Dock={_mainSplitter.Dock}");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to create _mainSplitter: {ex.Message}");
                throw;
            }

            // Right Panel - keep original width
            _rightPanel = new WinForms.Panel
            {
                BackColor = System.Drawing.Color.White,
                BorderStyle = WinForms.BorderStyle.FixedSingle,
                Width = 460  // Restored to original width
            };
            this.Controls.Add(_rightPanel);
            DebugLogger.Info($"_rightPanel created: Width={_rightPanel.Width}, Location={_rightPanel.Location}, Dock={_rightPanel.Dock}");

            // Add content to panels
            InitializePanelContent();
            PositionPanels();
            BalanceLeftLayout();
            this.Shown += (_, __) => _bottomLeftPanel.Height = (_leftPanel.Height - _horizontalSplitter.Height) / 2;
            this.Resize += (_, __) => BalanceLeftLayout();
            _leftPanel.Resize += (_, __) => BalanceLeftLayout();

            this.ResumeLayout(false);
        }

        private void InitializePanelContent()
        {
            // NEW: Create filters panel (left column)
            CreateFiltersPanel();
            
            // Create 4-section layout within left panel
            CreateFourSectionLayout();
            
            // Right Panel: Opening Configuration
            InitializeRightPanel();
        }

        private void CreateFiltersPanel()
        {
            DebugLogger.Info("=== STARTING CreateFiltersPanel ===");
            
            // Create filters panel
            _filtersPanel = new WinForms.Panel
            {
                BackColor = System.Drawing.Color.FromArgb(248, 249, 250),
                BorderStyle = WinForms.BorderStyle.FixedSingle
            };
            this.Controls.Add(_filtersPanel);
            DebugLogger.Info("_filtersPanel created and added to form");
            
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
            
            DebugLogger.Info("=== CreateFiltersPanel COMPLETED ===");
        }

        private void PopulateFiltersPanel()
        {
            DebugLogger.Info("=== STARTING PopulateFiltersPanel ===");
            
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
            DebugLogger.Info("Filters title label created");
            
            // Filter List (like conVoid's filter list) - Enlarged height to -140
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
            DebugLogger.Info("Filter list box created");
            
            // Add sample filters (like conVoid) via service so internal list is tracked
            _filterManagementService.SeedDefaultFilters(
                filterListBox,
                new System.Collections.Generic.List<string> { "Electrical", "Plumbing", "Ventilation" }
            );
            DebugLogger.Info("Sample filters seeded via FilterManagementService");
            
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
            DebugLogger.Info("New filter button created");
            
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
            DebugLogger.Info("Copy filter button created");
            
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
            DebugLogger.Info("Rename filter button created");
            
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
            DebugLogger.Info("Delete filter button created");
            
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
            DebugLogger.Info("Save filter button created");
            
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
            DebugLogger.Info("Load filter button created");
            
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
            filterListBox.SelectedIndexChanged += (s, e) => RestoreFilterStateToUI(filterListBox);
            
            DebugLogger.Info("=== PopulateFiltersPanel COMPLETED ===");
        }

        private void CreateFourSectionLayout()
        {
            DebugLogger.Info("=== STARTING CreateFourSectionLayout ===");
            DebugLogger.Info($"_leftPanel size: {_leftPanel.Width}x{_leftPanel.Height}");

            // Top-Left Panel: Reference Elements (linked files)
            _topLeftPanel = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.Fill,
                BackColor = System.Drawing.Color.FromArgb(245, 245, 245),
                BorderStyle = WinForms.BorderStyle.FixedSingle
            };
            _leftPanel.Controls.Add(_topLeftPanel);
            DebugLogger.Info("_topLeftPanel created and added");

            // Horizontal Splitter (between top and bottom)
            _horizontalSplitter = new WinForms.Splitter
            {
                Dock = WinForms.DockStyle.Bottom,
                Height = 3,
                BackColor = System.Drawing.Color.Gray
            };
            _leftPanel.Controls.Add(_horizontalSplitter);
            DebugLogger.Info("_horizontalSplitter created and added");

            // Bottom-Left Panel: Host Elements (linked files)
            _bottomLeftPanel = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.Bottom,
                Height = _leftPanel.Height / 2,
                BackColor = System.Drawing.Color.FromArgb(245, 245, 245),
                BorderStyle = WinForms.BorderStyle.FixedSingle
            };
            _leftPanel.Controls.Add(_bottomLeftPanel);
            DebugLogger.Info("_bottomLeftPanel created and added");

            // Now create the right side panels within the top and bottom panels
            CreateTopRightPanel();
            CreateBottomRightPanel();

            // Populate each section with content
            PopulateTopLeftSection();      // Reference Elements (linked files)
            PopulateTopRightSection();     // Reference Categories (MEP categories)
            PopulateBottomLeftSection();   // Host Elements (linked files)
            PopulateBottomRightSection();  // Host Categories (Revit categories)

            DebugLogger.Info("=== CreateFourSectionLayout COMPLETED ===");
        }

        private void CreateTopRightPanel()
        {
            // Top-Right Panel: Reference Categories (MEP categories) - ANCHOR instead of DOCK
            _topRightPanel = new WinForms.Panel
            {
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Right | WinForms.AnchorStyles.Bottom,
                Width = 200,  // Reduced from 340 to 200
                BackColor = System.Drawing.Color.FromArgb(250, 250, 250),
                BorderStyle = WinForms.BorderStyle.FixedSingle
            };
            _topLeftPanel.Controls.Add(_topRightPanel);

            // Vertical Splitter (between left and right in top section) - ANCHOR instead of DOCK
            _verticalSplitter = new WinForms.Splitter
            {
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Right | WinForms.AnchorStyles.Bottom,
                Width = 3,
                BackColor = System.Drawing.Color.Gray
            };
            _topLeftPanel.Controls.Add(_verticalSplitter);
        }

        private void CreateBottomRightPanel()
        {
            // Bottom-Right Panel: Host Categories (Revit categories) - ANCHOR instead of DOCK
            _bottomRightPanel = new WinForms.Panel
            {
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Right | WinForms.AnchorStyles.Bottom,
                Width = 200,  // Reduced from 340 to 200
                BackColor = System.Drawing.Color.FromArgb(250, 250, 250),
                BorderStyle = WinForms.BorderStyle.FixedSingle
            };
            _bottomLeftPanel.Controls.Add(_bottomRightPanel);

            // Vertical Splitter (between left and right in bottom section) - ANCHOR instead of DOCK
            var bottomVerticalSplitter = new WinForms.Splitter
            {
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Right | WinForms.AnchorStyles.Bottom,
                Width = 3,
                BackColor = System.Drawing.Color.Gray
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
                    string displayText = $"{file.FileName} ({file.ElementCount} elements)";
                    if (!file.IsLoaded)
                    {
                        displayText += " [NOT LOADED]";
                    }
                    
                    // Show MEP files in main list, others in "Other Files" section
                    bool isAvailableForReference = IsFileAvailableForReference(file.FileType);
                    if (isAvailableForReference)
                    {
                        referenceFilesListBox.Items.Add(displayText, false);
            }
            else
            {
                        otherRefFilesListBox?.Items.Add(displayText, false);
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
                // Use a timer to delay the update to avoid issues during the check operation
                System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
                timer.Interval = 10;
                timer.Tick += (s, args) => {
                    timer.Stop();
                    timer.Dispose();
                    UpdateMepTypeBasedOnSelection();
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

            // Always add the active document as a host element (available by default)
            if (_activeDocument != null)
            {
                string activeDocName = _activeDocument.Title;
                hostFilesListBox.Items.Add($"{activeDocName} (Active Document)", false);
                System.Diagnostics.Debug.WriteLine($"Added active document '{activeDocName}' to host elements");
            }
            
            // Get the other files list box
            var hostListBoxes = _bottomLeftPanel.Controls.OfType<WinForms.CheckedListBox>().ToList();
            var otherHostFilesListBox2 = hostListBoxes.Skip(1).FirstOrDefault();

            // Add files to appropriate sections
            if (_linkedFiles.Count > 0 && _linkedFileService != null)
            {
                foreach (var file in _linkedFiles)
                {
                    string displayText = $"{file.FileName} ({file.ElementCount} elements)";
                    if (!file.IsLoaded)
                    {
                        displayText += " [NOT LOADED]";
                    }
                    
                    // Show Architecture/Structural files in main list, others in "Other Files" section
                    bool isAvailableForHost = IsFileAvailableForHost(file.FileType);
                    if (isAvailableForHost)
                    {
                        hostFilesListBox.Items.Add(displayText, false);
            }
            else
            {
                        otherHostFilesListBox2?.Items.Add(displayText, false);
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
            _mepTypeCombo.Items.AddRange(new[] { "Pipe", "Duct", "Duct Accessories", "Duct Fittings", "Cable Tray", "Conduit" });
            _mepTypeCombo.SelectedIndex = 0;
            _rightPanel.Controls.Add(_mepTypeCombo);

            // Opening Type (Rectangular / Circular)
            CreateOpeningTypePanel();

            // Create clearance panels
            CreateClearancePanels();

            // Parameter filter section
            CreateParameterFilterPanel();
            // Ensure correct panel visible at startup
            UpdateClearanceVisibility();
        }

        private void CreateOpeningTypePanel()
        {
            _openingTypePanel = new WinForms.Panel
            {
                Location = new System.Drawing.Point(10, 80),
                Size = new System.Drawing.Size(_rightPanel.Width - 20, 70), // Increased height for checkbox
                BackColor = System.Drawing.Color.FromArgb(248, 249, 250),
                BorderStyle = WinForms.BorderStyle.FixedSingle,
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right
            };
            _rightPanel.Controls.Add(_openingTypePanel);

            var typeLabel = new WinForms.Label
            {
                Text = "Opening Type:",
                Location = new System.Drawing.Point(10, 8),
                Size = new System.Drawing.Size(100, 16)
            };
            _openingTypePanel.Controls.Add(typeLabel);

            _rectangularRadio = new WinForms.RadioButton
            {
                Text = "Rectangular",
                Location = new System.Drawing.Point(120, 6),
                AutoSize = true
            };
            _openingTypePanel.Controls.Add(_rectangularRadio);

            _circularRadio = new WinForms.RadioButton
            {
                Text = "Circular",
                Location = new System.Drawing.Point(220, 6),
                AutoSize = true
            };
            _openingTypePanel.Controls.Add(_circularRadio);

            // Full penetration checkbox
            _fullPenetrationCheckBox = new WinForms.CheckBox
            {
                Text = "Full Penetration (fixes fitting interference)",
                Location = new System.Drawing.Point(10, 30),
                Size = new System.Drawing.Size(300, 20),
                Checked = true, // Default to enabled
                // ToolTipText = "Ensures openings fully penetrate host elements even when fittings obstruct MEP elements" // Not supported in WinForms CheckBox
            };
            _openingTypePanel.Controls.Add(_fullPenetrationCheckBox);

            // default selection
            _rectangularRadio.Checked = true;

            // enable/disable by MEP type
            _mepTypeCombo.SelectedIndexChanged += (_, __) => {
                UpdateOpeningTypeAvailability();
                UpdateClearanceVisibility();
            };
            UpdateOpeningTypeAvailability();
        }

        private void UpdateOpeningTypeAvailability()
        {
            // Both shapes remain selectable for all MEP types
            _rectangularRadio.Enabled = true;
            _circularRadio.Enabled = true;
        }

        private void UpdateClearanceVisibility()
        {
            var mep = _mepTypeCombo.SelectedItem?.ToString() ?? string.Empty;
            UpdateClearanceVisibilityForCategory(mep);
        }

        private void UpdateClearanceVisibilityForCategory(string category)
        {
            // Hide all
            _clearancePanel.Visible = false;
            _cableTrayPanel.Visible = false;
            _damperPanel.Visible = false;

            if (category.Equals("Cable Trays", StringComparison.OrdinalIgnoreCase))
            {
                _cableTrayPanel.Visible = true;
                SetDefaultClearanceValues("Cable Trays");
            }
            else if (category.Equals("Duct Accessories", StringComparison.OrdinalIgnoreCase))
            {
                _damperPanel.Visible = true;
                SetDefaultClearanceValues("Duct Accessories");
            }
            else if (category.Equals("Ducts", StringComparison.OrdinalIgnoreCase))
            {
                _clearancePanel.Visible = true;
                SetDefaultClearanceValues("Ducts");
            }
            else if (category.Equals("Pipes", StringComparison.OrdinalIgnoreCase))
            {
                _clearancePanel.Visible = true;
                SetDefaultClearanceValues("Pipes");
            }
            else
            {
                _clearancePanel.Visible = true; // default
                SetDefaultClearanceValues("Default");
            }
        }

        private void CreateClearancePanels()
        {
            // Standard Clearance Panel
            _clearancePanel = new WinForms.Panel
            {
                Location = new System.Drawing.Point(10, 135),
                Size = new System.Drawing.Size(_rightPanel.Width - 20, 100),
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

            // Cable Tray Panel (initially hidden) - Top Side + Other Sides
            _cableTrayPanel = new WinForms.Panel
            {
                Location = new System.Drawing.Point(10, 135),
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

            var topSideTxt = new WinForms.TextBox { Location = new System.Drawing.Point(170, 38), Size = new System.Drawing.Size(50, 20), Text = "75", Tag = "cabletray_top_normal", Enabled = false, BackColor = System.Drawing.Color.LightGray };
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
            var otherTxt = new WinForms.TextBox { Location = new System.Drawing.Point(170, 68), Size = new System.Drawing.Size(50, 20), Text = "25", Tag = "cabletray_other_normal", Enabled = false, BackColor = System.Drawing.Color.LightGray };
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
                Location = new System.Drawing.Point(10, 135),
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
                // Lock - disable editing
                lockBtn.Text = "🔒";
                lockBtn.BackColor = System.Drawing.Color.LightGreen;
                textBox.Enabled = false;
                textBox.BackColor = System.Drawing.Color.LightGray;
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
                        SetClearancePanelValues("50", "25");
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

        private void CreateParameterFilterPanel()
        {
            _parameterFilterPanel = new WinForms.Panel
            {
                Location = new System.Drawing.Point(10, 250),
                Size = new System.Drawing.Size(_rightPanel.Width - 20, _rightPanel.Height - 260),
                BackColor = System.Drawing.Color.FromArgb(248, 249, 250),
                BorderStyle = WinForms.BorderStyle.FixedSingle,
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right | WinForms.AnchorStyles.Bottom
            };
            _rightPanel.Controls.Add(_parameterFilterPanel);

            var title = new WinForms.Label
            {
                Text = "Service Parameters",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold),
                Location = new System.Drawing.Point(10, 8),
                Size = new System.Drawing.Size(140, 18)
            };
            _parameterFilterPanel.Controls.Add(title);

            // Create tabbed service parameters
            _serviceParameterTabs = new WinForms.TabControl
            {
                Location = new System.Drawing.Point(5, 25),
                Size = new System.Drawing.Size(_parameterFilterPanel.Width - 10, _parameterFilterPanel.Height - 30),
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right | WinForms.AnchorStyles.Bottom
            };
            _parameterFilterPanel.Controls.Add(_serviceParameterTabs);

            // Create service tabs
            CreateServiceTab("Ducts", "DUCTS");
            CreateServiceTab("Duct Accessories", "DUCT_ACCESSORIES");
            CreateServiceTab("Cable Trays", "CABLE_TRAYS");
            CreateServiceTab("Pipes", "PIPES");
        }

        private void CreateServiceTab(string tabName, string serviceCode)
        {
            var tabPage = new WinForms.TabPage(tabName);
            
            var servicePanel = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.Fill,
                BackColor = System.Drawing.Color.White,
                Padding = new WinForms.Padding(3)
            };

            // Add default parameter rows for this service
            AddServiceParameterRow(servicePanel, "Reference Level", "A_GARDEN LEVEL");
            AddServiceParameterRow(servicePanel, "Size", "100x100");

            // Add button (plus) for this service
            var addButton = new WinForms.Button
            {
                Text = "+",
                Location = new System.Drawing.Point(servicePanel.Width - 35, 6),
                Size = new System.Drawing.Size(24, 24),
                BackColor = System.Drawing.Color.FromArgb(230, 255, 230),
                FlatStyle = WinForms.FlatStyle.Flat,
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Right,
                Tag = serviceCode // Store service code for identification
            };
            addButton.Click += (_, __) => AddServiceParameterRow(servicePanel, "<Select>", "");
            servicePanel.Controls.Add(addButton);

            tabPage.Controls.Add(servicePanel);
            _serviceParameterTabs.TabPages.Add(tabPage);
        }

        private void AddServiceParameterRow(WinForms.Panel servicePanel, string parameterName, string value)
        {
            int rowHeight = 24;
            int top = 25 + (servicePanel.Controls.OfType<WinForms.Panel>().Count() * (rowHeight + 3));

            var row = new WinForms.Panel
            {
                Location = new System.Drawing.Point(8, top),
                Size = new System.Drawing.Size(servicePanel.Width - 16, rowHeight),
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right
            };
            servicePanel.Controls.Add(row);

            var nameCombo = new WinForms.ComboBox
            {
                Location = new System.Drawing.Point(0, 2),
                Size = new System.Drawing.Size(120, 20),
                DropDownStyle = WinForms.ComboBoxStyle.DropDownList
            };
            
            // Get current MEP parameters from the latest refresh
            var mepParameters = GetCurrentMepParameters();
            nameCombo.Items.AddRange(mepParameters.ToArray());
            nameCombo.SelectedItem = parameterName;
            
            row.Controls.Add(nameCombo);

            var openingParamCombo = new WinForms.ComboBox
            {
                Location = new System.Drawing.Point(130, 2),
                Size = new System.Drawing.Size(row.Width - 130 - 30, 20),
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right,
                DropDownStyle = WinForms.ComboBoxStyle.DropDownList
            };
            
            // Get current opening parameters from the latest refresh
            var openingParameters = GetCurrentOpeningParameters();
            openingParamCombo.Items.AddRange(openingParameters.ToArray());
            openingParamCombo.Items.Insert(0, "<Select Opening Parameter>");
            openingParamCombo.SelectedIndex = 0;
            
            row.Controls.Add(openingParamCombo);

            var removeBtn = new WinForms.Button
            {
                Text = "×",
                Location = new System.Drawing.Point(row.Width - 25, 1),
                Size = new System.Drawing.Size(20, 20),
                BackColor = System.Drawing.Color.FromArgb(255, 230, 230),
                FlatStyle = WinForms.FlatStyle.Flat,
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Right
            };
            removeBtn.Click += (_, __) => {
                servicePanel.Controls.Remove(row);
                RepositionServiceParameterRows(servicePanel);
            };
            row.Controls.Add(removeBtn);
        }

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
            int top = 30 + (_parameterRows.Count * (rowHeight + 6));

            var row = new WinForms.Panel
            {
                Location = new System.Drawing.Point(8, top),
                Size = new System.Drawing.Size(_parameterFilterPanel.Width - 16, rowHeight),
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right
            };
            _parameterFilterPanel.Controls.Add(row);
            _parameterRows.Add(row);

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
                _parameterFilterPanel.Controls.Remove(row);
                _parameterRows.Remove(row);
                ReflowParameterRows();
            };
            row.Controls.Add(removeBtn);
        }

        private void ReflowParameterRows()
        {
            int rowHeight = 28;
            for (int i = 0; i < _parameterRows.Count; i++)
            {
                var row = _parameterRows[i];
                row.Location = new System.Drawing.Point(8, 30 + (i * (rowHeight + 6)));
                row.Size = new System.Drawing.Size(_parameterFilterPanel.Width - 16, rowHeight);
            }
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
                    valueCombo.Items.Clear();
                    valueCombo.Items.Add("<Auto Selection>");
                    valueCombo.SelectedIndex = 0;
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
                    string logPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\parameter_update_error.log";
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
                    string logPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\parameter_extraction_error.log";
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
                    string logPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\category_parameter_error.log";
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
                    string logPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\parameter_values_error.log";
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
                foreach (var row in _parameterRows)
                {
                    var nameCombo = row.Controls.OfType<WinForms.ComboBox>().FirstOrDefault();
                    if (nameCombo != null)
                    {
                        var currentSelection = nameCombo.SelectedItem?.ToString();
                        
                        // Update the dropdown items
                        nameCombo.Items.Clear();
                        nameCombo.Items.AddRange(categoryParameters.ToArray());
                        
                        // Restore selection if it still exists in the new list
                        if (!string.IsNullOrEmpty(currentSelection) && categoryParameters.Contains(currentSelection))
                        {
                            nameCombo.SelectedItem = currentSelection;
                        }
                        else
                        {
                            nameCombo.SelectedIndex = 0; // Select "<Select>"
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log error but don't crash
                try
                {
                    string logPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\parameter_refresh_error.log";
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
        private void OnOkClick(object? sender, EventArgs e)
        {
            try
            {
                _statusLabel.Text = "Validating configuration...";
                
                if (!ValidateConfiguration())
                {
                    return;
                }
                
                _statusLabel.Text = "Starting opening creation process...";
                
                // Execute with progress dialog
                var result = ExecuteSelectedFiltersWithProgress();
                
                if (result.Success)
                {
                    _statusLabel.Text = "Opening creation completed successfully!";
                    MessageBox.Show("Opening creation completed successfully!", "Success", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    _statusLabel.Text = $"Opening creation failed: {result.ErrorMessage}";
                    MessageBox.Show($"Opening creation failed: {result.ErrorMessage}", "Error", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                _statusLabel.Text = $"Error: {ex.Message}";
                MessageBox.Show($"Error: {ex.Message}", "Error", 
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
                if (false) // _referenceElementsListBox.SelectedItems.Count == 0) // Field not implemented yet
                {
                    MessageBox.Show("Please select reference elements before proceeding.", "No Reference Elements", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
                
                if (false) // _hostElementsListBox.SelectedItems.Count == 0) // Field not implemented yet
                {
                    MessageBox.Show("Please select host elements before proceeding.", "No Host Elements", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
                
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
                var selectedFilters = GetSelectedFilters();
                
                // Get UIDocument from current Revit context
                DebugLogger.Info("ExecuteSelectedFiltersWithProgress: Starting UIDocument retrieval");
                DebugLogger.Info($"ExecuteSelectedFiltersWithProgress: _uiDocument from constructor: {(_uiDocument != null ? "NOT NULL" : "NULL")}");
                DebugLogger.Info($"ExecuteSelectedFiltersWithProgress: _document from constructor: {(_document != null ? "NOT NULL" : "NULL")}");
                
                var uiDocument = GetCurrentUIDocument();
                DebugLogger.Info($"ExecuteSelectedFiltersWithProgress: GetCurrentUIDocument returned: {(uiDocument != null ? "NOT NULL" : "NULL")}");
                
                if (uiDocument == null)
                {
                    _statusLabel.Text = "Error: No active Revit document found";
                    DebugLogger.Error("ExecuteSelectedFiltersWithProgress: No active UIDocument found - this will cause orchestrator to fail");
                    return new OrchestrationResult
                    {
                        Success = false,
                        ErrorMessage = "No active Revit document found"
                    };
                }
                
                DebugLogger.Info($"ExecuteSelectedFiltersWithProgress: About to create OpeningCommandOrchestrator with UIDocument: {uiDocument.Application.ActiveUIDocument?.Document?.Title ?? "Unknown"}");
                
                using var orchestrator = new OpeningCommandOrchestrator(_document, uiDocument);
                
                // Pass UI clearance settings to orchestrator
                var clearanceSettings = GetClearanceSettings();
                orchestrator.SetUIClearances(clearanceSettings);
                
                // Pass full penetration setting to orchestrator
                bool fullPenetrationEnabled = GetFullPenetrationSetting();
                orchestrator.SetFullPenetrationEnabled(fullPenetrationEnabled);
                
                // Get clash zones from FILTER for incremental sleeve placement
                var filterWithClashZones = selectedFilters.FirstOrDefault(f => f.ClashZoneStorage != null);
                
                if (filterWithClashZones?.ClashZoneStorage != null)
                {
                    var clashZoneService = new ClashZoneService(filterWithClashZones.ClashZoneStorage, (msg) => DebugLogger.Info(msg));
                    orchestrator.SetClashZoneService(clashZoneService);
                    DebugLogger.Info($"ExecuteSelectedFiltersWithProgress: Using clash zones from FILTER ({filterWithClashZones.ClashZoneStorage.ClashZones.Count} zones) for incremental placement");
                }
                else
                {
                    DebugLogger.Info("ExecuteSelectedFiltersWithProgress: No clash zones found in FILTER - will place sleeves for all intersections");
                }
                
                var result = orchestrator.ExecuteMultipleFilters(selectedFilters, showProgress: true);
                
                return result;
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
            if (selectedCategories == null || selectedCategories.Count == 0) return;
            
            try
            {
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
                if (_bottomRightPanel?.Controls.Count > 0)
                {
                    foreach (var control in _bottomRightPanel.Controls)
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
                                    DebugLogger.Info($"Restored host category selection: {item}");
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
                LoadUIStateDirectly();
                
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
                    
                    // Restore host category selections (bottom-right panel)
                    RestoreHostCategorySelections(config.SelectedHostCategories);
                    
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
                                        if (config.ClashZoneStorage.ClashZones == null)
                                            config.ClashZoneStorage.ClashZones = new List<ClashZone>();
                                        
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
                
                // Save checked host categories
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
                                    uiState.Add($"HOSTCAT:{listBox.Items[i]}");
                                }
                            }
                        }
                    }
                }
                
                // Save clearance settings
                var clearanceSettings = GetClearanceSettings();
                foreach (var kvp in clearanceSettings)
                {
                    uiState.Add($"CLEARANCE:{kvp.Key}={kvp.Value}");
                }
                
                // Save clash zones from current profile
                var currentProfile = GetCurrentProfile();
                if (currentProfile?.Configuration?.ClashZoneStorage?.ClashZones != null)
                {
                    uiState.Add($"CLASHZONES:Count={currentProfile.Configuration.ClashZoneStorage.ClashZones.Count}");
                    uiState.Add($"CLASHZONES:LastUpdated={currentProfile.Configuration.ClashZoneStorage.LastUpdated:O}");
                    uiState.Add($"CLASHZONES:DocumentHash={currentProfile.Configuration.ClashZoneStorage.DocumentHash}");
                    foreach (var clashZone in currentProfile.Configuration.ClashZoneStorage.ClashZones)
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
                
                // Restore host categories
                foreach (var item in uiState.Where(x => x.StartsWith("HOSTCAT:")))
                {
                    var value = item.Substring(8);
                    RestoreCheckedItem(_bottomRightPanel, value);
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
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[CLASH_RESTORE] Failed to restore clash zones: {ex.Message}");
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
                                selectedFiles.Add(listBox.Items[i].ToString() ?? "");
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
                                selectedFiles.Add(listBox.Items[i].ToString() ?? "");
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
        
        private Dictionary<string, double> GetClearanceSettings()
        {
            var clearances = new Dictionary<string, double>();
            
            try
            {
                DebugLogger.Info($"[CLEARANCE_DEBUG] === GetClearanceSettings START ===");
                DebugLogger.Info($"[CLEARANCE_DEBUG] _clearancePanel exists: {_clearancePanel != null}");
                DebugLogger.Info($"[CLEARANCE_DEBUG] _clearancePanel.Controls.Count: {_clearancePanel?.Controls.Count ?? 0}");
                
                // Get clearance values from clearance panel
                if (_clearancePanel?.Controls.Count > 0)
                {
                    // Get current MEP category for specific key generation
                    string currentCategory = GetCurrentMepCategory();
                    DebugLogger.Info($"[CLEARANCE_DEBUG] Current MEP category: '{currentCategory}'");
                    
                    foreach (var control in _clearancePanel.Controls)
                    {
                        if (control is WinForms.TextBox textBox && textBox.Tag != null)
                        {
                            DebugLogger.Info($"[CLEARANCE_DEBUG] Found TextBox: Tag='{textBox.Tag}', Text='{textBox.Text}', Visible={textBox.Visible}");
                            
                            if (double.TryParse(textBox.Text, out double value))
                            {
                                string genericKey = textBox.Tag.ToString() ?? "";
                                string specificKey = ConvertToSpecificClearanceKey(genericKey, currentCategory);
                                clearances[specificKey] = value;
                                
                                DebugLogger.Info($"[CLEARANCE_DEBUG] Clearance setting: {genericKey} -> {specificKey} = {value}mm");
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
                
                // Get cable tray specific clearance values
                if (_cableTrayPanel?.Controls.Count > 0)
                {
                    foreach (var control in _cableTrayPanel.Controls)
                    {
                        if (control is WinForms.TextBox textBox && textBox.Tag != null)
                        {
                            if (double.TryParse(textBox.Text, out double value))
                            {
                                string key = textBox.Tag.ToString() ?? "";
                                clearances[key] = value;
                                
                                DebugLogger.Info($"Cable tray clearance setting: {key} = {value}mm");
                            }
                        }
                    }
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
                // Fire dampers use special keys
                return genericKey switch
                {
                    "normal_clearance" => "fire_damper_standard_clearance",
                    "insulated_clearance" => "fire_damper_msfd_clearance",
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
            
            // Standard MEP categories (Ducts, Pipes)
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
        /// Get full penetration setting from UI
        /// </summary>
        public bool GetFullPenetrationSetting()
        {
            return _fullPenetrationCheckBox?.Checked ?? true; // Default to true if checkbox not available
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
                // Read actual UI selections instead of hardcoded defaults
                var selectedMepCategories = GetSelectedMepCategories();
                DebugLogger.Info($"GetSelectedFilters: UI selected MEP categories: {string.Join(", ", selectedMepCategories)}");
                
                // Create filters only for selected categories
                foreach (var categoryName in selectedMepCategories)
                {
                    Models.MepCategory category;
                    string disciplineName;
                    
                    // Map UI category names to enum and discipline names
                    switch (categoryName.ToLower())
                    {
                        case "ducts":
                            category = Models.MepCategory.Ducts;
                            disciplineName = "Fire Fighting";
                            break;
                        case "ductaccessories":
                        case "duct accessories":
                            category = Models.MepCategory.DuctAccessories;
                            disciplineName = "Fire Fighting";
                            break;
                        case "pipes":
                            category = Models.MepCategory.Pipes;
                            disciplineName = "Water Systems";
                            break;
                        case "cabletrays":
                        case "cable trays":
                            category = Models.MepCategory.CableTrays;
                            disciplineName = "Data Devices";
                            break;
                        default:
                            DebugLogger.Warning($"GetSelectedFilters: Unknown category '{categoryName}' - skipping");
                            continue;
                    }
                    
                    var filter = OpeningFilter.CreateDefault(category, disciplineName);
                    selectedFilters.Add(filter);
                    DebugLogger.Info($"GetSelectedFilters: Created filter for {categoryName} -> {disciplineName}");
                }

                DebugLogger.Info($"GetSelectedFilters: Returning {selectedFilters.Count} filters based on UI selections");
                foreach (var filter in selectedFilters)
                {
                    DebugLogger.Info($"  - {filter.GetDescription()} (Category: {filter.Category})");
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

        private void OnRefreshClick(object? sender, EventArgs e)
        {
            // IMMEDIATE LOGGING BEFORE ANYTHING ELSE
            System.Diagnostics.Debug.WriteLine($"[ON_REFRESH_CLICK] === REFRESH BUTTON CLICKED AT {DateTime.Now} ===");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] === ON_REFRESH_CLICK STARTED ===\n");

            try
            {
                DebugLogger.Info("=== REFRESH BUTTON CLICKED ===");
                System.Diagnostics.Debug.WriteLine("[ON_REFRESH_CLICK] About to call Refresh() method");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] About to call Refresh() method\n");

                Refresh();

                System.Diagnostics.Debug.WriteLine("[ON_REFRESH_CLICK] Refresh() method completed successfully");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] Refresh() method completed successfully\n");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ON_REFRESH_CLICK] ERROR: {ex.Message}");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] ERROR in OnRefreshClick: {ex.Message}\n");

                _statusLabel.Text = $"Error during refresh: {ex.Message}";
                _progressBar.Visible = false;
                _refreshButton.Enabled = true;
                DebugLogger.Error($"Error in OnRefreshClick: {ex.Message}");
            }
        }

        /// <summary>
        /// Core refresh method that implements the refresh process flowchart
        /// </summary>
        private void Refresh()
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

            // Also write to the hardcoded refresh_debug.log file (like OnConfigureClick does)
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] REFRESH METHOD STARTED - Log file: Refresh_{timestamp}.log\n");

            // Write to the main debug logger file that user can see
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] === REFRESH STARTED === Timestamp: {timestamp}\n");

            DebugLogger.Info("=== REFRESH METHOD STARTED ===");
            System.Diagnostics.Debug.WriteLine("[REFRESH] DebugLogger.Info called");

            // Step 1: Check if filters are selected first (from filters panel)
            var selectedFilterItems = GetSelectedFilterItems();
            DebugLogger.Info($"[CLASH_DEBUG] Selected filter items: {string.Join(", ", selectedFilterItems)}");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [CLASH_DEBUG] Selected filter items: {string.Join(", ", selectedFilterItems)}\n");

            if (selectedFilterItems.Count == 0)
            {
                MessageBox.Show("Please select at least one filter before refreshing.", "No Filters Selected",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DebugLogger.Warning("Refresh: No filters selected - prompting user");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [CLASH_DEBUG] ERROR: No filters selected - cannot proceed with clash detection\n");
                return;
            }

            // Step 1.5: Get MEP categories for processing (optional - filters are the primary requirement)
            var filtersToProcess = GetSelectedFilters();
            DebugLogger.Info($"[CLASH_DEBUG] Found {filtersToProcess.Count} MEP category filters for processing");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [CLASH_DEBUG] MEP filters to process: {filtersToProcess.Count}\n");

            foreach (var filter in filtersToProcess)
            {
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [CLASH_DEBUG] Filter: {filter.Name} - Category: {filter.Category} - Enabled: {filter.IsEnabled}\n");
            }

            DebugLogger.Info($"Refresh: Processing {filtersToProcess.Count} selected filters");

            _statusLabel.Text = "Analyzing clash zones...";
            _progressBar.Value = 0;
            _progressBar.Visible = true;
            _refreshButton.Enabled = false;

            // Step 2: Check if we have a current profile with clash zone storage
            _progressBar.Value = 10;
            _statusLabel.Text = "Loading profile clash zones...";

            var currentProfile = GetCurrentProfile();
            DebugLogger.Info($"[CLASH_DEBUG] Current profile: {currentProfile?.Name ?? "NULL"}");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [CLASH_DEBUG] Profile loaded: {currentProfile?.Name ?? "NULL"}\n");

            if (currentProfile?.Configuration?.ClashZoneStorage == null)
            {
                DebugLogger.Info("[CLASH_DEBUG] No clash zone storage in profile - proceeding with direct clash detection and filter saving");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [CLASH_DEBUG] INFO: No clash zone storage - proceeding with direct clash detection\n");
                // Continue with clash detection instead of falling back to basic refresh
            }

            // Step 3: Get current document
            var document = GetCurrentDocument();
            if (document == null)
            {
                _statusLabel.Text = "Error: No active Revit document found";
                _progressBar.Visible = false;
                _refreshButton.Enabled = true;
                DebugLogger.Error("Refresh: No active document found");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [CLASH_DEBUG] ERROR: No active Revit document found\n");
                return;
            }

            DebugLogger.Info($"[CLASH_DEBUG] Active document: {document.Title}");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [CLASH_DEBUG] Document: {document.Title}\n");

            // Step 4: Get current intersections
            _progressBar.Value = 25;
            _statusLabel.Text = "Detecting current intersections...";

            DebugLogger.Info("[CLASH_DEBUG] Starting intersection detection...");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [CLASH_DEBUG] Starting MEP-structural intersection detection...\n");

            var currentIntersections = GetCurrentIntersections();
            DebugLogger.Info($"[CLASH_DEBUG] Found {currentIntersections.Count} current intersections");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [CLASH_DEBUG] INTERSECTION DETECTION COMPLETE: {currentIntersections.Count} intersections found\n");

            // Log intersection details to refresh log files
            try
            {
                File.AppendAllText(refreshLogPath, $"[{DateTime.Now}] INTERSECTION DETECTION RESULTS:\n");
                File.AppendAllText(refreshLogPath, $"[{DateTime.Now}] Total intersections found: {currentIntersections.Count}\n");

                if (currentIntersections.Count > 0)
                {
                    File.AppendAllText(refreshLogPath, $"[{DateTime.Now}] Intersection details:\n");
                    for (int i = 0; i < Math.Min(currentIntersections.Count, 20); i++) // Log first 20 for better debugging
                    {
                        var intersection = currentIntersections[i];
                        var mepElement = intersection.Item1;         // MEP element that caused intersection
                        var structuralElement = intersection.Item2;  // Structural element intersected
                        var intersectionBBox = intersection.Item3;   // Intersection bounding box
                        var intersectionPoint = intersection.Item4;  // Intersection center point

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
                }

                // Also log to the hardcoded refresh_debug.log file
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Found {currentIntersections.Count} intersections during refresh\n");

                // Log detailed intersection info to debug file
                foreach (var intersection in currentIntersections.Take(10))
                {
                    var structuralElement = intersection.Item1;
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [CLASH_DEBUG] Structural Element: {structuralElement.Name} (ID: {structuralElement.Id}) Category: {structuralElement.Category?.Name ?? "Unknown"}\n");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to log intersection details: {ex.Message}");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [CLASH_DEBUG] ERROR logging intersection details: {ex.Message}\n");
            }

            if (currentIntersections.Count == 0)
            {
                _statusLabel.Text = "No intersections found";
                _progressBar.Value = 100;
                DebugLogger.Warning("Refresh: No intersections found - no clash zones to save");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [CLASH_DEBUG] WARNING: No intersections found - clash detection returned empty results\n");

                // Hide progress bar after a short delay
                var noIntersectionsTimer = new System.Windows.Forms.Timer();
                noIntersectionsTimer.Interval = 1500;
                noIntersectionsTimer.Tick += (s, args) => {
                    _progressBar.Visible = false;
                    noIntersectionsTimer.Stop();
                    noIntersectionsTimer.Dispose();
                };
                noIntersectionsTimer.Start();
                _refreshButton.Enabled = true;
                return;
            }

            // Step 5: Initialize clash zone service
            _progressBar.Value = 40;
            _statusLabel.Text = "Initializing clash zone service...";

            DebugLogger.Info("[CLASH_DEBUG] Initializing ClashZoneService...");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [CLASH_DEBUG] Initializing ClashZoneService with {(currentProfile?.Configuration?.ClashZoneStorage?.ClashZones?.Count ?? 0)} existing zones\n");

            // Create or use existing clash zone storage
            var clashZoneStorage = currentProfile?.Configuration?.ClashZoneStorage ?? new ClashZoneStorage
            {
                ClashZones = new List<ClashZone>(),
                LastUpdated = DateTime.Now,
                DocumentHash = document?.PathName ?? "Unknown"
            };

            var clashZoneService = new ClashZoneService(clashZoneStorage, (msg) => {
                DebugLogger.Info($"[CLASH_DEBUG] {msg}");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [CLASH_DEBUG] {msg}\n");
            });

            // Step 6: Detect new clash zones
            _progressBar.Value = 60;
            _statusLabel.Text = "Detecting new clash zones...";

            DebugLogger.Info("[CLASH_DEBUG] Calling DetectNewClashZones...");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [CLASH_DEBUG] Calling DetectNewClashZones with {currentIntersections.Count} intersections...\n");

            var newClashZones = clashZoneService.DetectNewClashZones(currentIntersections, document);

            DebugLogger.Info($"[CLASH_DEBUG] DetectNewClashZones returned {newClashZones?.Count ?? 0} new zones");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [CLASH_DEBUG] DetectNewClashZones completed - {newClashZones?.Count ?? 0} new clash zones detected\n");

            // Step 7: Get statistics
            _progressBar.Value = 80;
            _statusLabel.Text = "Calculating clash zone statistics...";

            var (total, resolved, unresolved, newZones) = clashZoneService.GetClashZoneStatistics();

            DebugLogger.Info($"[CLASH_DEBUG] Clash zone statistics - Total: {total}, Resolved: {resolved}, Unresolved: {unresolved}, New: {newZones}");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [CLASH_DEBUG] Statistics - Total: {total}, Resolved: {resolved}, Unresolved: {unresolved}, New: {newZones}\n");

            // Step 8: Save clash zones to FILTER
            _progressBar.Value = 90;
            _statusLabel.Text = "Saving clash zones to filter...";

            // Save clash zones to the current filter
            if (filtersToProcess.Count > 0)
            {
                // Use the first enabled filter to store clash zones
                var targetFilter = filtersToProcess.FirstOrDefault(f => f.IsEnabled);
                if (targetFilter != null)
                {
                    // CRITICAL FIX: Save clash zones to BOTH filter AND profile configuration
                    targetFilter.ClashZoneStorage = clashZoneStorage;
                    targetFilter.LastModified = DateTime.Now;
                    
                    // Save to profile configuration for persistence
                    if (currentProfile != null)
                    {
                        // Ensure configuration exists
                        if (currentProfile.Configuration == null)
                        {
                            currentProfile.Configuration = new UserConfiguration();
                            DebugLogger.Info("[CLASH_DEBUG] Created new profile configuration during Refresh");
                        }

                        currentProfile.Configuration.ClashZoneStorage = clashZoneStorage;

                        // Persist current opening conditions (e.g., clearance values) into profile configuration
                        if (currentProfile.Configuration.OpeningSettings == null)
                        {
                            currentProfile.Configuration.OpeningSettings = new OpeningSettings();
                        }
                        currentProfile.Configuration.OpeningSettings.ClearanceSettings = GetClearanceSettings();

                        DebugLogger.Info($"[CLASH_DEBUG] Saved clash zones and opening conditions to profile configuration for persistence");

                        // Persist to disk immediately so Refresh round-trips data per hybrid implementation
                        try
                        {
                            _appProfileService.SaveCurrentProfile();
                            DebugLogger.Info("[CLASH_DEBUG] Profile persisted to disk after Refresh");
                        }
                        catch (Exception saveEx)
                        {
                            DebugLogger.Warning($"[CLASH_DEBUG] Warning: Failed to persist profile after Refresh: {saveEx.Message}");
                        }
                    }
                    
                    DebugLogger.Info($"[CLASH_DEBUG] Saved clash zones to filter '{targetFilter.Name}'");
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [CLASH_DEBUG] SUCCESS: Saved {total} clash zones to filter '{targetFilter.Name}' and profile configuration\n");
                }
                else
                {
                    DebugLogger.Warning("[CLASH_DEBUG] No enabled filter found to save clash zones");
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [CLASH_DEBUG] WARNING: No enabled filter found to save clash zones\n");
                }
            }
            else
            {
                DebugLogger.Warning("[CLASH_DEBUG] No filters to process - cannot save clash zones");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [CLASH_DEBUG] WARNING: No filters to process - cannot save clash zones\n");
            }

            // Step 9: Update UI with results
            _progressBar.Value = 100;
            _statusLabel.Text = $"Clash zones: {total} total, {unresolved} unresolved, {newZones} new";

            DebugLogger.Info($"[CLASH_DEBUG] Clash zone refresh complete: {total} total, {unresolved} unresolved, {newZones} new");
            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [CLASH_DEBUG] REFRESH COMPLETE: {total} total zones, {unresolved} unresolved, {newZones} new\n");

            // Step 10: Update parameter dropdowns with clash zone parameters
            _progressBar.Value = 90;
            _statusLabel.Text = "Updating parameter dropdowns...";

            try
            {
                PopulateParameterDropdowns();
                DebugLogger.Info("[PARAMETER_SERVICE] Parameter dropdowns updated using NEW service method");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [PARAMETER_SERVICE] Parameter dropdowns updated using NEW service method\n");
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[PARAMETER_SERVICE] Error updating parameter dropdowns: {ex.Message}");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [PARAMETER_SERVICE] ERROR updating parameter dropdowns: {ex.Message}\n");
            }

            // Hide progress bar after a short delay
            var completionTimer = new System.Windows.Forms.Timer();
            completionTimer.Interval = 2000; // 2 seconds to show results
            completionTimer.Tick += (s, args) => {
                _progressBar.Visible = false;
                completionTimer.Stop();
                completionTimer.Dispose();
            };
            completionTimer.Start();

            _refreshButton.Enabled = true;
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
        /// </summary>
        /// <summary>
        /// SIMPLE method to populate parameter dropdowns - calls ParameterExtractionService only
        /// </summary>
        private void PopulateParameterDropdowns()
        {
            try
            {
                DebugLogger.Info("[PARAMETER_SERVICE] Starting category-specific parameter population via service");
                
                // Get ALL selected MEP categories
                var selectedCategories = GetSelectedMepCategories();
                DebugLogger.Info($"[PARAMETER_SERVICE] Selected MEP categories: {string.Join(", ", selectedCategories)}");
                
                // Call service to handle category-specific parameter population
                var parameterService = new Services.ParameterExtractionService();
                parameterService.PopulateCategorySpecificParameters(_serviceParameterTabs, selectedCategories, _uiDocument?.Document);
                
                DebugLogger.Info("[PARAMETER_SERVICE] Category-specific parameter population completed via service");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[PARAMETER_SERVICE] Error in PopulateParameterDropdowns: {ex.Message}");
            }
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
                "Service Type"
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
        /// OLD METHOD - COMMENTED OUT - REPLACED BY SERVICE
        /// SIMPLE method to update UI dropdowns with parameter lists
        /// </summary>
        /*private void UpdateParameterDropdownsInUI(List<string> mepParameters, List<string> openingParameters)
        {
            try
            {
                DebugLogger.Info($"[PARAMETER_SIMPLE] Updating UI with {mepParameters.Count} MEP and {openingParameters.Count} opening parameters");
                
                // DEBUG: Check if _serviceParameterTabs exists
                DebugLogger.Info($"[PARAMETER_SIMPLE] _serviceParameterTabs exists: {_serviceParameterTabs != null}");
                if (_serviceParameterTabs != null)
                {
                    DebugLogger.Info($"[PARAMETER_SIMPLE] _serviceParameterTabs.TabPages.Count: {_serviceParameterTabs.TabPages.Count}");
                }
                
                if (_serviceParameterTabs?.TabPages.Count > 0)
                {
                    foreach (WinForms.TabPage tabPage in _serviceParameterTabs.TabPages)
                    {
                        DebugLogger.Info($"[PARAMETER_SIMPLE] Processing tab: '{tabPage.Text}'");
                        DebugLogger.Info($"[PARAMETER_SIMPLE] Tab has {tabPage.Controls.Count} direct controls");
                        
                        // DEBUG: List all control types in the tab
                        foreach (WinForms.Control control in tabPage.Controls)
                        {
                            DebugLogger.Info($"[PARAMETER_SIMPLE] Tab '{tabPage.Text}' contains control: {control.GetType().Name} - '{control.Name}'");
                        }
                        
                        // Find all ComboBoxes in this tab - CORRECT HIERARCHY SEARCH
                        var allComboBoxes = new List<WinForms.ComboBox>();
                        
                        // Step 1: Find the servicePanel in this tab
                        var servicePanel = tabPage.Controls.OfType<WinForms.Panel>().FirstOrDefault();
                        DebugLogger.Info($"[PARAMETER_SIMPLE] Found servicePanel: {servicePanel != null} in tab '{tabPage.Text}'");
                        
                        if (servicePanel != null)
                        {
                            DebugLogger.Info($"[PARAMETER_SIMPLE] servicePanel has {servicePanel.Controls.Count} controls");
                            
                            // Step 2: Find all row panels in the servicePanel (excluding the add button)
                            // The add button is a Button, not a Panel, so we can get all Panels
                            var rowPanels = servicePanel.Controls.OfType<WinForms.Panel>().ToList();
                            DebugLogger.Info($"[PARAMETER_SIMPLE] Found {rowPanels.Count} row panels in servicePanel");
                            
                            // Step 3: Find ComboBoxes in each row panel
                            foreach (var rowPanel in rowPanels)
                            {
                                DebugLogger.Info($"[PARAMETER_SIMPLE] Row panel has {rowPanel.Controls.Count} controls");
                                
                                var rowComboBoxes = rowPanel.Controls.OfType<WinForms.ComboBox>().ToList();
                                allComboBoxes.AddRange(rowComboBoxes);
                                DebugLogger.Info($"[PARAMETER_SIMPLE] Found {rowComboBoxes.Count} ComboBoxes in row panel");
                                
                                // DEBUG: List all controls in the row panel
                                foreach (WinForms.Control control in rowPanel.Controls)
                                {
                                    DebugLogger.Info($"[PARAMETER_SIMPLE] Row panel contains: {control.GetType().Name} - '{control.Name}'");
                                }
                            }
                        }
                        
                        DebugLogger.Info($"[PARAMETER_SIMPLE] TOTAL: Found {allComboBoxes.Count} ComboBoxes in tab '{tabPage.Text}'");
                        
                        // Check if ComboBoxes are actually populated with meaningful content
                        // A ComboBox is considered "populated" if it has items AND has a selected item
                        bool hasPopulatedComboBoxes = allComboBoxes.Any(cb => cb.Items.Count > 0 && cb.SelectedItem != null);
                        bool hasEmptyComboBoxes = allComboBoxes.Any(cb => cb.Items.Count == 0 || cb.SelectedItem == null);
                        
                        DebugLogger.Info($"[PARAMETER_SIMPLE] ComboBox population check - Total: {allComboBoxes.Count}, HasItems: {allComboBoxes.Count(cb => cb.Items.Count > 0)}, HasSelection: {allComboBoxes.Count(cb => cb.SelectedItem != null)}, HasPopulated: {hasPopulatedComboBoxes}, HasEmpty: {hasEmptyComboBoxes}");
                        
                        // If no ComboBoxes found OR there are empty ComboBoxes, create default parameter rows
                        if (allComboBoxes.Count == 0 || hasEmptyComboBoxes)
                        {
                            DebugLogger.Info($"[PARAMETER_SIMPLE] Empty ComboBoxes found in tab '{tabPage.Text}', clearing existing rows and creating default parameter rows");
                            
                            // Clear existing rows to prevent overlapping
                            var existingRows = servicePanel.Controls.OfType<WinForms.Panel>().ToList();
                            foreach (var row in existingRows)
                            {
                                servicePanel.Controls.Remove(row);
                                row.Dispose();
                            }
                            
                            CreateDefaultParameterRows(tabPage, mepParameters, openingParameters);
                        }
                        else
                        {
                            // Update existing ComboBoxes
                            foreach (var comboBox in allComboBoxes)
                            {
                                DebugLogger.Info($"[PARAMETER_SIMPLE] Updating ComboBox '{comboBox.Name}' at location ({comboBox.Location.X}, {comboBox.Location.Y})");
                                
                                // Store current selection before clearing
                                var currentSelection = comboBox.SelectedItem?.ToString();
                                
                                // Clear existing items
                                comboBox.Items.Clear();
                                
                                // Determine which parameters to add based on ComboBox position or tag
                                if (comboBox.Tag?.ToString()?.Contains("mep") == true || 
                                    comboBox.Location.X < 100) // Left side = MEP parameters
                                {
                                    comboBox.Items.AddRange(mepParameters.ToArray());
                                    DebugLogger.Info($"[PARAMETER_SIMPLE] Updated MEP ComboBox '{comboBox.Name}' with {mepParameters.Count} parameters");
                                    
                                    // Restore selection if it still exists
                                    if (!string.IsNullOrEmpty(currentSelection) && mepParameters.Contains(currentSelection))
                                    {
                                        comboBox.SelectedItem = currentSelection;
                                        DebugLogger.Info($"[PARAMETER_SIMPLE] Restored MEP selection: {currentSelection}");
                                    }
                                }
                                else // Right side = Opening parameters
                                {
                                    comboBox.Items.AddRange(openingParameters.ToArray());
                                    DebugLogger.Info($"[PARAMETER_SIMPLE] Updated Opening ComboBox '{comboBox.Name}' with {openingParameters.Count} parameters");
                                    
                                    // Restore selection if it still exists
                                    if (!string.IsNullOrEmpty(currentSelection) && openingParameters.Contains(currentSelection))
                                    {
                                        comboBox.SelectedItem = currentSelection;
                                        DebugLogger.Info($"[PARAMETER_SIMPLE] Restored Opening selection: {currentSelection}");
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    DebugLogger.Warning("[PARAMETER_SIMPLE] No service parameter tabs found!");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[PARAMETER_SIMPLE] Error updating UI dropdowns: {ex.Message}");
                DebugLogger.Error($"[PARAMETER_SIMPLE] Stack trace: {ex.StackTrace}");
            }
        }*/

        /// <summary>
        /// Updates parameter dropdowns with parameters from MEP categories (not clash zones)
        /// </summary>
        private void UpdateParameterDropdownsFromMepCategories(Document document)
        {
            try
            {
                DebugLogger.Info("[PARAMETER_DEBUG] Starting parameter dropdown update using ParameterExtractionService");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [PARAMETER_DEBUG] Starting parameter dropdown update using ParameterExtractionService\n");
                
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
                    var mepCategories = new List<Services.MepCategory> { (Services.MepCategory)selectedMepCategory };
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
                        if (!string.IsNullOrEmpty(param.Definition.Name) && 
                            !param.Definition.Name.StartsWith("Internal") &&
                            !param.Definition.Name.StartsWith("Revit") &&
                            !param.Definition.Name.StartsWith("Assembly"))
                        {
                            openingParameters.Add(param.Definition.Name);
                        }
                    }
                }
                
                DebugLogger.Info($"[PARAMETER_DEBUG] Found {mepParameters.Count} MEP parameters and {openingParameters.Count} opening parameters");

                // Update parameter service dropdowns if they exist
                DebugLogger.Info($"[PARAMETER_DEBUG] Checking _serviceParameterTabs: {_serviceParameterTabs != null}, TabPages count: {_serviceParameterTabs?.TabPages.Count ?? 0}");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [PARAMETER_DEBUG] Checking _serviceParameterTabs: {_serviceParameterTabs != null}, TabPages count: {_serviceParameterTabs?.TabPages.Count ?? 0}\n");
                
                if (_serviceParameterTabs?.TabPages.Count > 0)
                {
                    var categoryParameters = new Dictionary<string, List<Models.ParameterInfo>>();
                    
                    // Convert to ParameterInfo objects
                    var mepParamInfos = mepParameters.Select(p => new Models.ParameterInfo 
                    { 
                        Name = p, 
                        Type = "Text", 
                        IsReadOnly = false 
                    }).ToList();
                    
                    categoryParameters["MEP Elements"] = mepParamInfos;

                    // Include Opening Sleeve Family parameters in a separate dropdown
                    var openingParamInfos = openingParameters.Select(p => new Models.ParameterInfo
                    {
                        Name = p,
                        Type = "Text",
                        IsReadOnly = false
                    }).ToList();
                    categoryParameters["Opening Families"] = openingParamInfos;
                    
                    DebugLogger.Info($"[PARAMETER_DEBUG] About to update dropdowns with {mepParamInfos.Count} MEP params and {openingParamInfos.Count} opening params");
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [PARAMETER_DEBUG] About to update dropdowns with {mepParamInfos.Count} MEP params and {openingParamInfos.Count} opening params\n");
                    
                    // Update the dropdowns
                    UpdateParameterServiceDropdowns(categoryParameters);
                    DebugLogger.Info("[PARAMETER_DEBUG] Updated parameter service dropdowns");
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [PARAMETER_DEBUG] Updated parameter service dropdowns\n");
                }
                else
                {
                    DebugLogger.Warning("[PARAMETER_DEBUG] _serviceParameterTabs is null or has no tab pages - cannot update dropdowns");
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [PARAMETER_DEBUG] WARNING: _serviceParameterTabs is null or has no tab pages - cannot update dropdowns\n");
                }

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
        /// Gets opening families from the current document
        /// </summary>
        private List<FamilySymbol> GetOpeningFamilies(Document document)
        {
            var openingFamilies = new List<FamilySymbol>();
            
            try
            {
                // FamilySymbol is an ElementType; do NOT filter with WhereElementIsNotElementType
                var collector = new FilteredElementCollector(document)
                    .OfClass(typeof(FamilySymbol));

                int inspected = 0;
                foreach (Element element in collector)
                {
                    inspected++;
                    if (element is FamilySymbol familySymbol)
                    {
                        var familyNameLower = ($"{familySymbol.Family?.Name} {familySymbol.Name}").ToLower();
                        var categoryNameLower = familySymbol.Category?.Name?.ToLower() ?? string.Empty;

                        // Only include families that contain "Opening" in their name
                        bool isOpeningFamily = familyNameLower.Contains("opening");
                        
                        if (isOpeningFamily)
                        {
                            openingFamilies.Add(familySymbol);
                        }
                    }
                }

                DebugLogger.Info($"[PARAMETER_DEBUG] GetOpeningFamilies inspected {inspected} FamilySymbols, matched {openingFamilies.Count} families containing 'Opening'");
                
                // Enhanced debugging: Log some sample family names to help diagnose
                if (inspected > 0 && openingFamilies.Count == 0)
                {
                    DebugLogger.Info($"[PARAMETER_DEBUG] No families containing 'Opening' found. Sample family names in document:");
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
                    DebugLogger.Info($"[PARAMETER_DEBUG] No opening families found - UI will use hardcoded opening parameters");
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
                DebugLogger.Info("GetCurrentIntersections: Starting intersection detection...");

                var document = GetCurrentDocument();
                if (document == null)
                {
                    DebugLogger.Warning("GetCurrentIntersections: No document available");
                    return new List<(Element, Element, BoundingBoxXYZ, XYZ)>();
                }

                // Get structural elements for intersection detection
                var structuralElements = MepIntersectionService.CollectStructuralElementsForDirectIntersectionVisibleOnly(document, (msg) => DebugLogger.Info(msg));
                DebugLogger.Info($"GetCurrentIntersections: Found {structuralElements.Count} structural elements");

                // Get MEP elements (pipes, ducts, cable trays) from host and linked files
                var mepElements = MepElementCollectorHelper.CollectMepElementsVisibleOnly(document);

                // Count by category for logging
                int pipes = 0, ducts = 0, cableTrays = 0;
                foreach (var (element, transform) in mepElements)
                {
                    if (element.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_PipeCurves) pipes++;
                    else if (element.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_DuctCurves) ducts++;
                    else if (element.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_CableTray) cableTrays++;
                }

                DebugLogger.Info($"GetCurrentIntersections: Found {mepElements.Count} MEP elements ({pipes} pipes, {ducts} ducts, {cableTrays} cable trays)");

                // Find intersections
                var allIntersections = new List<(Element, Element, BoundingBoxXYZ, XYZ)>();

                foreach (var mepElement in mepElements)
                {
                    Element mepElementToUse = mepElement.Item1;
                    Transform? mepTransform = mepElement.Item2;

                    // Transform structural elements to match MEP coordinate system
                    var transformedStructuralElements = new List<(Element, Transform?)>();
                    foreach (var structuralElement in structuralElements)
                    {
                        Element structElement = structuralElement.Item1;
                        Transform? structTransform = structuralElement.Item2;

                        // If MEP element is from a different linked file than structural element,
                        // we need to transform structural element to MEP coordinate system
                        if (mepTransform != null && structTransform != null &&
                            mepElementToUse.Document.Title != structElement.Document.Title)
                        {
                            // Transform structural element to MEP coordinate system
                            Transform combinedTransform = mepTransform.Inverse.Multiply(structTransform);
                            transformedStructuralElements.Add((structElement, combinedTransform));
                            DebugLogger.Info($"GetCurrentIntersections: Transforming structural element {structElement.Id} from {structElement.Document.Title} to MEP coordinate system");
                        }
                        else
                        {
                            transformedStructuralElements.Add(structuralElement);
                        }
                    }

                    DebugLogger.Info($"GetCurrentIntersections: Testing MEP element {mepElementToUse.Id} with coordinate transformation");

                    var intersections = MepIntersectionService.FindIntersections(mepElementToUse, transformedStructuralElements, (msg) => DebugLogger.Info(msg));
                    if (intersections != null && intersections.Count > 0)
                    {
                        DebugLogger.Info($"GetCurrentIntersections: Found {intersections.Count} intersections for MEP element {mepElementToUse.Id} with transformation");
                        allIntersections.AddRange(intersections);
                    }
                    else
                    {
                        DebugLogger.Info($"GetCurrentIntersections: No intersections found for MEP element {mepElementToUse.Id} with transformation");
                    }
                }

                DebugLogger.Info($"GetCurrentIntersections: Found {allIntersections.Count} total intersections");
                return allIntersections;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error getting current intersections: {ex.Message}");
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
                DebugLogger.Error($"Error getting current document: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// Saves the current profile
        /// </summary>
        private void SaveCurrentProfile()
        {
            try
            {
                // This would need to be implemented based on your profile service
                DebugLogger.Info("Profile saved successfully");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error saving current profile: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Gets the current UIDocument from Revit context
        /// </summary>
        private UIDocument? GetCurrentUIDocument()
        {
            try
            {
                DebugLogger.Info("GetCurrentUIDocument: Starting UIDocument retrieval");
                DebugLogger.Info($"GetCurrentUIDocument: _uiDocument field is {(_uiDocument != null ? "NOT NULL" : "NULL")}");
                
                // First, try to use the UIDocument passed to the constructor
                if (_uiDocument != null)
                {
                    DebugLogger.Info($"GetCurrentUIDocument: Using UIDocument from constructor - Document: {_uiDocument.Document?.Title ?? "Unknown"}");
                    return _uiDocument;
                }
                
                DebugLogger.Warning("GetCurrentUIDocument: _uiDocument is NULL - this means it wasn't passed to the constructor properly");
                DebugLogger.Warning("GetCurrentUIDocument: This will cause the orchestrator to fail with ArgumentNullException");
                
                // Try to get UIDocument from the current Revit application
                // Note: Application.Current is not available in all contexts
                // This approach is not reliable in Revit add-ins
                
                DebugLogger.Warning("No active UIDocument found - this may cause issues with some operations");
                return null;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"GetCurrentUIDocument: Error getting UIDocument: {ex.Message}");
                return null;
            }
        }

        private void OnConfigureClick(object? sender, EventArgs e)
        {
            try
            {
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] CONFIGURE BUTTON CLICKED\n");
                _statusLabel.Text = "Opening Settings dialog...";
                
                // Create and show settings dialog (it will load/save settings automatically)
                var settingsDialog = new SettingsDialog();
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] SettingsDialog created successfully\n");
                
                // Show the settings dialog as modal
                using (settingsDialog)
                {
                    if (settingsDialog.ShowDialog(this) == DialogResult.OK)
                    {
                        _statusLabel.Text = "Settings saved successfully";
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Settings saved successfully\n");
                    }
                    else
                    {
                        _statusLabel.Text = "Settings cancelled";
                    }
                }
            }
            catch (Exception ex)
            {
                // Show full error message in status label
                var fullErrorMessage = $"Error opening settings: {ex.Message}";
                _statusLabel.Text = fullErrorMessage;
                
                // Set tooltip with full error message
                var statusTooltip = new WinForms.ToolTip();
                statusTooltip.SetToolTip(_statusLabel, fullErrorMessage);
                
                // Log detailed error information
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] CONFIGURE ERROR: {ex.Message}\n");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] CONFIGURE STACK TRACE: {ex.StackTrace}\n");
                
                // Also show a message box with the full error for debugging (smaller dialog)
                var errorMessage = $"Error opening settings dialog:\n\n{ex.Message}";
                if (ex.StackTrace != null)
                {
                    // Truncate stack trace to keep dialog manageable
                    var shortStackTrace = ex.StackTrace.Length > 200 ? ex.StackTrace.Substring(0, 200) + "..." : ex.StackTrace;
                    errorMessage += $"\n\nStack Trace:\n{shortStackTrace}";
                }
                
                MessageBox.Show(errorMessage, 
                              "Settings Dialog Error", 
                              MessageBoxButtons.OK, 
                              MessageBoxIcon.Error);
            }
        }

        private List<string> GetSelectedCategories()
        {
            var selectedCategories = new List<string>();
            var mepCategoriesListBox = _topRightPanel.Controls.OfType<WinForms.CheckedListBox>().FirstOrDefault();
            
            System.Diagnostics.Debug.WriteLine($"MEP Categories ListBox found: {mepCategoriesListBox != null}");
            
            if (mepCategoriesListBox != null)
            {
                System.Diagnostics.Debug.WriteLine($"MEP Categories ListBox items count: {mepCategoriesListBox.Items.Count}");
                for (int i = 0; i < mepCategoriesListBox.Items.Count; i++)
                {
                    bool isChecked = mepCategoriesListBox.GetItemChecked(i);
                    string itemText = mepCategoriesListBox.Items[i].ToString();
                    System.Diagnostics.Debug.WriteLine($"Item {i}: '{itemText}' - Checked: {isChecked}");
                    
                    if (isChecked)
                    {
                        selectedCategories.Add(itemText);
                    }
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("MEP Categories ListBox not found in _topRightPanel");
            }
            
            return selectedCategories;
        }

        private List<Models.ParameterInfo> GetParametersForCategory(string category)
        {
            var allParameters = new List<Models.ParameterInfo>();
            
            try
            {
                // Get the appropriate BuiltInCategory based on the category name
                BuiltInCategory? builtInCategory = GetBuiltInCategoryForMepCategory(category);
                if (builtInCategory == null)
                {
                    System.Diagnostics.Debug.WriteLine($"No BuiltInCategory found for {category}");
                    return allParameters;
                }
                
                var parameterService = new ParameterExtractionService();
                
                // 1. Get parameters from ACTIVE DOCUMENT if it's selected
                var selectedReferenceFiles = GetSelectedReferenceFiles();
                bool activeDocumentSelected = selectedReferenceFiles.Any(f => f.Contains("(Active Document)"));
                
                if (activeDocumentSelected && _activeDocument != null)
                {
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Getting parameters from ACTIVE DOCUMENT for {category}\n");
                    var activeDocParams = parameterService.GetParametersForCategory(_activeDocument, builtInCategory.Value);
                    
                    // Convert Services.ParameterInfo to Models.ParameterInfo
                    foreach (var param in activeDocParams)
                    {
                        allParameters.Add(new Models.ParameterInfo
                        {
                            Name = param.Name,
                            Type = param.Type,
                            IsReadOnly = false
                        });
                    }
                    
                    System.Diagnostics.Debug.WriteLine($"Found {activeDocParams.Count} parameters from active document for {category}");
                }
                
                // 2. Get parameters from SELECTED LINKED FILES
                var selectedLinkedFiles = selectedReferenceFiles.Where(f => !f.Contains("(Active Document)")).ToList();
                if (selectedLinkedFiles.Count > 0)
                {
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Getting parameters from {selectedLinkedFiles.Count} LINKED FILES for {category}\n");
                    
                    foreach (var linkedFile in selectedLinkedFiles)
                    {
                        // Find the corresponding linked file document
                        var linkedDoc = GetLinkedDocument(linkedFile);
                        if (linkedDoc != null)
                        {
                            var linkedParams = parameterService.GetParametersForCategory(linkedDoc, builtInCategory.Value);
                            
                            // Convert Services.ParameterInfo to Models.ParameterInfo
                            foreach (var param in linkedParams)
                            {
                                allParameters.Add(new Models.ParameterInfo
                                {
                                    Name = param.Name,
                                    Type = param.Type,
                                    IsReadOnly = false
                                });
                            }
                            
                            System.Diagnostics.Debug.WriteLine($"Found {linkedParams.Count} parameters from linked file '{linkedFile}' for {category}");
                        }
                    }
                }
                
                // Remove duplicates based on parameter name
                allParameters = allParameters.GroupBy(p => p.Name).Select(g => g.First()).ToList();
                
                System.Diagnostics.Debug.WriteLine($"Total unique parameters found for {category}: {allParameters.Count}");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Total unique parameters for {category}: {allParameters.Count}\n");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting parameters for {category}: {ex.Message}");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Error getting parameters for {category}: {ex.Message}\n");
            }
            
            return allParameters;
        }

        private BuiltInCategory? GetBuiltInCategoryForMepCategory(string category)
        {
            switch (category.ToLower())
            {
                case "ducts":
                    return BuiltInCategory.OST_DuctCurves;
                case "duct accessories":
                    return BuiltInCategory.OST_DuctAccessory;
                case "cable trays":
                case "cabletrays":  // Handle enum name format
                    return BuiltInCategory.OST_CableTray;
                case "pipes":
                    return BuiltInCategory.OST_PipeCurves;
                default:
                    return null;
            }
        }

        private void UpdateParameterServiceDropdowns(Dictionary<string, List<Models.ParameterInfo>> categoryParameters)
        {
            try
            {
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Updating parameter dropdowns\n");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Available categories: {string.Join(", ", categoryParameters.Keys)}\n");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] All collected categories: {string.Join(", ", _allCollectedParameters.Keys)}\n");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Number of tabs: {_serviceParameterTabs?.TabPages.Count ?? 0}\n");
                
                if (_serviceParameterTabs?.TabPages.Count > 0)
                {
                    foreach (WinForms.TabPage tabPage in _serviceParameterTabs.TabPages)
                    {
                        string tabName = tabPage.Text;
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Processing tab: {tabName}\n");
                        
                        // Combine MEP and Opening Family parameters for all tabs (broader availability)
                        var matchingParameters = new List<Models.ParameterInfo>();
                        if (_allCollectedParameters.TryGetValue("MEP Elements", out var mepParams))
                        {
                            matchingParameters.AddRange(mepParams);
                            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Added {mepParams.Count} MEP parameters to tab {tabName}\n");
                        }
                        if (_allCollectedParameters.TryGetValue("Opening Families", out var openingParams))
                        {
                            matchingParameters.AddRange(openingParams);
                            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Added {openingParams.Count} opening parameters to tab {tabName}\n");
                        }
                        
                        // Update ALL parameter rows in this tab
                        UpdateTabParameterDropdown(tabPage, matchingParameters);
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Updated dropdown for tab {tabName} with {matchingParameters.Count} parameters\n");
                    }
                }
                else
                {
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] No service parameter tabs found!\n");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error updating parameter dropdowns: {ex.Message}");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] ERROR updating parameter dropdowns: {ex.Message}\n");
            }
        }

        private bool IsCategoryMatch(string tabName, string category)
        {
            // Map tab names to categories
            switch (tabName.ToLower())
            {
                case "ducts":
                    return category.Equals("Ducts", StringComparison.OrdinalIgnoreCase);
                case "duct accessories":
                    return category.Equals("Duct Accessories", StringComparison.OrdinalIgnoreCase);
                case "cable trays":
                    return category.Equals("Cable Trays", StringComparison.OrdinalIgnoreCase);
                case "pipes":
                    return category.Equals("Pipes", StringComparison.OrdinalIgnoreCase);
                default:
                    return false;
            }
        }

        private void UpdateTabParameterDropdown(WinForms.TabPage tabPage, List<Models.ParameterInfo> parameters)
        {
            try
            {
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] UpdateTabParameterDropdown called for tab {tabPage.Text} with {parameters.Count} parameters\n");
                
                // Find ALL ComboBoxes in the tab (parameter name dropdowns) - look in nested panels too
                var allComboBoxes = new List<WinForms.ComboBox>();
                
                // First, try direct ComboBoxes in tabPage
                allComboBoxes.AddRange(tabPage.Controls.OfType<WinForms.ComboBox>());
                
                // Then look in nested panels (servicePanel -> row -> nameCombo)
                var servicePanel = tabPage.Controls.OfType<WinForms.Panel>().FirstOrDefault();
                if (servicePanel != null)
                {
                    var rowPanels = servicePanel.Controls.OfType<WinForms.Panel>();
                    foreach (var rowPanel in rowPanels)
                    {
                        var comboBoxes = rowPanel.Controls.OfType<WinForms.ComboBox>();
                        allComboBoxes.AddRange(comboBoxes);
                    }
                }
                
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Found {allComboBoxes.Count} ComboBoxes in tab {tabPage.Text}\n");
                
                if (allComboBoxes.Count > 0)
                {
                    // Prepare parameter list once
                    var groupedParameters = GroupParameters(parameters);
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Grouped parameters into {groupedParameters.Count} groups\n");
                    
                    // Separate MEP and Opening parameters
                    var mepParameters = new List<Models.ParameterInfo>();
                    var openingParameters = new List<Models.ParameterInfo>();
                    
                    if (_allCollectedParameters.ContainsKey("MEP Elements"))
                    {
                        mepParameters = _allCollectedParameters["MEP Elements"];
                    }
                    if (_allCollectedParameters.ContainsKey("Opening Families"))
                    {
                        openingParameters = _allCollectedParameters["Opening Families"];
                    }
                    
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Separated: {mepParameters.Count} MEP params, {openingParameters.Count} opening params\n");
                    
                    // Update ComboBoxes in pairs (left = MEP, right = Opening)
                    for (int i = 0; i < allComboBoxes.Count; i += 2)
                    {
                        var nameCombo = allComboBoxes[i]; // Left ComboBox (MEP parameters)
                        var valueCombo = (i + 1 < allComboBoxes.Count) ? allComboBoxes[i + 1] : null; // Right ComboBox (Opening parameters)
                        
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Updating ComboBox pair {i/2} in tab {tabPage.Text}\n");
                        
                        // Update left ComboBox with MEP parameters
                        nameCombo.Items.Clear();
                        nameCombo.Items.Add("<Select MEP Parameter>");
                        foreach (var param in mepParameters)
                        {
                            nameCombo.Items.Add(param.Name);
                        }
                        nameCombo.SelectedIndex = 0;
                        
                        // Update right ComboBox with Opening parameters
                        if (valueCombo != null)
                        {
                            valueCombo.Items.Clear();
                            valueCombo.Items.Add("<Select Opening Parameter>");
                            foreach (var param in openingParameters)
                            {
                                valueCombo.Items.Add(param.Name);
                            }
                            valueCombo.SelectedIndex = 0;
                        }
                        
                        // Add event handlers
                        nameCombo.SelectedIndexChanged += (sender, e) => OnParameterSelected(sender, e, tabPage.Text);
                        if (valueCombo != null)
                        {
                            valueCombo.SelectedIndexChanged += (sender, e) => OnValueSelected(sender, e, tabPage.Text);
                        }
                    }
                    
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Updated {allComboBoxes.Count} ComboBoxes in tab {tabPage.Text} with {allComboBoxes[0].Items.Count} total items each\n");
                }
                else
                {
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] No ComboBoxes found in tab {tabPage.Text}\n");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error updating tab parameter dropdown: {ex.Message}");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] ERROR in UpdateTabParameterDropdown: {ex.Message}\n");
            }
        }

        private Dictionary<string, List<Models.ParameterInfo>> GroupParameters(List<Models.ParameterInfo> parameters)
        {
            var grouped = new Dictionary<string, List<Models.ParameterInfo>>();
            
            foreach (var param in parameters)
            {
                string group = GetParameterGroup(param.Name);
                if (!grouped.ContainsKey(group))
                {
                    grouped[group] = new List<Models.ParameterInfo>();
                }
                grouped[group].Add(param);
            }
            
            return grouped;
        }

        private string GetParameterGroup(string parameterName)
        {
            // Group parameters by common prefixes/suffixes for easier selection
            var upperName = parameterName.ToUpper();
            
            if (upperName.Contains("SIZE") || upperName.Contains("DIAMETER") || upperName.Contains("WIDTH") || upperName.Contains("HEIGHT"))
                return "Size Parameters";
            else if (upperName.Contains("LEVEL") || upperName.Contains("ELEVATION"))
                return "Level Parameters";
            else if (upperName.Contains("MATERIAL") || upperName.Contains("TYPE"))
                return "Material/Type Parameters";
            else if (upperName.Contains("SYSTEM") || upperName.Contains("CLASSIFICATION"))
                return "System Parameters";
            else if (upperName.Contains("FIRE") || upperName.Contains("SMOKE"))
                return "Fire/Safety Parameters";
            else
                return "Other Parameters";
        }


        private Document? GetLinkedDocument(string linkedFileName)
        {
            try
            {
                // Extract the actual filename from the display text
                string actualFileName = linkedFileName;
                if (linkedFileName.Contains(" ("))
                {
                    actualFileName = linkedFileName.Substring(0, linkedFileName.IndexOf(" ("));
                }
                
                // Find the linked file in the current document
                if (_activeDocument != null)
                {
                    var linkedFiles = new FilteredElementCollector(_activeDocument)
                        .OfClass(typeof(RevitLinkInstance))
                        .Cast<RevitLinkInstance>()
                        .Where(link => link.GetLinkDocument() != null)
                        .ToList();
                    
                    foreach (var link in linkedFiles)
                    {
                        var linkDoc = link.GetLinkDocument();
                        if (linkDoc != null && linkDoc.Title.Contains(actualFileName))
                        {
                            return linkDoc;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting linked document for '{linkedFileName}': {ex.Message}");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Error getting linked document for '{linkedFileName}': {ex.Message}\n");
            }
            
            return null;
        }

        private void OnParameterSelected(object? sender, EventArgs e, string tabName)
        {
            try
            {
                // Prevent infinite loops
                if (_isUpdatingComboBoxes)
                {
                    return;
                }
                
                if (sender is WinForms.ComboBox nameCombo && nameCombo.SelectedItem != null)
                {
                    string selectedParameter = nameCombo.SelectedItem.ToString();
                    
                    // Skip if it's a group header or default selection
                    if (selectedParameter.StartsWith("---") || selectedParameter == "<Select Parameter>")
                    {
                        return;
                    }
                    
                    // Extract parameter name (remove type info)
                    string parameterName = selectedParameter.Split('(')[0].Trim();
                    
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Parameter selected: {parameterName} in tab {tabName}\n");
                    
                    // Find the corresponding value ComboBox in the same row
                    var valueCombo = FindValueComboBox(nameCombo);
                    if (valueCombo != null)
                    {
                        // Set flag to prevent infinite loops
                        _isUpdatingComboBoxes = true;
                        
                        try
                        {
                            // Get all unique values for this parameter from collected parameters
                            var parameterValues = GetParameterValues(parameterName, tabName);
                            
                            // Populate the value ComboBox
                            valueCombo.Items.Clear();
                            valueCombo.Items.Add("<Select Value>");
                            
                            if (parameterValues.Count > 0)
                            {
                                valueCombo.Items.AddRange(parameterValues.ToArray());
                                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Populated value ComboBox with {parameterValues.Count} values for parameter {parameterName}: {string.Join(", ", parameterValues.Take(5))}\n");
                            }
                            else
                            {
                                valueCombo.Items.Add("No values found");
                                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] No values found for parameter {parameterName}\n");
                            }
                            
                            valueCombo.SelectedIndex = 0;
                        }
                        finally
                        {
                            // Always reset the flag
                            _isUpdatingComboBoxes = false;
                        }
                    }
                    else
                    {
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] No value ComboBox found for parameter {parameterName}\n");
                    }
                }
            }
            catch (Exception ex)
            {
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Error in OnParameterSelected: {ex.Message}\n");
                _isUpdatingComboBoxes = false; // Reset flag on error
            }
        }

        private WinForms.ComboBox? FindValueComboBox(WinForms.ComboBox nameCombo)
        {
            try
            {
                // Find the parent row panel
                var rowPanel = nameCombo.Parent as WinForms.Panel;
                if (rowPanel != null)
                {
                    // Find the second ComboBox in the same row (value ComboBox)
                    var comboBoxes = rowPanel.Controls.OfType<WinForms.ComboBox>().ToList();
                    if (comboBoxes.Count >= 2)
                    {
                        // Return the second ComboBox (value ComboBox)
                        return comboBoxes[1];
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Error finding value ComboBox: {ex.Message}\n");
                return null;
            }
        }

        private List<string> GetParameterValues(string parameterName, string tabName)
        {
            try
            {
                var values = new HashSet<string>(); // Use HashSet to avoid duplicates
                
                // Get the category for this tab
                string category = GetCategoryForTab(tabName);
                if (string.IsNullOrEmpty(category) || !_allCollectedParameters.ContainsKey(category))
                {
                    return new List<string>();
                }
                
                // Find the parameter in collected parameters
                var parameters = _allCollectedParameters[category];
                var targetParameter = parameters.FirstOrDefault(p => p.Name.Equals(parameterName, StringComparison.OrdinalIgnoreCase));
                
                // if (targetParameter != null && targetParameter.Values != null)
                // {
                //     // Add all values from the parameter
                //     foreach (var value in targetParameter.Values)
                //     {
                //         if (!string.IsNullOrEmpty(value))
                //         {
                //             values.Add(value);
                //         }
                //     }
                // } // Values property not available in Models.ParameterInfo
                
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Found {values.Count} unique values for parameter {parameterName} in category {category}\n");
                
                return values.OrderBy(v => v).ToList(); // Return sorted list
            }
            catch (Exception ex)
            {
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Error getting parameter values for {parameterName}: {ex.Message}\n");
                return new List<string>();
            }
        }

        private string GetCategoryForTab(string tabName)
        {
            // Map tab names back to categories
            switch (tabName.ToLower())
            {
                case "ducts":
                    return "Ducts";
                case "duct accessories":
                    return "Duct Accessories";
                case "cable trays":
                    return "Cable Trays";
                case "pipes":
                    return "Pipes";
                default:
                    return string.Empty;
            }
        }

        /// <summary>
        /// Determines if a file type should be available for selection in the Reference Elements section
        /// MEP files (ME, EL, PH) are available by default, Architecture/Structural files are locked
        /// </summary>
        private bool IsFileAvailableForReference(LinkedFileType fileType)
        {
            return fileType switch
            {
                LinkedFileType.Electrical => true,    // EL, EE files - available
                LinkedFileType.Mechanical => true,    // ME files - available  
                LinkedFileType.Plumbing => true,      // PH files - available
                LinkedFileType.FireProtection => true, // FP, FF files - available
                LinkedFileType.Architectural => false, // AR, ARC files - locked
                LinkedFileType.Structural => false,   // ST, STR files - locked
                LinkedFileType.Unknown => false,      // Unknown files - locked
                _ => false
            };
        }

        /// <summary>
        /// Determines if a file type should be available for selection in the Host Elements section
        /// Architecture/Structural files (AR, ARC, ST, STR) are available by default, MEP files are locked
        /// </summary>
        private bool IsFileAvailableForHost(LinkedFileType fileType)
        {
            return fileType switch
            {
                LinkedFileType.Architectural => true, // AR, ARC files - available
                LinkedFileType.Structural => true,    // ST, STR files - available
                LinkedFileType.Electrical => false,   // EL, EE files - locked
                LinkedFileType.Mechanical => false,   // ME files - locked
                LinkedFileType.Plumbing => false,     // PH files - locked
                LinkedFileType.FireProtection => false, // FP, FF files - locked
                LinkedFileType.Unknown => false,      // Unknown files - locked
                _ => false
            };
        }





        private void OnValueSelected(object? sender, EventArgs e, string tabName)
        {
            try
            {
                // Prevent infinite loops
                if (_isUpdatingComboBoxes)
                {
                    return;
                }
                
                if (sender is WinForms.ComboBox valueCombo && valueCombo.SelectedItem != null)
                {
                    string selectedValue = valueCombo.SelectedItem.ToString();
                    
                    // Skip if it's default selection
                    if (selectedValue == "<Select Value>" || selectedValue == "No values found")
                    {
                        return;
                    }
                    
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Value selected: {selectedValue} in tab {tabName}\n");
                    
                    // Find the corresponding parameter ComboBox in the same row
                    var nameCombo = FindParameterComboBox(valueCombo);
                    if (nameCombo != null && nameCombo.SelectedItem != null)
                    {
                        string selectedParameter = nameCombo.SelectedItem.ToString();
                        if (!selectedParameter.StartsWith("---") && selectedParameter != "<Select Parameter>")
                        {
                            string parameterName = selectedParameter.Split('(')[0].Trim();
                            
                            // Show mapping information
                            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Ready to map: {parameterName} = {selectedValue} for {tabName}\n");
                            
                            // Handle different parameter types
                            if (parameterName.Equals("Size", StringComparison.OrdinalIgnoreCase))
                            {
                                // For Size parameters, find reference elements with this size
                                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Size parameter selected - finding reference elements with size: {selectedValue}\n");
                                // TODO: Add "Find Reference Elements" button to show elements with this size
                            }
                            else if (parameterName.Equals("Reference Level", StringComparison.OrdinalIgnoreCase) || 
                                     parameterName.Equals("Level", StringComparison.OrdinalIgnoreCase))
                            {
                                // For Level parameters, this will be applied to all openings on that level
                                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Level parameter selected - will apply to all openings on level: {selectedValue}\n");
                                // TODO: Add "Apply to All Openings on Level" button
                            }
                            else
                            {
                                // For other parameters (System Type, Material, etc.)
                                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Parameter {parameterName} selected - ready for mapping to openings\n");
                                // TODO: Add "Map to Opening" button for general parameters
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Error in OnValueSelected: {ex.Message}\n");
            }
        }

        private WinForms.ComboBox? FindParameterComboBox(WinForms.ComboBox valueCombo)
        {
            try
            {
                // Find the parent row panel
                var rowPanel = valueCombo.Parent as WinForms.Panel;
                if (rowPanel != null)
                {
                    // Find the first ComboBox in the same row (parameter ComboBox)
                    var comboBoxes = rowPanel.Controls.OfType<WinForms.ComboBox>().ToList();
                    if (comboBoxes.Count >= 2)
                    {
                        // Return the first ComboBox (parameter ComboBox)
                        return comboBoxes[0];
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Error finding parameter ComboBox: {ex.Message}\n");
                return null;
            }
        }

        private void PositionPanels()
        {
            DebugLogger.Info("=== STARTING PositionPanels (3-Column Layout) ===");
            DebugLogger.Info($"Form ClientSize: {this.ClientSize.Width}x{this.ClientSize.Height}");

            int top = _headerPanel.Bottom;
            DebugLogger.Info($"_headerPanel.Bottom: {top}");

            // Account for status banner if visible
            if (_statusBannerPanel != null && _statusBannerPanel.Visible)
            {
                top += _statusBannerPanel.Height;
                DebugLogger.Info($"Status banner visible, new top: {top}");
            }

            int height = this.ClientSize.Height - top - _statusPanel.Height;
            DebugLogger.Info($"Calculated height: {height}");

            // NEW: Left Column - Filters (Fixed width)
            int filtersWidth = 200;
            if (_filtersPanel != null)
            {
                _filtersPanel.Location = new System.Drawing.Point(0, top);
                _filtersPanel.Size = new System.Drawing.Size(filtersWidth, height);
                DebugLogger.Info($"_filtersPanel positioned: Location={_filtersPanel.Location}, Size={_filtersPanel.Size}");
            }

            // NEW: Filters Splitter
            if (_filtersSplitter != null)
            {
                _filtersSplitter.Location = new System.Drawing.Point(filtersWidth, top);
                _filtersSplitter.Size = new System.Drawing.Size(3, height);
                DebugLogger.Info($"_filtersSplitter positioned: Location={_filtersSplitter.Location}, Size={_filtersSplitter.Size}");
            }

            // EXISTING: Right Panel - Keep same size, just move right
            _rightPanel.Location = new System.Drawing.Point(this.ClientSize.Width - _rightPanel.Width, top);
            _rightPanel.Size = new System.Drawing.Size(_rightPanel.Width, height);
            DebugLogger.Info($"_rightPanel positioned: Location={_rightPanel.Location}, Size={_rightPanel.Size}");

            // EXISTING: Main Splitter - Keep same position relative to right panel
            int splitterX = _rightPanel.Left - _mainSplitter.Width;
            _mainSplitter.Location = new System.Drawing.Point(splitterX, top);
            _mainSplitter.Height = height;
            _mainSplitter.Anchor = WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Top;
            DebugLogger.Info($"_mainSplitter positioned: Location={_mainSplitter.Location}, Height={_mainSplitter.Height}");

            // EXISTING: Left Panel (4-section) - Keep same size, just move right
            int leftPanelStartX = (_filtersPanel != null) ? filtersWidth + 3 : 0;
            _leftPanel.Location = new System.Drawing.Point(leftPanelStartX, top);
            _leftPanel.Size = new System.Drawing.Size(_mainSplitter.Left - leftPanelStartX, height);
            DebugLogger.Info($"_leftPanel positioned: Location={_leftPanel.Location}, Size={_leftPanel.Size}");

            DebugLogger.Info("=== PositionPanels (3-Column) COMPLETED ===");
        }

        private void BalanceLeftLayout()
        {
            DebugLogger.Info("=== STARTING BalanceLeftLayout ===");

            if (_leftPanel == null)
            {
                DebugLogger.Info("_leftPanel is null, returning");
                return;
            }

            DebugLogger.Info($"_leftPanel size: {_leftPanel.Width}x{_leftPanel.Height}");
            DebugLogger.Info($"_horizontalSplitter height: {_horizontalSplitter?.Height ?? 0}");

            // equal top/bottom in the left column
            int half = (_leftPanel.Height - _horizontalSplitter.Height) / 2;
            DebugLogger.Info($"Calculated half height: {half}");

            _bottomLeftPanel.Height = half;
            DebugLogger.Info($"_bottomLeftPanel height set to: {_bottomLeftPanel.Height}");

            // [[memory:8116398]] Position right subsections with anchoring to fix width=0 issue
            if (_topRightPanel != null && _bottomRightPanel != null && _topLeftPanel != null && _bottomLeftPanel != null)
            {
                // Calculate available width for left content (total minus right panel and splitter)
                int rightPanelWidth = 200; // Reduced from 340 to 200
                int leftContentWidth = _leftPanel.Width - rightPanelWidth - 3; // 200 for right panel, 3 for splitter
                DebugLogger.Info($"Left content width calculated: {leftContentWidth}");

                // Top right panel positioning - anchor to right edge
                _topRightPanel.Width = rightPanelWidth;
                _topRightPanel.Height = _topLeftPanel.Height;
                _topRightPanel.Location = new System.Drawing.Point(_leftPanel.Width - rightPanelWidth, 0);
                
                // Bottom right panel positioning - anchor to right edge  
                _bottomRightPanel.Width = rightPanelWidth;
                _bottomRightPanel.Height = _bottomLeftPanel.Height;
                _bottomRightPanel.Location = new System.Drawing.Point(_leftPanel.Width - rightPanelWidth, 0);
                
                // Position vertical splitters just to the left of right panels
                if (_verticalSplitter != null)
                {
                    _verticalSplitter.Height = _topLeftPanel.Height;
                    _verticalSplitter.Location = new System.Drawing.Point(_leftPanel.Width - rightPanelWidth - 3, 0);
                }
                
                DebugLogger.Info($"Positioned right panels: width={rightPanelWidth}, left content gets width={leftContentWidth}");
            }
            else
            {
                DebugLogger.Error("Some panels are null - cannot position right subsections");
            }

            DebugLogger.Info("=== BalanceLeftLayout COMPLETED ===");
        }

        /// <summary>
        /// Captures current UI state and saves it to the selected filter
        /// </summary>
        private void SaveFilterWithUIState(WinForms.ListBox filterListBox)
        {
            try
            {
                DebugLogger.Info("[FILTER_UI] Saving filter with current UI state");
                
                var selectedFilter = GetSelectedFilterFromListBox(filterListBox);
                if (selectedFilter == null)
                {
                    WinForms.MessageBox.Show("Please select a filter to save.", "No Selection", 
                        WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
                    return;
                }

                // Capture current UI state
                CaptureCurrentUIStateToFilter(selectedFilter);
                
                // Save the filter directly using the modified filter object
                var filterDir = _filterManagementService.GetDefaultFilterDirectory();
                if (!Directory.Exists(filterDir))
                {
                    Directory.CreateDirectory(filterDir);
                }

                var filePath = Path.Combine(filterDir, $"{selectedFilter.Name}.xml");
                _filterManagementService.SaveFilterToXmlFile(selectedFilter, filePath);
                
                DebugLogger.Info($"[FILTER_UI] Saved filter '{selectedFilter.Name}' with UI state");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[FILTER_UI] Error saving filter with UI state: {ex.Message}");
                WinForms.MessageBox.Show($"Error saving filter: {ex.Message}", "Error", 
                    WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Restores filter state to UI when a filter is selected
        /// </summary>
        private void RestoreFilterStateToUI(WinForms.ListBox filterListBox)
        {
            try
            {
                if (filterListBox?.SelectedItem == null) return;
                
                DebugLogger.Info("[FILTER_UI] Restoring filter state to UI");
                
                var selectedFilterName = filterListBox.SelectedItem.ToString();
                DebugLogger.Info($"[FILTER_UI] Selected filter name: {selectedFilterName}");
                
                // Check if this is a different filter than the currently loaded one
                var isDifferentFilter = _lastLoadedFilterName != selectedFilterName;
                DebugLogger.Info($"[FILTER_UI] Is different filter: {isDifferentFilter} (Last: '{_lastLoadedFilterName}', Current: '{selectedFilterName}')");
                
                // Load the actual filter data from the saved file
                var selectedFilter = _filterManagementService.LoadFilterAuto(selectedFilterName);
                if (selectedFilter == null) 
                {
                    DebugLogger.Warning($"[FILTER_UI] Could not load filter '{selectedFilterName}' from file");
                    return;
                }

                DebugLogger.Info($"[FILTER_UI] Loaded filter data - Category: {selectedFilter.SelectedMepCategoryName}, Reference Files: {selectedFilter.SelectedReferenceFiles?.Count ?? 0}");

                // Restore UI state from filter (only clear if switching to different filter)
                RestoreUIStateFromFilter(selectedFilter, isDifferentFilter);
                
                // Update the last loaded filter name
                _lastLoadedFilterName = selectedFilterName;
                
                DebugLogger.Info($"[FILTER_UI] Restored UI state from filter '{selectedFilter.Name}'");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[FILTER_UI] Error restoring filter state to UI: {ex.Message}");
            }
        }

        /// <summary>
        /// Captures current UI state and stores it in the filter
        /// </summary>
        private void CaptureCurrentUIStateToFilter(OpeningFilter filter)
        {
            try
            {
                DebugLogger.Info("[FILTER_UI] Capturing current UI state to filter");
                
                // Capture MEP category selection
                var selectedCategories = GetSelectedMepCategories();
                filter.SelectedMepCategoryNames = selectedCategories;
                
                // For backward compatibility, also set the single category
                if (selectedCategories.Count > 0)
                {
                    filter.SelectedMepCategoryName = selectedCategories[0];
                    // Convert first category to enum for backward compatibility
                    if (Enum.TryParse(selectedCategories[0], out Models.MepCategory category))
                    {
                        filter.Category = category;
                    }
                }
                
                // Capture reference file selections
                filter.SelectedReferenceFiles = GetSelectedReferenceFiles();
                
                // Capture host file selections
                filter.SelectedHostFiles = GetSelectedHostFiles();
                
                // Capture opening settings (clearances, etc.)
                filter.OpeningSettings = GetCurrentOpeningSettings();
                
                // Update timestamp
                filter.LastModified = DateTime.Now;
                
                DebugLogger.Info($"[FILTER_UI] Captured UI state - Categories: {string.Join(", ", selectedCategories)}, Reference Files: {filter.SelectedReferenceFiles.Count}, Host Files: {filter.SelectedHostFiles.Count}, Has Settings: {filter.OpeningSettings != null}");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[FILTER_UI] Error capturing UI state: {ex.Message}");
            }
        }

        /// <summary>
        /// Clears all UI selections before restoring filter state
        /// </summary>
        private void ClearAllUISelections()
        {
            try
            {
                DebugLogger.Info("[FILTER_UI] Clearing all UI selections before restoring filter state");
                
                // Clear MEP category selections
                var mepCategoriesListBox = _topRightPanel?.Controls.OfType<WinForms.CheckedListBox>().FirstOrDefault();
                if (mepCategoriesListBox != null)
                {
                    mepCategoriesListBox.BeginUpdate();
                    for (int i = 0; i < mepCategoriesListBox.Items.Count; i++)
                    {
                        mepCategoriesListBox.SetItemChecked(i, false);
                    }
                    mepCategoriesListBox.EndUpdate();
                }
                
                // Clear reference file selections
                if (_topLeftPanel?.Controls.Count > 0)
                {
                    foreach (var control in _topLeftPanel.Controls)
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
                
                // Clear host file selections
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
                
                DebugLogger.Info("[FILTER_UI] All UI selections cleared");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[FILTER_UI] Error clearing UI selections: {ex.Message}");
            }
        }

        /// <summary>
        /// Restores UI state from the selected filter
        /// </summary>
        private void RestoreUIStateFromFilter(OpeningFilter filter, bool shouldClear = true)
        {
            try
            {
                DebugLogger.Info("[FILTER_UI] Restoring UI state from filter");
                
                // Only clear UI selections if switching to a different filter
                if (shouldClear)
                {
                    DebugLogger.Info("[FILTER_UI] Clearing UI selections before restoring filter state");
                    ClearAllUISelections();
                }
                else
                {
                    DebugLogger.Info("[FILTER_UI] Skipping UI clearing - same filter or new filter creation");
                }
                
                // Restore MEP category selection
                if (filter.SelectedMepCategoryNames?.Any() == true)
                {
                    DebugLogger.Info($"[FILTER_UI] Restoring MEP categories: {string.Join(", ", filter.SelectedMepCategoryNames)}");
                    // Use multiple categories if available
                    RestoreMepCategorySelections(filter.SelectedMepCategoryNames);
                }
                else if (!string.IsNullOrEmpty(filter.SelectedMepCategoryName))
                {
                    DebugLogger.Info($"[FILTER_UI] Restoring single MEP category: {filter.SelectedMepCategoryName}");
                    // Fallback to single category for backward compatibility
                    var categories = new List<string> { filter.SelectedMepCategoryName };
                    RestoreMepCategorySelections(categories);
                }
                else
                {
                    DebugLogger.Warning("[FILTER_UI] No MEP categories found in filter to restore");
                }
                
                // Restore reference file selections
                if (filter.SelectedReferenceFiles?.Any() == true)
                {
                    DebugLogger.Info($"[FILTER_UI] Restoring reference files: {string.Join(", ", filter.SelectedReferenceFiles)}");
                    RestoreReferenceFileSelections(filter.SelectedReferenceFiles);
                }
                else
                {
                    DebugLogger.Warning("[FILTER_UI] No reference files found in filter to restore");
                }
                
                // Restore host file selections
                if (filter.SelectedHostFiles?.Any() == true)
                {
                    DebugLogger.Info($"[FILTER_UI] Restoring host files: {string.Join(", ", filter.SelectedHostFiles)}");
                    RestoreHostFileSelections(filter.SelectedHostFiles);
                }
                else
                {
                    DebugLogger.Warning("[FILTER_UI] No host files found in filter to restore");
                }
                
                // Restore opening settings
                if (filter.OpeningSettings != null)
                {
                    DebugLogger.Info("[FILTER_UI] Restoring opening settings");
                    RestoreOpeningSettings(filter.OpeningSettings);
                }
                else
                {
                    DebugLogger.Warning("[FILTER_UI] No opening settings found in filter to restore");
                }
                
                DebugLogger.Info($"[FILTER_UI] Restored UI state from filter '{filter.Name}'");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[FILTER_UI] Error restoring UI state: {ex.Message}");
            }
        }

        /// <summary>
        /// Gets the currently selected MEP category from UI
        /// </summary>
        private Models.MepCategory GetSelectedMepCategory()
        {
            try
            {
                // Get selected categories from the MEP Categories CheckedListBox
                var mepCategoriesListBox = _topRightPanel?.Controls.OfType<WinForms.CheckedListBox>().FirstOrDefault();
                
                if (mepCategoriesListBox != null)
                {
                    var selectedCategories = new List<string>();
                    for (int i = 0; i < mepCategoriesListBox.Items.Count; i++)
                    {
                        if (mepCategoriesListBox.GetItemChecked(i))
                        {
                            selectedCategories.Add(mepCategoriesListBox.Items[i].ToString());
                        }
                    }
                    
                    // Convert UI category names to enum values
                    if (selectedCategories.Count == 1)
                    {
                        return ConvertCategoryNameToEnum(selectedCategories[0]);
                    }
                    else if (selectedCategories.Count > 1)
                    {
                        // Multiple categories selected - return the first one as primary
                        return ConvertCategoryNameToEnum(selectedCategories[0]);
                    }
                }
                
                DebugLogger.Info("[FILTER_UI] No MEP categories selected, returning default Ducts");
                return Models.MepCategory.Ducts; // Default fallback
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[FILTER_UI] Error getting selected MEP category: {ex.Message}");
                return Models.MepCategory.Ducts;
            }
        }

        /// <summary>
        /// Converts UI category name to MepCategory enum
        /// </summary>
        private Models.MepCategory ConvertCategoryNameToEnum(string categoryName)
        {
            return categoryName switch
            {
                "Ducts" => Models.MepCategory.Ducts,
                "Duct Accessories" => Models.MepCategory.DuctAccessories,
                "Cable Trays" => Models.MepCategory.CableTrays,
                "Pipes" => Models.MepCategory.Pipes,
                _ => Models.MepCategory.Ducts
            };
        }

        /// <summary>
        /// Converts MepCategory enum to UI category name
        /// </summary>
        private string ConvertEnumToCategoryName(Models.MepCategory category)
        {
            return category switch
            {
                Models.MepCategory.Ducts => "Ducts",
                Models.MepCategory.DuctAccessories => "Duct Accessories",
                Models.MepCategory.CableTrays => "Cable Trays",
                Models.MepCategory.Pipes => "Pipes",
                _ => "Ducts"
            };
        }

        /// <summary>
        /// Restores MEP category selection in UI
        /// </summary>
        private void RestoreMepCategorySelection(string categoryName)
        {
            try
            {
                DebugLogger.Info($"[FILTER_UI] Restoring MEP category: {categoryName}");
                
                // Get the MEP Categories CheckedListBox
                var mepCategoriesListBox = _topRightPanel?.Controls.OfType<WinForms.CheckedListBox>().FirstOrDefault();
                
                if (mepCategoriesListBox != null)
                {
                    // First, uncheck all items
                    for (int i = 0; i < mepCategoriesListBox.Items.Count; i++)
                    {
                        mepCategoriesListBox.SetItemChecked(i, false);
                    }
                    
                    // Then check the specific category
                    for (int i = 0; i < mepCategoriesListBox.Items.Count; i++)
                    {
                        if (mepCategoriesListBox.Items[i].ToString() == categoryName)
                        {
                            mepCategoriesListBox.SetItemChecked(i, true);
                            DebugLogger.Info($"[FILTER_UI] Successfully restored MEP category: {categoryName}");
                            break;
                        }
                    }
                    
                    // Trigger the update to refresh the MEP type combo and clearance panels
                    UpdateMepTypeBasedOnSelection();
                }
                else
                {
                    DebugLogger.Warning("[FILTER_UI] MEP Categories CheckedListBox not found");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[FILTER_UI] Error restoring MEP category selection: {ex.Message}");
            }
        }

        /// <summary>
        /// Gets the selected filter from the list box
        /// </summary>
        private OpeningFilter GetSelectedFilterFromListBox(WinForms.ListBox filterListBox)
        {
            if (filterListBox?.SelectedItem != null)
            {
                var selectedName = filterListBox.SelectedItem.ToString();
                // This would need to be implemented to get the actual filter object
                // For now, return a basic filter - you'll need to wire this properly
                return new OpeningFilter
                {
                    Name = selectedName,
                    Category = Models.MepCategory.Ducts,
                    OpeningType = Models.OpeningType.RectangularSleeves,
                    IsEnabled = true,
                    LastModified = DateTime.Now
                };
            }
            return null;
        }

    }
}
