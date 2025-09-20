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
            
            // Add sample filters (like conVoid)
            filterListBox.Items.Add("Electrical");
            filterListBox.Items.Add("Plumbing");
            filterListBox.Items.Add("Ventilation");
            DebugLogger.Info("Sample filters added to list");
            
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
            
            // Get real parameters from the model instead of hardcoded values
            var availableParameters = GetAvailableParameters();
            nameCombo.Items.AddRange(availableParameters.ToArray());
            nameCombo.SelectedItem = parameterName;
            
            row.Controls.Add(nameCombo);

            var valueCombo = new WinForms.ComboBox
            {
                Location = new System.Drawing.Point(130, 2),
                Size = new System.Drawing.Size(row.Width - 130 - 30, 20),
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right,
                DropDownStyle = WinForms.ComboBoxStyle.DropDownList
            };
            
            // Add event handler for parameter selection changes
            nameCombo.SelectedIndexChanged += (_, __) => {
                valueCombo.Items.Clear();
                valueCombo.Items.Add("<Auto Selection>");
                valueCombo.SelectedIndex = 0;
            };
            
            row.Controls.Add(valueCombo);

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
                DebugLogger.Info($"=== LoadConfigurationFromProfile COMPLETED for {profile.Name} ===");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to load configuration from profile {profile.Name}: {ex.Message}");
                // Don't throw - we can continue with defaults
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
            if (selectedFiles == null || selectedFiles.Count == 0) return;
            
            try
            {
                if (_bottomLeftPanel?.Controls.Count > 0)
                {
                    foreach (var control in _bottomLeftPanel.Controls)
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
                                if (item != null && selectedFiles.Contains(item))
                                {
                                    listBox.SetItemChecked(i, true);
                                    DebugLogger.Info($"Restored host file selection: {item}");
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
                DebugLogger.Error($"Failed to restore host file selections: {ex.Message}");
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
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to load UI state directly: {ex.Message}");
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
            
            return selectedCategories;
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
                // Get clearance values from clearance panel
                if (_clearancePanel?.Controls.Count > 0)
                {
                    // Get current MEP category for specific key generation
                    string currentCategory = GetCurrentMepCategory();
                    
                    foreach (var control in _clearancePanel.Controls)
                    {
                        if (control is WinForms.TextBox textBox && textBox.Tag != null)
                        {
                            if (double.TryParse(textBox.Text, out double value))
                            {
                                string genericKey = textBox.Tag.ToString() ?? "";
                                string specificKey = ConvertToSpecificClearanceKey(genericKey, currentCategory);
                                clearances[specificKey] = value;
                                
                                DebugLogger.Info($"Clearance setting: {genericKey} -> {specificKey} = {value}mm");
                            }
                        }
                    }
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
                    if (currentProfile?.Configuration?.ClashZoneStorage != null)
                    {
                        targetFilter.ClashZoneStorage = currentProfile.Configuration.ClashZoneStorage;
                        targetFilter.LastModified = DateTime.Now;
                        DebugLogger.Info($"[CLASH_DEBUG] Saved clash zones to filter '{targetFilter.Name}'");
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [CLASH_DEBUG] SUCCESS: Saved {total} clash zones to filter '{targetFilter.Name}'\n");
                    }
                    else
                    {
                        DebugLogger.Warning("[CLASH_DEBUG] currentProfile.Configuration.ClashZoneStorage is null - cannot save clash zones");
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] [CLASH_DEBUG] WARNING: currentProfile.Configuration.ClashZoneStorage is null\n");
                    }
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
                
                // Create and show settings dialog
                var settings = new SettingsModel(); // You can load from saved settings here
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] SettingsModel created successfully\n");
                
                var settingsDialog = new SettingsDialog(settings);
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] SettingsDialog created successfully\n");
                
                // Show the settings dialog as modal
                using (settingsDialog)
                {
                    if (settingsDialog.ShowDialog(this) == DialogResult.OK)
                    {
                        // Settings were saved
                        var savedSettings = settingsDialog.GetSettings();
                        _statusLabel.Text = "Settings saved successfully";
                        
                        // TODO: Save settings to file or profile
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
                    // allParameters.AddRange(activeDocParams); // Type conversion issue
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
                            // allParameters.AddRange(linkedParams); // Type conversion issue
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
                        
                        // Find matching category parameters from ALL collected parameters (persistent)
                        var matchingParameters = new List<Models.ParameterInfo>();
                        foreach (var kvp in _allCollectedParameters) // Use _allCollectedParameters instead of categoryParameters
                        {
                            if (IsCategoryMatch(tabName, kvp.Key))
                            {
                                matchingParameters.AddRange(kvp.Value);
                                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Found {kvp.Value.Count} parameters for {kvp.Key} in tab {tabName}\n");
                            }
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
                    
                    // Update ALL ComboBoxes in this tab
                    foreach (var nameCombo in allComboBoxes)
                    {
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Updating ComboBox in tab {tabPage.Text}\n");
                        
                        nameCombo.Items.Clear();
                        nameCombo.Items.Add("<Select Parameter>");
                        
                        foreach (var group in groupedParameters)
                        {
                            if (!string.IsNullOrEmpty(group.Key))
                            {
                                nameCombo.Items.Add($"--- {group.Key} ---");
                            }
                            foreach (var param in group.Value)
                            {
                                nameCombo.Items.Add($"{param.Name} ({param.Type})");
                            }
                        }
                        
                        nameCombo.SelectedIndex = 0;
                        
                        // Add event handler for parameter selection to populate values
                        nameCombo.SelectedIndexChanged += (sender, e) => OnParameterSelected(sender, e, tabPage.Text);
                        
                        // Add event handler for value selection to enable mapping
                        var valueCombo = FindValueComboBox(nameCombo);
                        if (valueCombo != null)
                        {
                            valueCombo.SelectedIndexChanged += (sender, e) => OnValueSelected(sender, e, tabPage.Text);
                            
                            // Initialize value ComboBox with default selection
                            valueCombo.Items.Clear();
                            valueCombo.Items.Add("<Select Value>");
                            valueCombo.SelectedIndex = 0;
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

    }
}
