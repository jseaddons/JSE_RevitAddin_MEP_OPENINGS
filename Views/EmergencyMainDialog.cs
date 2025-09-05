using System;
using System.Drawing;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Models;
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
        
        // Main panels - 4-section layout
        private WinForms.Panel _leftPanel = null!;
        private WinForms.Panel _rightPanel = null!;
        private WinForms.Panel _mainSplitter = null!;  // Using Panel instead of Splitter to avoid docking requirement
        
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
        
        // Dynamic UI controls (only what we actually use)
        private WinForms.ComboBox _mepTypeCombo = null!;
        private WinForms.Panel _clearancePanel = null!;
        private WinForms.Panel _cableTrayPanel = null!;
        private WinForms.Panel _damperPanel = null!;
        private LinkedFileService? _linkedFileService;
        private List<LinkedFileInfo> _linkedFiles = new List<LinkedFileInfo>();
        private Document? _activeDocument;
        // Opening type controls (right section)
        private WinForms.Panel _openingTypePanel = null!;
        private WinForms.RadioButton _rectangularRadio = null!;
        private WinForms.RadioButton _circularRadio = null!;
        // Parameter filter controls (right section)
        private WinForms.Panel _parameterFilterPanel = null!;
        private List<WinForms.Panel> _parameterRows = new List<WinForms.Panel>();
        private WinForms.Button _addParameterButton = null!;

        // constants (top of class)
        private const int InnerRightWidth = 320;  // choose 300–360

        public EmergencyMainDialog(ApplicationProfileService appProfileService, Document? document = null)
        {
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
                Text = "⚠️ No linked files found. Please link MEP, architectural, and structural files to use this feature.",
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

            // Status Panel (bottom)
            _statusPanel = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.Bottom,
                Height = 30,
                BackColor = System.Drawing.Color.FromArgb(248, 249, 250),
                BorderStyle = WinForms.BorderStyle.FixedSingle
            };
            this.Controls.Add(_statusPanel);

            _statusLabel = new WinForms.Label
            {
                Text = "Ready",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular),
                ForeColor = System.Drawing.Color.FromArgb(102, 102, 102),
                Location = new System.Drawing.Point(10, 5),
                Size = new System.Drawing.Size(600, 20),
                AutoSize = false
            };
            _statusPanel.Controls.Add(_statusLabel);

            _progressBar = new WinForms.ProgressBar
            {
                Location = new System.Drawing.Point(this.Width - 220, 5),
                Size = new System.Drawing.Size(200, 20),
                Style = WinForms.ProgressBarStyle.Continuous,
                Minimum = 0,
                Maximum = 100,
                Value = 0
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
            // Create 4-section layout within left panel
            CreateFourSectionLayout();
            
            // Right Panel: Opening Configuration
            InitializeRightPanel();
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
            // Title: "Reference Elements (MEP Files)" - make it shorter
            var title = new WinForms.Label
            {
                Text = "Reference Elements",  // Remove "(MEP Files)" to make it shorter
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51),
                Location = new System.Drawing.Point(5, 2),
                Size = new System.Drawing.Size(150, 30),  // Make it wider to fit
                AutoSize = false
            };
            _topLeftPanel.Controls.Add(title);

            // ListBox for linked files with checkboxes - starts right after title
            var referenceFilesListBox = new WinForms.CheckedListBox
            {
                Location = new System.Drawing.Point(5, 40),
                Size = new System.Drawing.Size(_topLeftPanel.Width - 10, _topLeftPanel.Height - 25),
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right,
                CheckOnClick = true,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular)
            };
            _topLeftPanel.Controls.Add(referenceFilesListBox);

            // Use dynamic data only - no hardcoded fallback
            if (_linkedFiles.Count > 0 && _linkedFileService != null)
            {
                var referenceFiles = _linkedFileService.GetReferenceElementFiles(_linkedFiles);
                foreach (var file in referenceFiles)
                {
                    string displayText = $"{file.FileName} ({file.ElementCount} elements)";
                    if (!file.IsLoaded)
                    {
                        displayText += " [NOT LOADED]";
                    }
                    referenceFilesListBox.Items.Add(displayText, file.IsLoaded);
                }
                System.Diagnostics.Debug.WriteLine($"Populated top-left with {referenceFiles.Count} dynamic files");
            }
            else
            {
                // Show message that linked files need to be loaded
                var noDataLabel = new WinForms.Label
                {
                    Text = "No linked files loaded.\nPlease ensure linked files are present in the Revit project.",
                    Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Italic),
                    ForeColor = System.Drawing.Color.Gray,
                    Location = new System.Drawing.Point(10, 50),
                    Size = new System.Drawing.Size(_topLeftPanel.Width - 20, 60),
                    AutoSize = false,
                    TextAlign = System.Drawing.ContentAlignment.MiddleCenter
                };
                _topLeftPanel.Controls.Add(noDataLabel);
                System.Diagnostics.Debug.WriteLine("No linked files available - showing message to user");
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

            // Use dynamic categories only - no hardcoded fallback
            if (_linkedFiles.Count > 0 && _linkedFileService != null)
            {
                var referenceFiles = _linkedFileService.GetReferenceElementFiles(_linkedFiles);
                var availableCategories = new HashSet<string>();

                foreach (var file in referenceFiles)
                {
                    var categories = LinkedFileDetectionService.GetAvailableCategories(file.FileType);
                    foreach (var category in categories)
                    {
                        availableCategories.Add(LinkedFileDetectionService.GetCategoryDisplayName(category));
                    }
                }

                // Add available categories
                foreach (var category in availableCategories.OrderBy(c => c))
                {
                    referenceCategoriesListBox.Items.Add(category, false);
                }
                System.Diagnostics.Debug.WriteLine($"Populated MEP categories with {availableCategories.Count} dynamic categories");
            }
            else
            {
                // Show message that linked files need to be loaded
                var noDataLabel = new WinForms.Label
                {
                    Text = "No linked files loaded.\nMEP categories will be shown once linked files are available.",
                    Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Italic),
                    ForeColor = System.Drawing.Color.Gray,
                    Location = new System.Drawing.Point(10, 50),
                    Size = new System.Drawing.Size(_topRightPanel.Width - 20, 60),
                    AutoSize = false,
                    TextAlign = System.Drawing.ContentAlignment.MiddleCenter
                };
                _topRightPanel.Controls.Add(noDataLabel);
                System.Diagnostics.Debug.WriteLine("No linked files available - showing message to user for MEP categories");
            }
        }

        private void PopulateBottomLeftSection()
        {
            // Title: "Host Elements (ARC/STR Files)" - make it shorter
            var title = new WinForms.Label
            {
                Text = "Host Elements",  // Remove "(ARC/STR Files)" to make it shorter
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51),
                Location = new System.Drawing.Point(5, 2),
                Size = new System.Drawing.Size(150, 30),  // Make it wider to fit
                AutoSize = false
            };
            _bottomLeftPanel.Controls.Add(title);

            // ListBox for host files with checkboxes - starts right after title
            var hostFilesListBox = new WinForms.CheckedListBox
            {
                Location = new System.Drawing.Point(5, 30),
                Size = new System.Drawing.Size(_bottomLeftPanel.Width - 10, _bottomLeftPanel.Height - 25),
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right,
                CheckOnClick = true,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular)
            };
            _bottomLeftPanel.Controls.Add(hostFilesListBox);

            // Use dynamic data only - no hardcoded fallback
            if (_linkedFiles.Count > 0 && _linkedFileService != null)
            {
                var hostFiles = _linkedFileService.GetHostElementFiles(_linkedFiles);
                foreach (var file in hostFiles)
                {
                    string displayText = $"{file.FileName} ({file.ElementCount} elements)";
                    if (!file.IsLoaded)
                    {
                        displayText += " [NOT LOADED]";
                    }
                    hostFilesListBox.Items.Add(displayText, file.IsLoaded);
                }
                System.Diagnostics.Debug.WriteLine($"Populated bottom-left with {hostFiles.Count} dynamic host files");
            }
            else
            {
                // Show message that linked files need to be loaded
                var noDataLabel = new WinForms.Label
                {
                    Text = "No linked files loaded.\nPlease ensure architectural and structural linked files are present in the Revit project.",
                    Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Italic),
                    ForeColor = System.Drawing.Color.Gray,
                    Location = new System.Drawing.Point(10, 50),
                    Size = new System.Drawing.Size(_bottomLeftPanel.Width - 20, 60),
                    AutoSize = false,
                    TextAlign = System.Drawing.ContentAlignment.MiddleCenter
                };
                _bottomLeftPanel.Controls.Add(noDataLabel);
                System.Diagnostics.Debug.WriteLine("No linked files available - showing message to user for host elements");
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

            // Add horizontal categories based on available linked files
            if (_linkedFiles.Count > 0 && _linkedFileService != null)
            {
                var hostFiles = _linkedFileService.GetHostElementFiles(_linkedFiles);
                var horizontalCategories = new HashSet<string>();

                foreach (var file in hostFiles)
                {
                    var categories = LinkedFileDetectionService.GetAvailableHostCategories(file.FileType);
                    foreach (var category in categories)
                    {
                        var displayName = LinkedFileDetectionService.GetHostCategoryDisplayName(category);
                        if (displayName == "Walls" || displayName == "Structural Framing")
                        {
                            horizontalCategories.Add(displayName);
                        }
                    }
                }

                foreach (var category in horizontalCategories.OrderBy(c => c))
                {
                    horizontalCategoriesListBox.Items.Add(category, true);
                }
                System.Diagnostics.Debug.WriteLine($"Added {horizontalCategories.Count} horizontal host categories");
            }
            else
            {
                // Show message when no linked files
                var noDataLabel = new WinForms.Label
                {
                    Text = "No linked files.\nHorizontal openings require architectural/structural files.",
                    Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Italic),
                    ForeColor = System.Drawing.Color.Gray,
                    Location = new System.Drawing.Point(10, 40),
                    Size = new System.Drawing.Size(_bottomRightPanel.Width - 20, 40),
                    AutoSize = false,
                    TextAlign = System.Drawing.ContentAlignment.MiddleCenter
                };
                _bottomRightPanel.Controls.Add(noDataLabel);
            }

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
                    verticalCategoriesListBox.Items.Add(category, true);
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
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right
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
                Size = new System.Drawing.Size(_rightPanel.Width - 20, 50),
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
            // Hide all
            _clearancePanel.Visible = false;
            _cableTrayPanel.Visible = false;
            _damperPanel.Visible = false;

            if (mep.Equals("Cable Tray", StringComparison.OrdinalIgnoreCase))
            {
                _cableTrayPanel.Visible = true;
            }
            else if (mep.Equals("Duct Accessories", StringComparison.OrdinalIgnoreCase))
            {
                _damperPanel.Visible = true;
            }
            else
            {
                _clearancePanel.Visible = true; // default
            }
        }

        private void CreateClearancePanels()
        {
            // Standard Clearance Panel
            _clearancePanel = new WinForms.Panel
            {
                Location = new System.Drawing.Point(10, 140),
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
                Size = new System.Drawing.Size(60, 20),
                Text = "25"
            };
            _clearancePanel.Controls.Add(normalText);

            var insulatedText = new WinForms.TextBox
            {
                Location = new System.Drawing.Point(300, 33),
                Size = new System.Drawing.Size(60, 20),
                Text = "30"
            };
            _clearancePanel.Controls.Add(insulatedText);

            // Cable Tray Panel (initially hidden) - Top Side + Other Sides
            _cableTrayPanel = new WinForms.Panel
            {
                Location = new System.Drawing.Point(10, 140),
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
            var topSideTxt = new WinForms.TextBox { Location = new System.Drawing.Point(170, 38), Size = new System.Drawing.Size(60, 20), Text = "50" };
            _cableTrayPanel.Controls.Add(topSideTxt);
            var ctNormalHeader = new WinForms.Label { Text = "Normal (mm)", Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold), Location = new System.Drawing.Point(170, 18), Size = new System.Drawing.Size(110, 18) };
            _cableTrayPanel.Controls.Add(ctNormalHeader);
            // Stack 'Other Sides' below
            var otherLbl = new WinForms.Label { Text = "Other Sides:", Location = new System.Drawing.Point(10, 70), Size = new System.Drawing.Size(120, 18) };
            _cableTrayPanel.Controls.Add(otherLbl);
            var otherTxt = new WinForms.TextBox { Location = new System.Drawing.Point(170, 68), Size = new System.Drawing.Size(60, 20), Text = "25" };
            _cableTrayPanel.Controls.Add(otherTxt);
            var ctInsHeader = new WinForms.Label { Text = "Insulated (mm)", Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold), Location = new System.Drawing.Point(300, 18), Size = new System.Drawing.Size(120, 18) };
            _cableTrayPanel.Controls.Add(ctInsHeader);
            var ctInsTopTxt = new WinForms.TextBox { Location = new System.Drawing.Point(300, 38), Size = new System.Drawing.Size(60, 20), Text = "35" };
            _cableTrayPanel.Controls.Add(ctInsTopTxt);
            var ctInsOtherTxt = new WinForms.TextBox { Location = new System.Drawing.Point(300, 68), Size = new System.Drawing.Size(60, 20), Text = "25" };
            _cableTrayPanel.Controls.Add(ctInsOtherTxt);

            // Damper Panel (initially hidden) - MEP Side + Other Sides
            _damperPanel = new WinForms.Panel
            {
                Location = new System.Drawing.Point(10, 140),
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
            var mepSideTxt = new WinForms.TextBox { Location = new System.Drawing.Point(170, 38), Size = new System.Drawing.Size(60, 20), Text = "100" };
            _damperPanel.Controls.Add(mepSideTxt);
            var dNormalHeader = new WinForms.Label { Text = "Normal (mm)", Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold), Location = new System.Drawing.Point(170, 18), Size = new System.Drawing.Size(110, 18) };
            _damperPanel.Controls.Add(dNormalHeader);
            // Stack 'Other Sides' below
            var otherDLbl = new WinForms.Label { Text = "Other Sides:", Location = new System.Drawing.Point(10, 70), Size = new System.Drawing.Size(120, 18) };
            _damperPanel.Controls.Add(otherDLbl);
            var otherDTxt = new WinForms.TextBox { Location = new System.Drawing.Point(170, 68), Size = new System.Drawing.Size(60, 20), Text = "50" };
            _damperPanel.Controls.Add(otherDTxt);
            var dInsHeader = new WinForms.Label { Text = "Insulated (mm)", Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold), Location = new System.Drawing.Point(300, 18), Size = new System.Drawing.Size(120, 18) };
            _damperPanel.Controls.Add(dInsHeader);
            var dInsTopTxt = new WinForms.TextBox { Location = new System.Drawing.Point(300, 38), Size = new System.Drawing.Size(60, 20), Text = "50" };
            _damperPanel.Controls.Add(dInsTopTxt);
            var dInsOtherTxt = new WinForms.TextBox { Location = new System.Drawing.Point(300, 68), Size = new System.Drawing.Size(60, 20), Text = "35" };
            _damperPanel.Controls.Add(dInsOtherTxt);
            var dInsUnit = new WinForms.Label { Text = "mm", Location = new System.Drawing.Point(400, 41), Size = new System.Drawing.Size(30, 16) };
            _damperPanel.Controls.Add(dInsUnit);
        }

        private void CreateParameterFilterPanel()
        {
            _parameterFilterPanel = new WinForms.Panel
            {
                Location = new System.Drawing.Point(10, 270),
                Size = new System.Drawing.Size(_rightPanel.Width - 20, 160),
                BackColor = System.Drawing.Color.FromArgb(248, 249, 250),
                BorderStyle = WinForms.BorderStyle.FixedSingle,
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right
            };
            _rightPanel.Controls.Add(_parameterFilterPanel);

            var title = new WinForms.Label
            {
                Text = "Parameter Filter",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold),
                Location = new System.Drawing.Point(10, 8),
                Size = new System.Drawing.Size(140, 18)
            };
            _parameterFilterPanel.Controls.Add(title);

            // Add first two rows as defaults
            AddParameterRow("Reference Level", "A_GARDEN LEVEL");
            AddParameterRow("Size", "100x100");

            // Add button (plus)
            _addParameterButton = new WinForms.Button
            {
                Text = "+",
                Location = new System.Drawing.Point(_parameterFilterPanel.Width - 35, 6),
                Size = new System.Drawing.Size(24, 24),
                BackColor = System.Drawing.Color.FromArgb(230, 255, 230),
                FlatStyle = WinForms.FlatStyle.Flat,
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Right
            };
            _addParameterButton.Click += (_, __) => AddParameterRow("<Select>", "");
            _parameterFilterPanel.Controls.Add(_addParameterButton);
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
            nameCombo.Items.AddRange(new[] { "Reference Level", "Size", "Service Type", "Category", "<Select>" });
            nameCombo.SelectedItem = parameterName;
            row.Controls.Add(nameCombo);

            var valueCombo = new WinForms.ComboBox
            {
                Location = new System.Drawing.Point(160, 3),
                Size = new System.Drawing.Size(row.Width - 160 - 60, 21),
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right,
                DropDownStyle = WinForms.ComboBoxStyle.DropDownList
            };
            valueCombo.Items.AddRange(new[] { value, "<Auto Selection>", "100x100", "200x200", "A_GARDEN LEVEL", "ESSENTIAL POWER" });
            valueCombo.SelectedIndex = 0;
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
            _statusLabel.Text = "OK clicked";
        }

        private void OnSaveClick(object? sender, EventArgs e)
        {
            try
            {
                DebugLogger.Info("=== STARTING Save Configuration ===");
                _statusLabel.Text = "Saving configuration...";
                
                // Save current profile with UI state
                SaveCurrentConfiguration();
                
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
                    DebugLogger.Warning("No current profile found - creating default profile");
                    currentProfile = _appProfileService.CreateDefaultProfile();
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
                            
                            // Now check only the saved selections
                            for (int i = 0; i < listBox.Items.Count; i++)
                            {
                                var item = listBox.Items[i]?.ToString();
                                if (item != null && selectedFiles.Contains(item))
                                {
                                    listBox.SetItemChecked(i, true);
                                    DebugLogger.Info($"Restored reference file selection: {item} (index {i})");
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
                
                var lines = System.IO.File.ReadAllLines(configFile);
                var config = new UserConfiguration();
                
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
                
                profile.Configuration = config;
                
                DebugLogger.Info($"Loaded configuration from text file: {config.SelectedReferenceFiles?.Count ?? 0} reference files, {config.SelectedMepCategories?.Count ?? 0} MEP categories");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to load configuration from text file: {ex.Message}");
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
                    foreach (var control in _clearancePanel.Controls)
                    {
                        if (control is WinForms.TextBox textBox && textBox.Tag != null)
                        {
                            if (double.TryParse(textBox.Text, out double value))
                            {
                                clearances[textBox.Tag.ToString() ?? ""] = value;
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

        private void OnCancelClick(object? sender, EventArgs e)
        {
            _statusLabel.Text = "Cancel clicked";
        }

        private void OnCloseClick(object? sender, EventArgs e)
        {
            this.Close();
        }

        private void PositionPanels()
        {
            DebugLogger.Info("=== STARTING PositionPanels ===");
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

            _rightPanel.Location = new System.Drawing.Point(this.ClientSize.Width - _rightPanel.Width, top);
            _rightPanel.Size = new System.Drawing.Size(_rightPanel.Width, height);
            DebugLogger.Info($"_rightPanel positioned: Location={_rightPanel.Location}, Size={_rightPanel.Size}");
            DebugLogger.Info($"_rightPanel visibility: Visible={_rightPanel.Visible}, Parent={_rightPanel.Parent != null}");
            DebugLogger.Info($"Form ClientSize: {this.ClientSize}, Right edge calculation: {this.ClientSize.Width - _rightPanel.Width}");

            DebugLogger.Info($"Before splitter positioning: _rightPanel.Left={_rightPanel.Left}, _rightPanel.Location={_rightPanel.Location}");
            int splitterX = _rightPanel.Left - _mainSplitter.Width;
            DebugLogger.Info($"Calculated splitter X position: {splitterX} = {_rightPanel.Left} - {_mainSplitter.Width}");

            // FORCE the splitter position - don't let WinForms override it
            _mainSplitter.Location = new System.Drawing.Point(splitterX, top);
            _mainSplitter.Height = height;
            _mainSplitter.Anchor = WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Top; // Lock position
            DebugLogger.Info($"_mainSplitter positioned: Location={_mainSplitter.Location}, Height={_mainSplitter.Height}, Anchor={_mainSplitter.Anchor}");

            _leftPanel.Location = new System.Drawing.Point(0, top);
            _leftPanel.Size = new System.Drawing.Size(_mainSplitter.Left, height);
            DebugLogger.Info($"_leftPanel positioned: Location={_leftPanel.Location}, Size={_leftPanel.Size}");

            DebugLogger.Info("=== PositionPanels COMPLETED ===");
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
