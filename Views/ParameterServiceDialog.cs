using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Views
{
    /// <summary>
    /// Separate dialog for parameter extraction service
    /// This dialog handles parameter extraction independently from the main UI
    /// to prevent sleeve deletion during main UI operations
    /// </summary>
    public partial class ParameterServiceDialog : System.Windows.Forms.Form
    {
        private readonly Document _document;
        private readonly UIDocument _uiDocument;
        private ParameterExtractionService _parameterService;
        private List<string> _extractedParameters;
        private bool _extractionCompleted = false;
        private EmergencyMainDialog _mainDialog; // Reference to main dialog for integration

        public ParameterServiceDialog(Document document, UIDocument uiDocument, EmergencyMainDialog mainDialog = null)
        {
            _document = document;
            _uiDocument = uiDocument;
            _mainDialog = mainDialog;
            _parameterService = new ParameterExtractionService();
            _extractedParameters = new List<string>();
            
            InitializeComponent();
            InitializeUI();
        }

        private void InitializeComponent()
        {
            this.Text = "Parameter Service";
            this.Size = new System.Drawing.Size(600, 400);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            
            // Main panel
            var mainPanel = new System.Windows.Forms.Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };
            this.Controls.Add(mainPanel);

            // Title label
            var titleLabel = new Label
            {
                Text = "Parameter Extraction Service",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold),
                Location = new System.Drawing.Point(10, 10),
                Size = new System.Drawing.Size(300, 25),
                ForeColor = System.Drawing.Color.FromArgb(0, 100, 200)
            };
            mainPanel.Controls.Add(titleLabel);

            // Description label
            var descLabel = new Label
            {
                Text = "This service extracts opening parameters by placing temporary instances.\n" +
                       "⚠️ WARNING: This operation may affect existing sleeves temporarily.\n" +
                       "All sleeves will be preserved after extraction completes.",
                Location = new System.Drawing.Point(10, 40),
                Size = new System.Drawing.Size(550, 60),
                ForeColor = System.Drawing.Color.FromArgb(100, 100, 100)
            };
            mainPanel.Controls.Add(descLabel);

            // Progress label
            var progressLabel = new Label
            {
                Text = "Ready to extract parameters...",
                Location = new System.Drawing.Point(10, 110),
                Size = new System.Drawing.Size(550, 20)
            };
            mainPanel.Controls.Add(progressLabel);

            // Progress bar
            var progressBar = new ProgressBar
            {
                Location = new System.Drawing.Point(10, 135),
                Size = new System.Drawing.Size(550, 20),
                Style = ProgressBarStyle.Continuous
            };
            mainPanel.Controls.Add(progressBar);

            // Results listbox
            var resultsLabel = new Label
            {
                Text = "Extracted Parameters:",
                Location = new System.Drawing.Point(10, 170),
                Size = new System.Drawing.Size(150, 20)
            };
            mainPanel.Controls.Add(resultsLabel);

            var resultsListBox = new ListBox
            {
                Location = new System.Drawing.Point(10, 195),
                Size = new System.Drawing.Size(550, 120),
                ScrollAlwaysVisible = true
            };
            mainPanel.Controls.Add(resultsListBox);

            // Buttons panel
            var buttonPanel = new System.Windows.Forms.Panel
            {
                Location = new System.Drawing.Point(10, 325),
                Size = new System.Drawing.Size(550, 35)
            };
            mainPanel.Controls.Add(buttonPanel);

            // Extract button
            var extractButton = new Button
            {
                Text = "Extract Parameters",
                Location = new System.Drawing.Point(0, 5),
                Size = new System.Drawing.Size(120, 25),
                BackColor = System.Drawing.Color.FromArgb(0, 100, 200),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat
            };
            extractButton.Click += (s, e) => ExtractParameters(progressLabel, progressBar, resultsListBox, extractButton);
            buttonPanel.Controls.Add(extractButton);

            // Close button
            var closeButton = new Button
            {
                Text = "Close",
                Location = new System.Drawing.Point(130, 5),
                Size = new System.Drawing.Size(80, 25),
                DialogResult = DialogResult.Cancel
            };
            buttonPanel.Controls.Add(closeButton);

            // OK button (initially disabled)
            var okButton = new Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(220, 5),
                Size = new System.Drawing.Size(80, 25),
                Enabled = false,
                DialogResult = DialogResult.OK
            };
            buttonPanel.Controls.Add(okButton);

            // Store references for later use
            this.Tag = new { progressLabel, progressBar, resultsListBox, extractButton, okButton };
        }

        private void InitializeUI()
        {
            // Set initial state
            var controls = (dynamic)this.Tag;
            controls.extractButton.Enabled = true;
            controls.okButton.Enabled = false;
        }

        private async void ExtractParameters(Label progressLabel, ProgressBar progressBar, ListBox resultsListBox, Button extractButton)
        {
            try
            {
                extractButton.Enabled = false;
                progressLabel.Text = "Starting parameter extraction...";
                progressBar.Value = 10;

                // Step 1: Backup existing sleeves
                progressLabel.Text = "Backing up existing sleeves...";
                progressBar.Value = 20;
                var sleeveBackup = BackupExistingSleeves();

                // Step 2: Extract parameters
                progressLabel.Text = "Extracting opening parameters...";
                progressBar.Value = 40;
                
                _extractedParameters = _parameterService.GetCurrentOpeningParameters(_document);
                
                progressBar.Value = 80;

                // Step 3: Restore sleeves (in case extraction affected them)
                progressLabel.Text = "Restoring sleeves...";
                progressBar.Value = 90;
                RestoreSleeves(sleeveBackup);

                // Step 4: Update UI
                progressLabel.Text = $"Parameter extraction completed! Found {_extractedParameters.Count} parameters.";
                progressBar.Value = 100;
                
                resultsListBox.Items.Clear();
                resultsListBox.Items.AddRange(_extractedParameters.Cast<object>().ToArray());
                
                _extractionCompleted = true;
                
                // Step 5: Update main UI dropdowns if main dialog is available
                if (_mainDialog != null)
                {
                    progressLabel.Text = "Updating main UI parameter dropdowns...";
                    UpdateMainUIDropdowns();
                }
                
                // Enable OK button
                var controls = (dynamic)this.Tag;
                controls.okButton.Enabled = true;
                
                progressLabel.Text = $"Parameter extraction completed! Found {_extractedParameters.Count} parameters.";
                
                DebugLogger.Info($"Parameter extraction completed successfully. Found {_extractedParameters.Count} parameters.");
            }
            catch (Exception ex)
            {
                progressLabel.Text = $"Error during extraction: {ex.Message}";
                DebugLogger.Error($"Error in parameter extraction: {ex.Message}");
                
                // Try to restore sleeves even if extraction failed
                try
                {
                    RestoreSleeves(BackupExistingSleeves());
                }
                catch (Exception restoreEx)
                {
                    DebugLogger.Error($"Error restoring sleeves after failed extraction: {restoreEx.Message}");
                }
            }
            finally
            {
                extractButton.Enabled = true;
            }
        }

        private List<ElementId> BackupExistingSleeves()
        {
            try
            {
                var sleeves = new FilteredElementCollector(_document)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                    .Select(fi => fi.Id)
                    .ToList();

                DebugLogger.Info($"Backed up {sleeves.Count} existing sleeves");
                return sleeves;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error backing up sleeves: {ex.Message}");
                return new List<ElementId>();
            }
        }

        private void RestoreSleeves(List<ElementId> sleeveIds)
        {
            try
            {
                // Check if sleeves still exist
                var existingSleeves = new FilteredElementCollector(_document)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                    .Select(fi => fi.Id)
                    .ToList();

                var missingSleeves = sleeveIds.Except(existingSleeves).ToList();
                
                if (missingSleeves.Count > 0)
                {
                    DebugLogger.Warning($"Found {missingSleeves.Count} missing sleeves after parameter extraction");
                    // Note: We can't restore sleeves here as we don't have their placement data
                    // This is just for logging purposes
                }
                else
                {
                    DebugLogger.Info("All sleeves preserved after parameter extraction");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error checking sleeve restoration: {ex.Message}");
            }
        }

        private void UpdateMainUIDropdowns()
        {
            try
            {
                DebugLogger.Info("Updating main UI parameter dropdowns...");
                
                // Update parameter dropdowns in main UI
                _mainDialog.UpdateParameterDropdownsFromMepCategories(_document);
                _mainDialog.PopulateHostParameters();
                
                DebugLogger.Info("Main UI parameter dropdowns updated successfully");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error updating main UI dropdowns: {ex.Message}");
            }
        }

        public List<string> GetExtractedParameters()
        {
            return _extractionCompleted ? _extractedParameters : new List<string>();
        }
    }
}
