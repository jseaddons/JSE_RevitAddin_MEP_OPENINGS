using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using WinForms = System.Windows.Forms;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for managing Host Element parameter transfer functionality
    /// Handles Wall, Structural Framing, and Floor parameter extraction and UI management
    /// </summary>
    public class HostParameterService
    {
        private readonly LinkedFileService _linkedFileService;
        private readonly ParameterExtractionService _parameterExtractionService;
        private Document _document; // Store document reference for row creation
        private List<string> _cachedHostParameters = new List<string>();
        private List<string> _cachedOpeningParameters = new List<string>();

        public HostParameterService()
        {
            _linkedFileService = new LinkedFileService();
            _parameterExtractionService = new ParameterExtractionService();
        }

        /// <summary>
        /// Creates host element tabs in the provided TabControl
        /// </summary>
        public void CreateHostTabs(WinForms.TabControl hostParameterTabs)
        {
            try
            {
                DebugLogger.Info("[HOST_SERVICE] Creating host element tabs");
                
                CreateHostTab(hostParameterTabs, "Walls", "WALLS");
                CreateHostTab(hostParameterTabs, "Structural Framing", "STRUCTURAL_FRAMING");
                CreateHostTab(hostParameterTabs, "Floors", "FLOORS");
                
                DebugLogger.Info("[HOST_SERVICE] Host element tabs created successfully");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[HOST_SERVICE] Error creating host tabs: {ex.Message}");
            }
        }

        /// <summary>
        /// Creates a single host tab with empty parameter rows (no default parameters)
        /// </summary>
        private void CreateHostTab(WinForms.TabControl hostParameterTabs, string tabName, string hostCode)
        {
            try
            {
                var tabPage = new WinForms.TabPage(tabName);
                
                var hostPanel = new WinForms.Panel
                {
                    Dock = WinForms.DockStyle.Fill,
                    BackColor = System.Drawing.Color.White,
                    Padding = new WinForms.Padding(3)
                };

                // Add default parameter rows (mimic Reference defaults for Host):
                // 1) Fire Rating → opening Fire Rating (if exists)
                AddHostParameterRow(hostPanel, hostCode, _document, "Fire Rating", "Fire Rating");
                // 2) Nearby Grids → opening Grid Names (try common names)
                AddHostParameterRow(hostPanel, hostCode, _document, "Grid", "Grid Names");

                // Add button (plus) for adding new parameter rows - positioned at right like Reference Elements
                var addButton = new WinForms.Button
                {
                    Text = "+",
                    Location = new System.Drawing.Point(hostPanel.Width - 35, 6),
                    Size = new System.Drawing.Size(24, 24),
                    Tag = hostCode,
                    BackColor = System.Drawing.Color.FromArgb(230, 255, 230), // Same green color as Reference Elements
                    FlatStyle = WinForms.FlatStyle.Flat,
                    Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Right // Same anchor as Reference Elements
                };
                addButton.Click += (_, __) =>
                {
                    // Create a new row - parameters will be populated by PopulateHostParameters
                    // Use stored document reference for opening parameter population
                    AddHostParameterRow(hostPanel, hostCode, _document);
                };
                hostPanel.Controls.Add(addButton);

                tabPage.Controls.Add(hostPanel);
                hostParameterTabs.TabPages.Add(tabPage);
                
                DebugLogger.Info($"[HOST_SERVICE] Created host tab: {tabName}");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[HOST_SERVICE] Error creating host tab '{tabName}': {ex.Message}");
            }
        }

        /// <summary>
        /// Adds a new parameter row to the host panel with populated dropdowns
        /// Same layout as Reference Elements
        /// </summary>
        private void AddHostParameterRow(WinForms.Panel hostPanel, string hostCode, Document document, string preselectHostParam = null, string preselectOpeningParam = null)
        {
            try
            {
                // =====  DIAGNOSTIC – DO NOT DELETE  =====
                if (document != null)
                {
                    var liveOpeningParams = _parameterExtractionService.GetEssentialOpeningParameters(document);
                    DebugLogger.Info($"[LIVE-DIAG] {nameof(AddHostParameterRow)} about to fill Opening combo with {liveOpeningParams.Count} items");
                }
                // =======================================

                int rowHeight = 24;
                
                // Count only parameter rows (panels with Tag = "parameterRow")
                var existingRows = hostPanel.Controls.OfType<WinForms.Panel>().Where(p => p.Tag?.ToString() == "parameterRow").ToList();
                int top = 25 + (existingRows.Count * (rowHeight + 3));

                var row = new WinForms.Panel
                {
                    Location = new System.Drawing.Point(8, top),
                    Size = new System.Drawing.Size(hostPanel.Width - 16, rowHeight),
                    Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right,
                    Tag = "parameterRow" // Tag to identify parameter rows
                };
                hostPanel.Controls.Add(row);

                // Host Parameter ComboBox (left side) - Empty for user selection
                var hostCombo = new WinForms.ComboBox
                {
                    Location = new System.Drawing.Point(0, 2),
                    Size = new System.Drawing.Size(120, 20),
                    DropDownStyle = WinForms.ComboBoxStyle.DropDownList,
                    Tag = "host"
                };
                hostCombo.Items.Add("<Select Host Parameter>");
                // Populate left combo from cached host parameters if available (ensures + rows get data)
                if (_cachedHostParameters != null && _cachedHostParameters.Count > 0)
                {
                    hostCombo.Items.AddRange(_cachedHostParameters.Cast<object>().ToArray());
                }
                hostCombo.SelectedIndex = 0;
                row.Controls.Add(hostCombo);

                // Opening Parameter ComboBox (right side) - Populate with cached opening parameters
                var openingCombo = new WinForms.ComboBox
                {
                    Location = new System.Drawing.Point(130, 2),
                    Size = new System.Drawing.Size(row.Width - 130 - 30, 24),
                    Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right,
                    DropDownStyle = WinForms.ComboBoxStyle.DropDownList,
                    Tag = "opening"
                };

                // Populate with live opening parameters using bootstrap routine - always fresh (same as Reference)
                if (document != null)
                {
                    var openingList = (_cachedOpeningParameters != null && _cachedOpeningParameters.Count > 0)
                        ? _cachedOpeningParameters
                        : _parameterExtractionService.GetEssentialOpeningParameters(document);
                    openingCombo.Items.AddRange(openingList.Cast<object>().ToArray());

                    // =====  DIAGNOSTIC – DO NOT DELETE  =====
                    LoggingConfiguration.ConditionalAppendAllText(
                        SafeFileLogger.GetLogFilePath("logger_debug.txt"),
                        $"[HOST-CHECK] Opening combo now contains {openingCombo.Items.Count} items (after AddRange){Environment.NewLine}");
                    // =======================================

                    // Force dropdown to show all items with scrolling (match Reference tab behavior)
                    openingCombo.IntegralHeight = false; // Allow partial items for better scrolling
                    openingCombo.MaxDropDownItems = 100; // Align with Reference tab
                    openingCombo.DropDownHeight = 400;
                }
                openingCombo.Items.Insert(0, "<Select Opening Parameter>");
                openingCombo.SelectedIndex = 0;
                row.Controls.Add(openingCombo);

                // Remove button - Same styling as Reference Elements
                var removeBtn = new WinForms.Button
                {
                    Text = "×",
                    Location = new System.Drawing.Point(row.Width - 25, 1),
                    Size = new System.Drawing.Size(20, 20),
                    BackColor = System.Drawing.Color.FromArgb(255, 230, 230),
                    FlatStyle = WinForms.FlatStyle.Flat,
                    Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Right // Same anchor as Reference Elements
                };
                removeBtn.Click += (s, e) => {
                    hostPanel.Controls.Remove(row);
                    row.Dispose();
                    // Re-layout remaining rows
                    RelayoutParameterRows(hostPanel);
                };
                row.Controls.Add(removeBtn);

                // Apply pre-selections if provided
                if (!string.IsNullOrWhiteSpace(preselectHostParam))
                {
                    var idx = hostCombo.FindStringExact(preselectHostParam);
                    if (idx >= 0) hostCombo.SelectedIndex = idx; // only if exists
                }
                if (!string.IsNullOrWhiteSpace(preselectOpeningParam))
                {
                    var idx2 = openingCombo.FindStringExact(preselectOpeningParam);
                    if (idx2 >= 0) openingCombo.SelectedIndex = idx2; // only if exists
                }

                DebugLogger.Info($"[HOST_SERVICE] Added parameter row for {hostCode}");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[HOST_SERVICE] Error adding parameter row for {hostCode}: {ex.Message}");
            }
        }

        /// <summary>
        /// Re-layouts parameter rows after one is removed
        /// </summary>
        private void RelayoutParameterRows(WinForms.Panel hostPanel)
        {
            int rowHeight = 24;
            int top = 25;

            var parameterRows = hostPanel.Controls.OfType<WinForms.Panel>()
                .Where(p => p.Tag?.ToString() == "parameterRow")
                .OrderBy(p => p.Location.Y)
                .ToList();

            foreach (var row in parameterRows)
            {
                row.Location = new System.Drawing.Point(8, top);
                top += rowHeight + 3;
            }
        }

        /// <summary>
        /// Populates host parameter dropdowns from selected host linked files
        /// </summary>
        public void PopulateHostParameters(WinForms.TabControl hostParameterTabs, List<string> selectedHostFiles, Document document)
        {
            try
            {
                // Store document reference for row creation
                _document = document;

                DebugLogger.Info($"[HOST_SERVICE] Populating host parameters for {selectedHostFiles.Count} selected host files");

                if (selectedHostFiles == null || selectedHostFiles.Count == 0)
                {
                    DebugLogger.Info("[HOST_SERVICE] No host files selected - skipping parameter population");
                    return;
                }

                // Get host parameters from selected linked files
                var hostParameters = GetHostParametersFromLinkedFiles(selectedHostFiles, document);
                DebugLogger.Info($"[HOST_SERVICE] Found {hostParameters.Count} host parameters from linked files");

                // Get opening parameters (use the optimized method to avoid destructive operations)
                var openingParameters = _parameterExtractionService.GetEssentialOpeningParameters(document);
                DebugLogger.Info($"[HOST_SERVICE] Found {openingParameters.Count} opening parameters");

                // Cache for add-row usage
                _cachedHostParameters = hostParameters ?? new List<string>();
                _cachedOpeningParameters = openingParameters ?? new List<string>();

                // Update each host tab with parameters
                foreach (WinForms.TabPage tabPage in hostParameterTabs.TabPages)
                {
                    UpdateHostTabParameters(tabPage, hostParameters, openingParameters);
                }

                DebugLogger.Info("[HOST_SERVICE] Host parameter population completed");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[HOST_SERVICE] Error populating host parameters: {ex.Message}");
            }
        }

        /// <summary>
        /// Gets host parameters from selected linked files
        /// </summary>
        private List<string> GetHostParametersFromLinkedFiles(List<string> selectedHostFiles, Document document)
        {
            var hostParameters = new List<string>();
            
            try
            {
                DebugLogger.Info($"[HOST_SERVICE] Getting host parameters from {selectedHostFiles.Count} selected host files");
                
                // Get all linked files
                var linkedFiles = _linkedFileService.GetLinkedFiles(document);
                DebugLogger.Info($"[HOST_SERVICE] Found {linkedFiles.Count} total linked files");

                // Also check the current document for host elements
                var currentDocParams = GetHostParametersFromDocument(document, "Current Document");
                AddUniqueParameters(hostParameters, currentDocParams);
                DebugLogger.Info($"[HOST_SERVICE] Found {currentDocParams.Count} host parameters from current document");

                foreach (var linkedFile in linkedFiles)
                {
                    // Check if this linked file is selected
                    var fileName = linkedFile.FileName;
                    var isSelected = selectedHostFiles.Any(selected => 
                        fileName.Contains(selected, StringComparison.OrdinalIgnoreCase) ||
                        selected.Contains(fileName, StringComparison.OrdinalIgnoreCase));

                    if (!isSelected)
                    {
                        DebugLogger.Info($"[HOST_SERVICE] Skipping unselected linked file: {fileName}");
                        continue;
                    }

                    DebugLogger.Info($"[HOST_SERVICE] Processing selected host file: {fileName}");

                    var linkedDoc = linkedFile.LinkInstance?.GetLinkDocument();
                    if (linkedDoc != null)
                    {
                        var linkedDocParams = GetHostParametersFromDocument(linkedDoc, fileName);
                        AddUniqueParameters(hostParameters, linkedDocParams);
                        DebugLogger.Info($"[HOST_SERVICE] Retrieved {linkedDocParams.Count} host parameters from {fileName}");
                    }
                }

                DebugLogger.Info($"[HOST_SERVICE] Total unique host parameters: {hostParameters.Count}");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[HOST_SERVICE] Error getting host parameters from linked files: {ex.Message}");
            }

            return hostParameters;
        }

        /// <summary>
        /// Gets host parameters from a specific document (current or linked)
        /// </summary>
        private List<string> GetHostParametersFromDocument(Document document, string documentName)
        {
            var parameters = new List<string>();
            
            try
            {
                DebugLogger.Info($"[HOST_SERVICE] Getting host parameters from document: {documentName}");
                
                // ✅ PERFORMANCE FIX: Only load Fire Rating parameter on startup (rest load when user clicks "+")
                // Check if Fire Rating parameter exists in any of the host categories
                var categories = new[] { BuiltInCategory.OST_Walls, BuiltInCategory.OST_StructuralFraming, BuiltInCategory.OST_Floors };
                
                foreach (var category in categories)
                {
                    var collector = new FilteredElementCollector(document)
                        .OfCategory(category)
                        .WhereElementIsNotElementType()
                        .Take(1);
                    
                    foreach (Element element in collector)
                    {
                        foreach (Parameter param in element.Parameters)
                        {
                            if (string.Equals(param.Definition?.Name, "Fire Rating", StringComparison.OrdinalIgnoreCase))
                            {
                                if (!parameters.Contains("Fire Rating"))
                                {
                                    parameters.Add("Fire Rating");
                                    break;
                                }
                            }
                        }
                        if (parameters.Contains("Fire Rating")) break;
                    }
                    if (parameters.Contains("Fire Rating")) break;
                }

                DebugLogger.Info($"[HOST_SERVICE] Retrieved {parameters.Count} host parameters from {documentName} (startup: only Fire Rating)");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[HOST_SERVICE] Error getting host parameters from document {documentName}: {ex.Message}");
            }

            return parameters;
        }

        /// <summary>
        /// ✅ PERFORMANCE FIX: Gets only essential/shared parameters for a specific category
        /// Only loads: Fire Rating, Grid, and shared parameters (from first element sample)
        /// This limits loading to ~1-5 parameters per category instead of scanning all parameters
        /// </summary>
        private List<string> GetParametersForCategory(Document document, BuiltInCategory category)
        {
            var parameters = new HashSet<string>();
            
            try
            {
                // ✅ PERFORMANCE: Only sample 1 element instead of 10 (user wants ~1 parameter per host type)
                var collector = new FilteredElementCollector(document)
                    .OfCategory(category)
                    .WhereElementIsNotElementType()
                    .Take(1); // ✅ PERFORMANCE: Only 1 element for essential parameters

                // ✅ PERFORMANCE: Pre-define essential parameters (user requirement: ~1 per host type)
                var essentialParams = new[] { "Fire Rating", "Grid", "Grid Names" };
                foreach (var essentialParam in essentialParams)
                {
                    parameters.Add(essentialParam);
                }

                // Get shared parameters from first element only (for the ~20 shared params user mentioned)
                foreach (Element element in collector)
                {
                    foreach (Parameter param in element.Parameters)
                    {
                        if (string.IsNullOrEmpty(param.Definition?.Name)) continue;
                        
                        var paramName = param.Definition.Name;
                        
                        // ✅ PERFORMANCE: Only add shared parameters (they have positive IDs and are part of the ~20)
                        // Skip instance/type parameters that aren't shared (user only wants shared params from families)
                        bool isSharedParameter = param.Id.IntegerValue > 0 && param.IsShared;
                        
                        if (isSharedParameter && 
                            !paramName.StartsWith("Internal") &&
                            !paramName.StartsWith("Revit") &&
                            !paramName.StartsWith("Assembly"))
                        {
                            parameters.Add(paramName);
                        }
                    }
                }

                DebugLogger.Info($"[HOST_SERVICE] Found {parameters.Count} essential/shared parameters for category {category}");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[HOST_SERVICE] Error getting parameters for category {category}: {ex.Message}");
            }

            return parameters.ToList();
        }

        /// <summary>
        /// Gets opening parameters from the current document (same logic as ParameterExtractionService)
        /// </summary>
        private List<string> GetOpeningParameters(Document document)
        {
            var openingParameters = new List<string>();

            try
            {
                DebugLogger.Info($"[HOST_SERVICE] Getting opening parameters from document");

                // Method 1: Get parameters from FamilySymbol (Type Parameters)
                var openingFamilySymbols = GetOpeningFamilies(document);
                foreach (var familySymbol in openingFamilySymbols)
                {
                    foreach (Parameter param in familySymbol.Parameters)
                    {
                        if (!string.IsNullOrEmpty(param.Definition?.Name) && !openingParameters.Contains(param.Definition.Name))
                        {
                            openingParameters.Add(param.Definition.Name);
                        }
                    }
                }
                DebugLogger.Info($"[HOST_SERVICE] Found {openingParameters.Count} parameters from FamilySymbols (type parameters)");

                // Method 2: Get parameters from FamilySymbol (Type Parameters) - works even without instances
                var targetFamilyNames = new List<string>
                {
                    "RectangularOpeningOnWall",
                    "RectangularOpeningOnSlab",
                    "CircularOpeningOnWall",
                    "CircularOpeningOnSlab"
                };

                var symbolCollector = new FilteredElementCollector(document)
                    .OfClass(typeof(FamilySymbol))
                    .WhereElementIsElementType();

                int symbolCount = 0;
                foreach (Element element in symbolCollector)
                {
                    if (element is FamilySymbol familySymbol)
                    {
                        var familyName = familySymbol.Family?.Name ?? "";
                        var symbolName = familySymbol.Name ?? "";

                        // Check if this is one of our opening families
                        bool isTargetFamily = targetFamilyNames.Any(targetName =>
                            familyName.Contains(targetName) ||
                            symbolName.Contains(targetName) ||
                            $"{familyName} {symbolName}".Contains(targetName));

                        if (isTargetFamily)
                        {
                            symbolCount++;

                            // Get TYPE parameters from FamilySymbol
                            foreach (Parameter param in familySymbol.Parameters)
                            {
                                if (!string.IsNullOrEmpty(param.Definition?.Name) && !openingParameters.Contains(param.Definition.Name))
                                {
                                    openingParameters.Add(param.Definition.Name);
                                    DebugLogger.Info($"[HOST_SERVICE] Added type parameter from FamilySymbol: '{param.Definition.Name}'");
                                }
                            }

                            // Only need to check a few symbols
                            if (symbolCount >= 10) break;
                        }
                    }
                }

                DebugLogger.Info($"[HOST_SERVICE] Checked {symbolCount} opening family symbols for type parameters");
                DebugLogger.Info($"[HOST_SERVICE] Total unique opening parameters found: {openingParameters.Count}");

                // If no parameters found from families, use fallback from shared parameter file
                if (openingParameters.Count == 0)
                {
                    DebugLogger.Info("[HOST_SERVICE] No parameters found from families - using fallback from shared parameter file");

                    var fallbackParameters = GetOpeningParametersFromSharedFile();
                    openingParameters = fallbackParameters;

                    DebugLogger.Info($"[HOST_SERVICE] Using {openingParameters.Count} parameters from shared parameter file as fallback");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[HOST_SERVICE] Error getting opening parameters: {ex.Message}");
            }

            return openingParameters;
        }

        /// <summary>
        /// Gets opening families from the current document (same as ParameterExtractionService)
        /// </summary>
        private List<FamilySymbol> GetOpeningFamilies(Document document)
        {
            var openingFamilies = new List<FamilySymbol>();

            try
            {
                // Specific opening family names to filter by (only the 4 current families)
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

                DebugLogger.Info($"[HOST_SERVICE] GetOpeningFamilies inspected {inspected} FamilySymbols, matched {openingFamilies.Count} families from the 4 specific opening families: {string.Join(", ", targetFamilyNames)}");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[HOST_SERVICE] Error getting opening families: {ex.Message}");
            }

            return openingFamilies;
        }

        /// <summary>
        /// Get opening parameters directly from the shared parameter file as fallback
        /// </summary>
        private List<string> GetOpeningParametersFromSharedFile()
        {
            var parameters = new List<string>();

            try
            {
                // Path to shared parameter file
                string sharedParamFile = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Resources\Opening family shared parameter.txt";

                if (!System.IO.File.Exists(sharedParamFile))
                {
                    DebugLogger.Warning($"[HOST_SERVICE] Shared parameter file not found: {sharedParamFile}");
                    return parameters;
                }

                // Read all lines from the shared parameter file
                var lines = File.ReadAllLines(sharedParamFile);

                foreach (var line in lines)
                {
                    // Look for parameter definition lines (start with "PARAM")
                    if (line.StartsWith("PARAM"))
                    {
                        var parts = line.Split('\t');
                        if (parts.Length >= 3)
                        {
                            // The parameter name is the 3rd field (index 2)
                            string paramName = parts[2];
                            if (!string.IsNullOrEmpty(paramName) && !parameters.Contains(paramName))
                            {
                                parameters.Add(paramName);
                                DebugLogger.Info($"[HOST_SERVICE] Added fallback parameter from shared file: '{paramName}'");
                            }
                        }
                    }
                }

                DebugLogger.Info($"[HOST_SERVICE] Loaded {parameters.Count} parameters from shared parameter file");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[HOST_SERVICE] Error reading shared parameter file: {ex.Message}");
            }

            return parameters;
        }

        /// <summary>
        /// Updates a single host tab with parameter lists
        /// </summary>
        private void UpdateHostTabParameters(WinForms.TabPage tabPage, List<string> hostParameters, List<string> openingParameters)
        {
            try
            {
                DebugLogger.Info($"[HOST_SERVICE] Updating tab '{tabPage.Text}' with {hostParameters.Count} host and {openingParameters.Count} opening parameters");

                // Find the hostPanel in this tab
                var hostPanel = tabPage.Controls.OfType<WinForms.Panel>().FirstOrDefault();
                
                if (hostPanel != null)
                {
                    // Find all ComboBoxes in existing rows
                    var existingRows = hostPanel.Controls.OfType<WinForms.Panel>().Where(p => p.Tag?.ToString() == "parameterRow").ToList();
                    
                    foreach (var row in existingRows)
                    {
                        var hostCombo = row.Controls.OfType<WinForms.ComboBox>().FirstOrDefault(cb => cb.Tag?.ToString() == "host");
                        var openingCombo = row.Controls.OfType<WinForms.ComboBox>().FirstOrDefault(cb => cb.Tag?.ToString() == "opening");

                        if (hostCombo != null)
                        {
                            hostCombo.Items.Clear();
                            hostCombo.Items.Add("<Select Host Parameter>");
                            hostCombo.Items.AddRange(hostParameters.ToArray());
                            hostCombo.SelectedIndex = 0;
                        }

                        if (openingCombo != null)
                        {
                            openingCombo.Items.Clear();
                            openingCombo.Items.Add("<Select Opening Parameter>");
                            openingCombo.Items.AddRange(openingParameters.ToArray());
                            openingCombo.SelectedIndex = 0;
                        }
                    }

                    DebugLogger.Info($"[HOST_SERVICE] Updated {existingRows.Count} existing rows in tab '{tabPage.Text}'");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[HOST_SERVICE] Error updating host tab parameters: {ex.Message}");
            }
        }

        /// <summary>
        /// Adds unique parameters to the list
        /// </summary>
        private void AddUniqueParameters(List<string> parameterList, List<string> newParameters)
        {
            foreach (var param in newParameters)
            {
                if (!parameterList.Contains(param))
                {
                    parameterList.Add(param);
                }
            }
        }
    }
}
