using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using System.Windows.Forms;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public class ParameterInfo
    {
        public string Name { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public bool IsInstanceParameter { get; set; }
        public BuiltInParameter? BuiltInParameter { get; set; }
        public List<string> Values { get; set; } = new List<string>();
    }

    public class ParameterExtractionService
    {
        public List<ParameterInfo> GetParametersForCategory(Document document, BuiltInCategory category, bool includeInstanceParams = true, bool includeTypeParams = true)
        {
            var parameters = new List<ParameterInfo>();

            try
            {
                // Get all elements of the specified category
                var collector = new FilteredElementCollector(document)
                    .OfCategory(category)
                    .WhereElementIsNotElementType();

                var elements = collector.ToElements();
                if (elements.Count == 0) return parameters;

                // Sample a few elements to extract parameters
                var sampleElements = elements.Take(10).ToList();

                // Get parameters from instance elements
                if (includeInstanceParams)
                {
                    foreach (var element in sampleElements)
                    {
                        ExtractParametersFromElement(element, parameters, true);
                    }
                }

                // Get parameters from element types
                if (includeTypeParams)
                {
                    foreach (var element in sampleElements)
                    {
                        var elementType = document.GetElement(element.GetTypeId());
                        if (elementType != null)
                        {
                            ExtractParametersFromElement(elementType, parameters, false);
                        }
                    }
                }

                // Remove duplicates and sort
                parameters = parameters
                    .GroupBy(p => p.Name)
                    .Select(g => g.First())
                    .OrderBy(p => p.Name)
                    .ToList();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error extracting parameters for category {category}: {ex.Message}");
            }

            return parameters;
        }

        public List<ParameterInfo> GetParametersForMepCategories(Document document, List<MepCategory> categories)
        {
            var allParameters = new List<ParameterInfo>();

            foreach (var category in categories)
            {
                var builtinCategories = GetBuiltInCategoriesForMepCategory(category);
                foreach (var builtinCategory in builtinCategories)
                {
                    var categoryParams = GetParametersForCategory(document, builtinCategory);
                    allParameters.AddRange(categoryParams);
                }
            }

            // Remove duplicates and sort
            return allParameters
                .GroupBy(p => p.Name)
                .Select(g => g.First())
                .OrderBy(p => p.Name)
                .ToList();
        }

        private void ExtractParametersFromElement(Element element, List<ParameterInfo> parameters, bool isInstance)
        {
            try
            {
                foreach (Parameter param in element.Parameters)
                {
                    if (param == null || string.IsNullOrEmpty(param.Definition?.Name)) continue;

                    var paramInfo = new ParameterInfo
                    {
                        Name = param.Definition?.Name ?? "Unknown",
                        Type = param.StorageType.ToString(),
                        IsInstanceParameter = isInstance,
                        BuiltInParameter = param.Id.IntegerValue < 0 ?
                            (BuiltInParameter)param.Id.IntegerValue : null
                    };

                    // Try to get some sample values
                    if (param.HasValue)
                    {
                        try
                        {
                            string value = GetParameterValueAsString(param);
                            if (!string.IsNullOrEmpty(value) && !paramInfo.Values.Contains(value))
                            {
                                paramInfo.Values.Add(value);
                            }
                        }
                        catch
                        {
                            // Ignore value extraction errors
                        }
                    }

                    parameters.Add(paramInfo);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error extracting parameters from element: {ex.Message}");
            }
        }

        private string GetParameterValueAsString(Parameter param)
        {
            try
            {
                switch (param.StorageType)
                {
                    case StorageType.String:
                        return param.AsString() ?? string.Empty;
                    case StorageType.Integer:
                        return param.AsInteger().ToString();
                    case StorageType.Double:
                        return param.AsDouble().ToString();
                    case StorageType.ElementId:
                        var elemId = param.AsElementId();
                        return elemId.IntegerValue.ToString();
                    default:
                        return string.Empty;
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        private List<BuiltInCategory> GetBuiltInCategoriesForMepCategory(MepCategory category)
        {
            return category switch
            {
                MepCategory.Pipes => new List<BuiltInCategory> { BuiltInCategory.OST_PipeCurves, BuiltInCategory.OST_PipeFitting },
                MepCategory.Ducts => new List<BuiltInCategory> { BuiltInCategory.OST_DuctCurves },
                MepCategory.DuctAccessories => new List<BuiltInCategory> { BuiltInCategory.OST_DuctAccessory },
                MepCategory.DuctFittings => new List<BuiltInCategory> { BuiltInCategory.OST_DuctFitting },
                MepCategory.CableTrays => new List<BuiltInCategory> { BuiltInCategory.OST_CableTray },
                MepCategory.Conduits => new List<BuiltInCategory> { BuiltInCategory.OST_Conduit },
                _ => new List<BuiltInCategory>()
            };
        }

        public List<string> GetParameterNamesForDisplay(List<ParameterInfo> parameters)
        {
            return parameters.Select(p => p.Name).ToList();
        }

        public List<string> GetParameterValuesForParameter(List<ParameterInfo> parameters, string parameterName)
        {
            var param = parameters.FirstOrDefault(p => p.Name == parameterName);
            return param?.Values ?? new List<string>();
        }

        /// <summary>
        /// Populates parameters for tabs matching the selected MEP categories
        /// </summary>
        public void PopulateCategorySpecificParameters(TabControl serviceParameterTabs, List<string> selectedCategories, Document document)
        {
            try
            {
                // Log to main log file
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] [PARAMETER_SERVICE] Starting category-specific population for {selectedCategories.Count} categories\n");
                
                // Get opening parameters (FIXED: now includes both type and instance parameters)
                var openingParameters = GetCurrentOpeningParameters(document);
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] [PARAMETER_SERVICE] Found {openingParameters.Count} opening parameters\n");

                // Get linked files for parameter extraction (both MEP and architectural files)
                var linkedFileService = new Services.LinkedFileService();
                var allLinkedFiles = linkedFileService.GetLinkedFiles(document);

                // DEBUG: Check what files are returned by LinkedFileService
                System.Diagnostics.Debug.WriteLine($"[HOST] LinkedFileService returned {allLinkedFiles.Count} files");
                foreach (var f in allLinkedFiles)
                    System.Diagnostics.Debug.WriteLine($"[HOST]   {f.FileName}  type={f.FileType}");

                // Get only host element files (architectural and structural)
                var hostLinkedFiles = linkedFileService.GetHostElementFiles(allLinkedFiles);
                System.Diagnostics.Debug.WriteLine($"[HOST] Filtered to {hostLinkedFiles.Count} host element files");
                foreach (var f in hostLinkedFiles)
                    System.Diagnostics.Debug.WriteLine($"[HOST]   Host file: {f.FileName}  type={f.FileType}");

                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] [PARAMETER_SERVICE] Found {allLinkedFiles.Count} total linked files, {hostLinkedFiles.Count} host element files\n");

                // Get host element parameters from linked architectural files (walls, floors, ceilings are in linked files)
                var linkedHostParameters = GetHostElementParametersFromLinkedFiles(hostLinkedFiles);
                System.Diagnostics.Debug.WriteLine($"[HOST] GetHostElementParametersFromLinkedFiles returned {linkedHostParameters.Count} parameters");

                // DEBUG: Check if we got any host parameters
                if (linkedHostParameters.Count == 0)
                {
                    System.Diagnostics.Debug.WriteLine($"[HOST] WARNING: No host parameters found from linked files!");
                    System.Diagnostics.Debug.WriteLine($"[HOST] This means either:");
                    System.Diagnostics.Debug.WriteLine($"[HOST] 1. No architectural files are linked");
                    System.Diagnostics.Debug.WriteLine($"[HOST] 2. Architectural files are not detected as 'Architectural' type");
                    System.Diagnostics.Debug.WriteLine($"[HOST] 3. No walls/floors/ceilings found in linked files");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[HOST] Found {linkedHostParameters.Count} host parameters: {string.Join(", ", linkedHostParameters.Take(5))}");
                }

                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] [PARAMETER_SERVICE] Found {linkedHostParameters.Count} host element parameters from linked files\n");

                // Combine opening and host parameters from linked files
                var combinedOpeningParameters = openingParameters.Union(linkedHostParameters).Distinct().OrderBy(p => p).ToList();
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] [PARAMETER_SERVICE] Combined total: {combinedOpeningParameters.Count} unique parameters\n");
                
                // Update each tab that matches selected categories
                foreach (var selectedCategory in selectedCategories)
                {
                    var matchingTab = FindTabByCategory(serviceParameterTabs, selectedCategory);
                    if (matchingTab != null)
                    {
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] [PARAMETER_SERVICE] Updating tab '{matchingTab.Text}' for category '{selectedCategory}'\n");

                        var categorySpecificParameters = GetParametersForSpecificCategoryFromLinkedFiles(selectedCategory, allLinkedFiles);
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] [PARAMETER_SERVICE] Found {categorySpecificParameters.Count} parameters for category '{selectedCategory}' from linked files\n");

                        UpdateSingleTabParameters(matchingTab, categorySpecificParameters, combinedOpeningParameters);
                    }
                    else
                    {
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] [PARAMETER_SERVICE] No tab found for category '{selectedCategory}'\n");
                    }
                }

                // Handle special tabs that don't correspond to MEP categories
                var hostTab = FindTabByName(serviceParameterTabs, "Host to Opening");
                if (hostTab != null)
                {
                    System.Diagnostics.Debug.WriteLine($"[HOST] Tab '{hostTab.Text}' gets {linkedHostParameters.Count} host params + {combinedOpeningParameters.Count} opening params");

                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] [PARAMETER_SERVICE] Updating 'Host to Opening' tab with host parameters\n");

                    // For Host tab: Source = host parameters, Target = opening parameters
                    UpdateSingleTabParameters(hostTab, linkedHostParameters, combinedOpeningParameters);
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[HOST] WARNING: 'Host to Opening' tab not found!");
                }
                
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] [PARAMETER_SERVICE] Category-specific parameter population completed successfully\n");
            }
            catch (Exception ex)
            {
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] [PARAMETER_SERVICE] Error in PopulateCategorySpecificParameters: {ex.Message}\n");
            }
        }

        private TabPage FindTabByCategory(TabControl serviceParameterTabs, string categoryName)
        {
            if (serviceParameterTabs?.TabPages.Count > 0)
            {
                foreach (TabPage tabPage in serviceParameterTabs.TabPages)
                {
                    if (string.Equals(tabPage.Text, categoryName, StringComparison.OrdinalIgnoreCase))
                    {
                        return tabPage;
                    }
                }
            }
            return null;
        }

        private TabPage FindTabByName(TabControl tabControl, string tabName)
        {
            if (tabControl?.TabPages.Count > 0)
            {
                foreach (TabPage tabPage in tabControl.TabPages)
                {
                    // Case-insensitive and space-tolerant comparison
                    if (tabPage.Text.Replace(" ", "").Equals(tabName.Replace(" ", ""), StringComparison.OrdinalIgnoreCase))
                    {
                        return tabPage;
                    }
                }
            }
            return null;
        }

        private List<string> GetParametersForSpecificCategoryFromLinkedFiles(string categoryName, List<Services.LinkedFileInfo> linkedFiles)
        {
            try
            {
                var categoryParameters = new List<string>();
                
                if (linkedFiles == null || linkedFiles.Count == 0)
                {
                    System.Diagnostics.Debug.WriteLine($"[PARAMETER_CATEGORY] No linked files for category '{categoryName}'");
                    return categoryParameters;
                }
                
                // Handle category name mapping for enum parsing
                var enumCategoryName = categoryName.Replace(" ", ""); // Remove spaces
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_CATEGORY] Original category: '{categoryName}', Enum category: '{enumCategoryName}'");
                if (Enum.TryParse<MepCategory>(enumCategoryName, out var category))
                {
                    // Get parameters from all linked files
                    foreach (var linkedFile in linkedFiles)
                    {
                        try
                        {
                            var linkedDoc = linkedFile.LinkInstance?.GetLinkDocument();
                            if (linkedDoc != null)
                            {
                                var parameters = GetParametersForMepCategories(linkedDoc, new List<MepCategory> { category });
                                var parameterNames = parameters.Select(p => p.Name).ToList();
                                
                                // Add unique parameters
                                foreach (var paramName in parameterNames)
                                {
                                    if (!categoryParameters.Contains(paramName))
                                    {
                                        categoryParameters.Add(paramName);
                                    }
                                }
                                
                                System.Diagnostics.Debug.WriteLine($"[PARAMETER_CATEGORY] Retrieved {parameterNames.Count} parameters from linked file '{linkedFile.FileName}' for category '{categoryName}'");
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[PARAMETER_CATEGORY] Error getting parameters from linked file '{linkedFile.FileName}': {ex.Message}");
                        }
                    }
                    
                    System.Diagnostics.Debug.WriteLine($"[PARAMETER_CATEGORY] Total unique parameters for category '{categoryName}': {categoryParameters.Count}");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[PARAMETER_CATEGORY] Could not parse category name '{categoryName}' to enum");
                }
                
                return categoryParameters;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_CATEGORY] Error getting parameters for category '{categoryName}' from linked files: {ex.Message}");
                return new List<string>();
            }
        }

        /// <summary>
        /// FIXED: Get opening parameters from both TYPE and INSTANCE parameters of placed opening families
        /// </summary>
        private List<string> GetCurrentOpeningParameters(Document document)
        {
            var openingParameters = new HashSet<string>(); // Use HashSet to avoid duplicates
            
            try
            {
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_OPENING] Starting to collect opening parameters...");
                
                // Method 1: Get parameters from FamilySymbol (Type Parameters)
                var openingFamilySymbols = GetOpeningFamilies(document);
                foreach (var familySymbol in openingFamilySymbols)
                {
                    foreach (Parameter param in familySymbol.Parameters)
                    {
                        if (!string.IsNullOrEmpty(param.Definition?.Name))
                        {
                            openingParameters.Add(param.Definition.Name);
                        }
                    }
                }
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_OPENING] Found {openingParameters.Count} parameters from FamilySymbols (type parameters)");
                
                // Method 2: Get parameters from PLACED INSTANCES (Instance Parameters - THIS WAS MISSING!)
                var targetFamilyNames = new List<string>
                {
                    "RectangularOpeningOnWall",
                    "RectangularOpeningOnSlab",
                    "CircularOpeningOnWall",
                    "CircularOpeningOnSlab"
                };

                var instanceCollector = new FilteredElementCollector(document)
                    .OfClass(typeof(FamilyInstance))
                    .WhereElementIsNotElementType();

                int instanceCount = 0;
                foreach (Element element in instanceCollector)
                {
                    if (element is FamilyInstance famInst)
                    {
                        var familyName = famInst.Symbol?.Family?.Name ?? "";
                        
                        // Check if this is one of our opening families
                        bool isTargetFamily = targetFamilyNames.Any(targetName =>
                            familyName.Contains(targetName));

                        if (isTargetFamily)
                        {
                            instanceCount++;
                            
                            // Get INSTANCE parameters
                            foreach (Parameter param in famInst.Parameters)
                            {
                                if (!string.IsNullOrEmpty(param.Definition?.Name))
                                {
                                    openingParameters.Add(param.Definition.Name);
                                }
                            }
                            
                            // Only need to check a few instances
                            if (instanceCount >= 5) break;
                        }
                    }
                }
                
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_OPENING] Checked {instanceCount} placed opening instances");
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_OPENING] Total unique opening parameters found: {openingParameters.Count}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_OPENING] Error getting opening parameters: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_OPENING] Stack trace: {ex.StackTrace}");
            }
            
            return openingParameters.OrderBy(p => p).ToList();
        }

        /// <summary>
        /// NEW METHOD: Get parameters from host elements (Walls, Floors, Ceilings) from linked architectural files
        /// </summary>
        private List<string> GetHostElementParametersFromLinkedFiles(List<Services.LinkedFileInfo> linkedFiles)
        {
            var hostParameters = new HashSet<string>();

            try
            {
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_HOST_LINKED] Starting to collect host element parameters from {linkedFiles.Count} linked files...");

                var hostCategories = new List<BuiltInCategory>
                {
                    BuiltInCategory.OST_Walls,
                    BuiltInCategory.OST_Floors,
                    BuiltInCategory.OST_Ceilings
                };

                foreach (var linkedFile in linkedFiles)
                {
                    try
                    {
                        var linkedDoc = linkedFile.LinkInstance?.GetLinkDocument();
                        if (linkedDoc != null)
                        {
                            System.Diagnostics.Debug.WriteLine($"[PARAMETER_HOST_LINKED] Processing linked file: {linkedFile.FileName} (type: {linkedFile.FileType})");

                            foreach (var category in hostCategories)
                            {
                                System.Diagnostics.Debug.WriteLine($"[PARAMETER_HOST_LINKED] Getting parameters for category: {category}");
                                var categoryParams = GetParametersForCategory(linkedDoc, category, includeInstanceParams: true, includeTypeParams: true);

                                System.Diagnostics.Debug.WriteLine($"[PARAMETER_HOST_LINKED] Found {categoryParams.Count} parameters for category {category}");
                                foreach (var param in categoryParams)
                                {
                                    hostParameters.Add(param.Name);
                                }
                            }

                            System.Diagnostics.Debug.WriteLine($"[PARAMETER_HOST_LINKED] Found parameters from linked file: {linkedFile.FileName}");
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[PARAMETER_HOST_LINKED] Could not get document for linked file: {linkedFile.FileName}");
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[PARAMETER_HOST_LINKED] Error reading linked file '{linkedFile.FileName}': {ex.Message}");
                    }
                }

                System.Diagnostics.Debug.WriteLine($"[PARAMETER_HOST_LINKED] Total unique host parameters from linked files: {hostParameters.Count}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_HOST_LINKED] Error: {ex.Message}");
            }

            return hostParameters.OrderBy(p => p).ToList();
        }

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

                System.Diagnostics.Debug.WriteLine($"[PARAMETER_OPENING] GetOpeningFamilies inspected {inspected} FamilySymbols, matched {openingFamilies.Count} families from the 4 specific opening families: {string.Join(", ", targetFamilyNames)}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_OPENING] Error getting opening families: {ex.Message}");
            }

            return openingFamilies;
        }

        private void UpdateSingleTabParameters(TabPage tabPage, List<string> mepParameters, List<string> openingParameters)
        {
            try
            {
                // Find the servicePanel in this tab
                var servicePanel = tabPage.Controls.OfType<System.Windows.Forms.Panel>().FirstOrDefault();

                System.Diagnostics.Debug.WriteLine($"[UI] Updating tab '{tabPage.Text}', panel found = {servicePanel != null}");

                // Additional debugging for Host tab
                if (tabPage.Text == "Host to Opening")
                {
                    System.Diagnostics.Debug.WriteLine($"[UI] Host tab controls: {tabPage.Controls.Count}");
                    foreach (System.Windows.Forms.Control control in tabPage.Controls)
                    {
                        System.Diagnostics.Debug.WriteLine($"[UI]   Control: {control.GetType().Name} - {control.Name}");
                    }
                }

                if (servicePanel != null)
                {
                    // Clear existing rows to prevent overlapping
                    var existingRows = servicePanel.Controls.OfType<System.Windows.Forms.Panel>().ToList();
                    foreach (var row in existingRows)
                    {
                        servicePanel.Controls.Remove(row);
                        row.Dispose();
                    }
                    
                    // Create automatic parameter rows with category-specific MEP parameters
                    CreateAutomaticParameterRows(servicePanel, mepParameters, openingParameters);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_UI] Error updating tab parameters: {ex.Message}");
            }
        }

        private void CreateAutomaticParameterRows(System.Windows.Forms.Panel servicePanel, List<string> mepParameters, List<string> openingParameters)
        {
            try
            {
                int rowHeight = 24;
                int top = 25;
                
                // Get the category name from the tab
                var tabPage = servicePanel.Parent as System.Windows.Forms.TabPage;
                var categoryName = tabPage?.Text ?? "";
                
                // Get specific parameters for this category
                var specificParameters = GetSpecificParametersForCategory(categoryName, mepParameters);
                
                // Create parameter rows for each specific parameter
                foreach (var param in specificParameters)
                {
                    var row = new System.Windows.Forms.Panel
                    {
                        Location = new System.Drawing.Point(8, top),
                        Size = new System.Drawing.Size(servicePanel.Width - 16, rowHeight),
                        Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
                    };
                    servicePanel.Controls.Add(row);

                    // MEP Parameter ComboBox (left side) - Pre-populated with specific parameter
                    var mepCombo = new System.Windows.Forms.ComboBox
                    {
                        Location = new System.Drawing.Point(0, 2),
                        Size = new System.Drawing.Size(120, 20),
                        DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList,
                        Tag = "mep"
                    };
                    
                    // Add all MEP parameters and select the specific one
                    mepCombo.Items.AddRange(mepParameters.ToArray());
                    mepCombo.SelectedItem = param;
                    row.Controls.Add(mepCombo);

                    // Opening Parameter ComboBox (right side) - Empty for user selection
                    var openingCombo = new System.Windows.Forms.ComboBox
                    {
                        Location = new System.Drawing.Point(130, 2),
                        Size = new System.Drawing.Size(row.Width - 130 - 30, 20),
                        Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right,
                        DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList,
                        Tag = "opening"
                    };
                    
                    // Add opening parameters but leave blank for user selection
                    openingCombo.Items.AddRange(openingParameters.ToArray());
                    openingCombo.Items.Insert(0, "<Select Opening Parameter>");
                    openingCombo.SelectedIndex = 0; // Leave blank
                    row.Controls.Add(openingCombo);

                    // Remove button
                    var removeBtn = new System.Windows.Forms.Button
                    {
                        Text = "×",
                        Location = new System.Drawing.Point(row.Width - 25, 1),
                        Size = new System.Drawing.Size(20, 20),
                        BackColor = System.Drawing.Color.FromArgb(255, 230, 230),
                        FlatStyle = System.Windows.Forms.FlatStyle.Flat,
                        Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold)
                    };
                    removeBtn.Click += (s, e) => {
                        servicePanel.Controls.Remove(row);
                        row.Dispose();
                    };
                    row.Controls.Add(removeBtn);

                    top += rowHeight + 3;
                }
                
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_UI] Created {specificParameters.Count} specific parameter rows for category '{categoryName}'");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_UI] Error creating specific parameter rows: {ex.Message}");
            }
        }

        private List<string> GetSpecificParametersForCategory(string categoryName, List<string> allMepParameters)
        {
            var specificParameters = new List<string>();
            
            try
            {
                // Define specific parameters for each category
                var categoryParams = categoryName.ToLower() switch
                {
                    "ducts" => new[] { "Reference Level", "Width", "Height", "System Type" },
                    "duct accessories" => new[] { "Reference Level", "Width", "Height", "System Type" },
                    "cable trays" => new[] { "Reference Level", "Width", "Height", "Service Type" },
                    "pipes" => new[] { "Reference Level", "Diameter", "System Type" },
                    _ => new string[0] // Unknown category
                };
                
                // Find matching parameters from the available MEP parameters
                foreach (var requiredParam in categoryParams)
                {
                    // Try exact match first
                    var exactMatch = allMepParameters.FirstOrDefault(p => 
                        string.Equals(p, requiredParam, StringComparison.OrdinalIgnoreCase));
                    
                    if (exactMatch != null)
                    {
                        specificParameters.Add(exactMatch);
                    }
                    else
                    {
                        // Try partial match (case insensitive)
                        var partialMatch = allMepParameters.FirstOrDefault(p => 
                            p.Contains(requiredParam, StringComparison.OrdinalIgnoreCase) ||
                            requiredParam.Contains(p, StringComparison.OrdinalIgnoreCase));
                        
                        if (partialMatch != null)
                        {
                            specificParameters.Add(partialMatch);
                        }
                        else
                        {
                            // If no match found, add the required parameter anyway
                            // This ensures the UI shows the expected parameter even if not found in linked files
                            specificParameters.Add(requiredParam);
                        }
                    }
                }
                
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_SPECIFIC] Category '{categoryName}': Found {specificParameters.Count} specific parameters");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_SPECIFIC] Error getting specific parameters for '{categoryName}': {ex.Message}");
            }
            
            return specificParameters;
        }
    }
}