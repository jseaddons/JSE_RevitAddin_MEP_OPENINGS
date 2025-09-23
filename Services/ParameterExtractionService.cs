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
                
                // Get opening parameters (global - same for all tabs)
                var openingParameters = GetCurrentOpeningParameters(document);
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] [PARAMETER_SERVICE] Found {openingParameters.Count} opening parameters\n");
                
                // Get linked files for parameter extraction
                var linkedFileService = new Services.LinkedFileService();
                var linkedFiles = linkedFileService.GetLinkedFiles(document);
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] [PARAMETER_SERVICE] Found {linkedFiles.Count} linked files\n");
                
                // Update each tab that matches selected categories
                foreach (var selectedCategory in selectedCategories)
                {
                    var matchingTab = FindTabByCategory(serviceParameterTabs, selectedCategory);
                    if (matchingTab != null)
                    {
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] [PARAMETER_SERVICE] Updating tab '{matchingTab.Text}' for category '{selectedCategory}'\n");
                        
                        var categorySpecificParameters = GetParametersForSpecificCategoryFromLinkedFiles(selectedCategory, linkedFiles);
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] [PARAMETER_SERVICE] Found {categorySpecificParameters.Count} parameters for category '{selectedCategory}' from linked files\n");
                        
                        UpdateSingleTabParameters(matchingTab, categorySpecificParameters, openingParameters);
                    }
                    else
                    {
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] [PARAMETER_SERVICE] No tab found for category '{selectedCategory}'\n");
                    }
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
                
                if (Enum.TryParse<MepCategory>(categoryName, out var category))
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

        private List<string> GetCurrentOpeningParameters(Document document)
        {
            var openingParameters = new List<string>();
            
            try
            {
                var openingFamilies = GetOpeningFamilies(document);
                foreach (var family in openingFamilies)
                {
                    foreach (Parameter param in family.Parameters)
                    {
                        if (!string.IsNullOrEmpty(param.Definition.Name) && !openingParameters.Contains(param.Definition.Name))
                        {
                            openingParameters.Add(param.Definition.Name);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_OPENING] Error getting opening parameters: {ex.Message}");
            }
            
            return openingParameters;
        }

        private List<FamilySymbol> GetOpeningFamilies(Document document)
        {
            var openingFamilies = new List<FamilySymbol>();
            
            try
            {
                var collector = new FilteredElementCollector(document)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .Where(fs => fs.Family.Name.Contains("Opening", StringComparison.OrdinalIgnoreCase));
                
                openingFamilies = collector.ToList();
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
                // EXACT COPY OF WORKING LOGIC FROM OLD METHOD
                var allComboBoxes = new List<ComboBox>();
                
                // Step 1: Find the servicePanel in this tab
                var servicePanel = tabPage.Controls.OfType<System.Windows.Forms.Panel>().FirstOrDefault();
                
                if (servicePanel != null)
                {
                    // Step 2: Find all row panels in the servicePanel
                    var rowPanels = servicePanel.Controls.OfType<System.Windows.Forms.Panel>().ToList();
                    
                    // Step 3: Find ComboBoxes in each row panel
                    foreach (var rowPanel in rowPanels)
                    {
                        var rowComboBoxes = rowPanel.Controls.OfType<ComboBox>().ToList();
                        allComboBoxes.AddRange(rowComboBoxes);
                    }
                }
                
                // Check if ComboBoxes are actually populated with meaningful content
                bool hasEmptyComboBoxes = allComboBoxes.Any(cb => cb.Items.Count == 0 || cb.SelectedItem == null);
                
                // If no ComboBoxes found OR there are empty ComboBoxes, create default parameter rows
                if (allComboBoxes.Count == 0 || hasEmptyComboBoxes)
                {
                    // Clear existing rows to prevent overlapping
                    if (servicePanel != null)
                    {
                        var existingRows = servicePanel.Controls.OfType<System.Windows.Forms.Panel>().ToList();
                        foreach (var row in existingRows)
                        {
                            servicePanel.Controls.Remove(row);
                            row.Dispose();
                        }
                    }
                    
                    // Skip creating default rows for now - just populate existing ComboBoxes
                }
                else
                {
                    // Update existing ComboBoxes
                    foreach (var comboBox in allComboBoxes)
                    {
                        // Store current selection before clearing
                        var currentSelection = comboBox.SelectedItem?.ToString();
                        
                        // Clear existing items
                        comboBox.Items.Clear();
                        
                        // Determine which parameters to add based on ComboBox position or tag
                        if (comboBox.Tag?.ToString()?.Contains("mep") == true || 
                            comboBox.Location.X < 100) // Left side = MEP parameters
                        {
                            comboBox.Items.AddRange(mepParameters.ToArray());
                            
                            // Restore selection if it still exists
                            if (!string.IsNullOrEmpty(currentSelection) && mepParameters.Contains(currentSelection))
                            {
                                comboBox.SelectedItem = currentSelection;
                            }
                        }
                        else // Right side = Opening parameters
                        {
                            comboBox.Items.AddRange(openingParameters.ToArray());
                            
                            // Restore selection if it still exists
                            if (!string.IsNullOrEmpty(currentSelection) && openingParameters.Contains(currentSelection))
                            {
                                comboBox.SelectedItem = currentSelection;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Silent fail - just populate what we can
            }
        }
    }
}