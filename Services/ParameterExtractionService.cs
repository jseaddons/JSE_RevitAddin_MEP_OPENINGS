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
        private Document _document; // Store document reference for row creation

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
                // Store document reference for row creation
                _document = document;

                // Log to main log file
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] [PARAMETER_SERVICE] Starting category-specific population for {selectedCategories.Count} categories\n");
                
                // Get opening parameters using the corrected harvest routine
                var openingParameters = GetCurrentOpeningParameters(document);

                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] [PARAMETER_SERVICE] Using cached opening parameters: {openingParameters.Count} parameters\n");

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

                        UpdateSingleTabParameters(matchingTab, categorySpecificParameters, combinedOpeningParameters, document);
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
                    UpdateSingleTabParameters(hostTab, linkedHostParameters, combinedOpeningParameters, document);
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[HOST] WARNING: 'Host to Opening' tab not found!");
                }

                // Handle Reference tab (MEP parameters)
                var referenceTab = FindTabByName(serviceParameterTabs, "Reference Element to Openings");
                if (referenceTab != null)
                {
                    System.Diagnostics.Debug.WriteLine($"[REFERENCE] Tab '{referenceTab.Text}' found, populating with MEP parameters");

                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt", $"[{DateTime.Now}] [PARAMETER_SERVICE] Updating 'Reference Element to Openings' tab with MEP parameters\n");

                    // For Reference tab: Source = MEP parameters, Target = opening parameters
                    // We need to harvest MEP parameters since they're not in the cache yet
                    var mepParameters = HarvestMepParametersFromDocument(document);
                    UpdateSingleTabParameters(referenceTab, mepParameters, combinedOpeningParameters, document);
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[REFERENCE] WARNING: 'Reference Element to Openings' tab not found!");
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

    public List<string> GetParametersForSpecificCategoryFromLinkedFiles(string categoryName, List<Services.LinkedFileInfo> linkedFiles)
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
        /// MASTER FIX: Bootstrap opening parameters by temporarily placing instances to force Revit to bind shared parameters
        /// </summary>
        public List<string> GetCurrentOpeningParameters(Document doc)
        {
            var set = new HashSet<string>();

            var targetFamilies = new[] {
                    "RectangularOpeningOnWall",
                    "RectangularOpeningOnSlab",
                    "CircularOpeningOnWall",
                    "CircularOpeningOnSlab"
                };

            // Check if we can safely start a transaction (not already in a transaction context)
            bool canStartTransaction = true;
            try
            {
                // Try to check if document is modifiable (this will fail if already in transaction)
                var testModifiable = doc.IsModifiable;
            }
            catch
            {
                // If we can't check modifiable status, we're likely in a transaction context
                canStartTransaction = false;
            }

            if (canStartTransaction)
            {
                using (var t = new Transaction(doc, "BootstrapOpeningParams"))
                {
                    t.Start();

                // 1.  Ensure at least one instance of each family exists
                foreach (var name in targetFamilies)
                {
                    // Already placed?  Skip.
                    bool alreadyPlaced = new FilteredElementCollector(doc)
                                        .OfClass(typeof(FamilyInstance))
                                        .WhereElementIsNotElementType()
                                        .Cast<FamilyInstance>()
                                        .Any(fi => fi.Symbol.Family.Name.Contains(name));

                    if (alreadyPlaced) continue;

                    var symbol = new FilteredElementCollector(doc)
                                .OfClass(typeof(FamilySymbol))
                                .Cast<FamilySymbol>()
                                .FirstOrDefault(fs => fs.Family.Name.Contains(name));

                    if (symbol == null) continue;

                    if (!symbol.IsActive) symbol.Activate();

                    // Get the first level in the document
                    var level = new FilteredElementCollector(doc)
                               .OfClass(typeof(Level))
                               .Cast<Level>()
                               .FirstOrDefault();

                    if (level != null)
                    {
                        // Place at origin – will be deleted in the same transaction
                        doc.Create.NewFamilyInstance(XYZ.Zero, symbol, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                    }
                }

                // 2.  Collect parameters from the freshly bound instances
                var instances = new FilteredElementCollector(doc)
                               .OfClass(typeof(FamilyInstance))
                               .WhereElementIsNotElementType()
                               .Cast<FamilyInstance>()
                               .Where(fi => targetFamilies.Any(n => fi.Symbol.Family.Name.Contains(n)))
                               .ToList();

                foreach (var fi in instances)
                    foreach (Parameter p in fi.Parameters)
                        if (!string.IsNullOrEmpty(p.Definition?.Name))
                            set.Add(p.Definition.Name);

                // 3.  Delete the temporary instances – binding stays
                foreach (var fi in instances)
                    doc.Delete(fi.Id);

                    t.Commit();
                }
            }
            else
            {
                // Can't start transaction - skip bootstrap and just harvest existing symbols
                System.Diagnostics.Debug.WriteLine($"[BOOTSTRAP] Cannot start transaction - skipping bootstrap, using existing parameters only");
            }

            // 4.  Also harvest the symbols (catches any type-only parameters)
            foreach (var name in targetFamilies)
            {
                var symbol = new FilteredElementCollector(doc)
                            .OfClass(typeof(FamilySymbol))
                            .Cast<FamilySymbol>()
                            .FirstOrDefault(fs => fs.Family.Name.Contains(name));

                if (symbol == null) continue;

                foreach (Parameter p in symbol.Parameters)
                    if (!string.IsNullOrEmpty(p.Definition?.Name))
                        set.Add(p.Definition.Name);
            }

            System.Diagnostics.Debug.WriteLine($"[BOOTSTRAP] Returning {set.Count} opening parameters");
            return set.OrderBy(p => p).ToList();
        }

        /// <summary>
        /// NEW METHOD: Get parameters from host elements (Walls, Floors, Ceilings) from linked architectural files
        /// </summary>
        public List<string> GetHostElementParametersFromLinkedFiles(List<LinkedFileInfo> linkedFiles)
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

        /// <summary>
        /// Harvests MEP parameters from the document (used for Reference tab when cache is not available)
        /// </summary>
    public List<string> HarvestMepParametersFromDocument(Document document)
        {
            var mepParameters = new HashSet<string>();

            try
            {
                System.Diagnostics.Debug.WriteLine($"[MEP_HARVEST] Starting to harvest MEP parameters from document");

                // Get selected MEP categories
                var selectedCategories = new List<string> { "Ducts", "Pipes", "Cable Trays", "Duct Accessories" };

                // Use LinkedFileService to get linked files (same as LoadMepParameters)
                var linkedFileService = new Services.LinkedFileService();
                var linkedFiles = linkedFileService.GetLinkedFiles(document);

                System.Diagnostics.Debug.WriteLine($"[MEP_HARVEST] Found {linkedFiles.Count} linked files");

                if (linkedFiles.Count > 0)
                {
                    // Get parameters from linked files
                    foreach (var linkedFile in linkedFiles)
                    {
                        try
                        {
                            System.Diagnostics.Debug.WriteLine($"[MEP_HARVEST] Processing linked file: {linkedFile.FileName} (type: {linkedFile.FileType})");
                            var linkedDoc = linkedFile.LinkInstance?.GetLinkDocument();
                            if (linkedDoc != null)
                            {
                                // Convert string categories to enum
                                var mepCategories = selectedCategories
                                    .Select(cat => GetMepCategoryFromString(cat))
                                    .Where(cat => cat.HasValue)
                                    .Select(cat => cat.Value!)
                                    .ToList();

                                System.Diagnostics.Debug.WriteLine($"[MEP_HARVEST] Getting parameters for categories: {string.Join(", ", mepCategories)}");
                                var parameters = GetParametersForMepCategories(linkedDoc, mepCategories);
                                var parameterNames = parameters.Select(p => p.Name).ToList();

                                System.Diagnostics.Debug.WriteLine($"[MEP_HARVEST] Found {parameterNames.Count} parameters from linked file");

                                // Add unique parameters
                                foreach (var paramName in parameterNames)
                                {
                                    if (!string.IsNullOrEmpty(paramName))
                                    {
                                        mepParameters.Add(paramName);
                                    }
                                }
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"[MEP_HARVEST] Could not get document for linked file: {linkedFile.FileName}");
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error getting parameters from linked file '{linkedFile.FileName}': {ex.Message}");
                        }
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[MEP_HARVEST] No linked files found, checking current document");
                    // Fallback to current document if no linked files
                    var mepCategories = selectedCategories
                        .Select(cat => GetMepCategoryFromString(cat))
                        .Where(cat => cat.HasValue)
                        .Select(cat => cat.Value!)
                        .ToList();

                    System.Diagnostics.Debug.WriteLine($"[MEP_HARVEST] Getting parameters from current document for categories: {string.Join(", ", mepCategories)}");
                    var parameters = GetParametersForMepCategories(document, mepCategories);
                    var parameterNames = parameters.Select(p => p.Name).ToList();

                    System.Diagnostics.Debug.WriteLine($"[MEP_HARVEST] Found {parameterNames.Count} parameters from current document");

                    foreach (var paramName in parameterNames)
                    {
                        if (!string.IsNullOrEmpty(paramName))
                        {
                            mepParameters.Add(paramName);
                        }
                    }
                }

                System.Diagnostics.Debug.WriteLine($"[MEP_HARVEST] Total MEP parameters harvested: {mepParameters.Count}");
                return mepParameters.OrderBy(p => p).ToList();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[MEP_HARVEST] Error harvesting MEP parameters: {ex.Message}");
                return new List<string>();
            }
        }

        /// <summary>
        /// Convert category name string to MepCategory enum (helper for HarvestMepParametersFromDocument)
        /// </summary>
        private MepCategory? GetMepCategoryFromString(string categoryName)
        {
            return categoryName.ToLower() switch
            {
                "pipes" or "pipe" => MepCategory.Pipes,
                "ducts" or "duct" => MepCategory.Ducts,
                "cable trays" or "cable tray" => MepCategory.CableTrays,
                "conduits" or "conduit" => MepCategory.CableTrays, // Conduits not available, use CableTrays instead
                "duct accessories" or "duct accessory" => MepCategory.DuctAccessories,
                "duct fittings" or "duct fitting" => MepCategory.DuctAccessories, // DuctFittings not available, use DuctAccessories instead
                _ => null
            };
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

        private void UpdateSingleTabParameters(TabPage tabPage, List<string> mepParameters, List<string> openingParameters, Document document)
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
                    CreateAutomaticParameterRows(servicePanel, mepParameters, openingParameters, document);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_UI] Error updating tab parameters: {ex.Message}");
            }
        }

        /// <summary>
        /// Add this method to handle dynamic row addition with proper parameter population
        /// This should be called by your "Add Row" button click handler
        /// </summary>
        public void AddNewParameterRow(System.Windows.Forms.Panel servicePanel, List<string> mepParameters, List<string> openingParameters, Document document)
        {
            try
            {
                LoggingConfiguration.ConditionalAppendAllText(
                    @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt",
                    $"[{DateTime.Now}] [ADD_ROW] Starting to add new row{Environment.NewLine}");

                int rowHeight = 24;

                // Find the last row's Y position
                var existingRows = servicePanel.Controls.OfType<System.Windows.Forms.Panel>().ToList();
                int top = existingRows.Count > 0
                    ? existingRows.Max(r => r.Location.Y + r.Height) + 3
                    : 25;

                // Create new row panel
                var row = new System.Windows.Forms.Panel
                {
                    Location = new System.Drawing.Point(8, top),
                    Size = new System.Drawing.Size(servicePanel.Width - 16, rowHeight),
                    Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
                };
                servicePanel.Controls.Add(row);

                // MEP Parameter ComboBox (left side)
                var mepCombo = new System.Windows.Forms.ComboBox
                {
                    Location = new System.Drawing.Point(0, 2),
                    Size = new System.Drawing.Size(120, 20),
                    DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList,
                    Tag = "mep"
                };

                // Populate MEP combo
                mepCombo.Items.AddRange(mepParameters.ToArray());
                mepCombo.IntegralHeight = false;
                mepCombo.MaxDropDownItems = 25;
                mepCombo.DropDownHeight = 400;

                if (mepCombo.Items.Count > 0)
                    mepCombo.SelectedIndex = 0;

                row.Controls.Add(mepCombo);

                LoggingConfiguration.ConditionalAppendAllText(
                    @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt",
                    $"[{DateTime.Now}] [ADD_ROW] MEP combo populated with {mepCombo.Items.Count} items{Environment.NewLine}");

                // Opening Parameter ComboBox (right side) - THIS IS THE CRITICAL PART
                var openingCombo = new System.Windows.Forms.ComboBox
                {
                    Location = new System.Drawing.Point(130, 2),
                    Size = new System.Drawing.Size(row.Width - 130 - 30, 20),
                    Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right,
                    DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList,
                    Tag = "opening"
                };

                LoggingConfiguration.ConditionalAppendAllText(
                    @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt",
                    $"[{DateTime.Now}] [ADD_ROW] About to populate opening combo - received {openingParameters.Count} parameters from caller{Environment.NewLine}");

                // CRITICAL FIX: Use the LIVE bootstrap routine instead of cached parameters
                var liveOpeningParams = GetCurrentOpeningParameters(document);

                LoggingConfiguration.ConditionalAppendAllText(
                    @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt",
                    $"[{DateTime.Now}] [ADD_ROW] Live bootstrap returned {liveOpeningParams.Count} opening parameters{Environment.NewLine}");

                // Add opening parameters
                openingCombo.Items.AddRange(liveOpeningParams.Cast<object>().ToArray());
                openingCombo.Items.Insert(0, "<Select Opening Parameter>");

                // Configure dropdown display
                openingCombo.IntegralHeight = false;
                openingCombo.MaxDropDownItems = 25;
                openingCombo.DropDownHeight = 400;
                openingCombo.SelectedIndex = 0;

                row.Controls.Add(openingCombo);

                LoggingConfiguration.ConditionalAppendAllText(
                    @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt",
                    $"[{DateTime.Now}] [ADD_ROW] Opening combo populated - Items.Count = {openingCombo.Items.Count}{Environment.NewLine}");

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

                LoggingConfiguration.ConditionalAppendAllText(
                    @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt",
                    $"[{DateTime.Now}] [ADD_ROW] Row successfully added to panel{Environment.NewLine}");
            }
            catch (Exception ex)
            {
                LoggingConfiguration.ConditionalAppendAllText(
                    @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt",
                    $"[{DateTime.Now}] [ADD_ROW] ERROR: {ex.Message}{Environment.NewLine}{ex.StackTrace}{Environment.NewLine}");
            }
        }

        private void CreateAutomaticParameterRows(System.Windows.Forms.Panel servicePanel, List<string> mepParameters, List<string> openingParameters, Document document)
        {
            try
            {
                // =====  DIAGNOSTIC – DO NOT DELETE  =====
                var diagOpeningParams = GetCurrentOpeningParameters(document);
                LoggingConfiguration.ConditionalAppendAllText(
                    @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt",
                    $"[LIVE-DIAG] {nameof(CreateAutomaticParameterRows)} about to fill Opening combo with {diagOpeningParams.Count} items{Environment.NewLine}");
                // =======================================

                int rowHeight = 24;
                int top = 25;
                
                // Get the category name from the tab
                var tabPage = servicePanel.Parent as System.Windows.Forms.TabPage;
                var categoryName = tabPage?.Text ?? "";
                
                // Get specific parameters for this category
                var specificParameters = GetSpecificParametersForCategory(categoryName, mepParameters);

                // =====  DIAGNOSTIC – DO NOT DELETE  =====
                LoggingConfiguration.ConditionalAppendAllText(
                    @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt",
                    $"[ROW-CHECK] About to create {specificParameters.Count} rows for category '{categoryName}'{Environment.NewLine}");
                // =======================================
                
                    // Get opening parameters using the live bootstrap routine - always fresh
                    var liveOpeningParams = GetCurrentOpeningParameters(document);
                
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

                    // Ensure MEP combo also has proper dropdown settings for large parameter lists
                    mepCombo.IntegralHeight = false;
                    mepCombo.MaxDropDownItems = 25;
                    mepCombo.DropDownHeight = 400;
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
                    
                        // Add opening parameters using the live bootstrap routine
                    openingCombo.Items.AddRange(liveOpeningParams.Cast<object>().ToArray());
                        openingCombo.Items.Insert(0, "<Select Opening Parameter>");

                    // =====  DIAGNOSTIC – DO NOT DELETE  =====
                    LoggingConfiguration.ConditionalAppendAllText(
                        @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt",
                        $"[REF-CHECK] Opening combo now contains {openingCombo.Items.Count - 1} items (after AddRange){Environment.NewLine}");
                    // =======================================

                    // Force dropdown to show all items with scrolling (fix UI truncation)
                    openingCombo.IntegralHeight = false; // Allow partial items for better scrolling
                    openingCombo.MaxDropDownItems = 25; // Show 25 items at a time with scroll bar (increased from 20)
                    openingCombo.DropDownHeight = 400; // Increase dropdown height to accommodate more items

                    // =====  DIAGNOSTIC – DO NOT DELETE  =====
                    LoggingConfiguration.ConditionalAppendAllText(
                        @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt",
                        $"[UI-TRUTH] Added-row combo count = {openingCombo.Items.Count - 1}{Environment.NewLine}");
                    // =======================================

                        // Set default opening parameter based on the MEP parameter selected
                        SetDefaultOpeningParameterSelection(openingCombo, param);
                        if (openingCombo.SelectedIndex <= 0) // If no matching parameter found
                            openingCombo.SelectedIndex = 0; // Leave as "Select" option
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

                // =====  DIAGNOSTIC – DO NOT DELETE  =====
                var finalOpeningParams = GetCurrentOpeningParameters(document);
                LoggingConfiguration.ConditionalAppendAllText(
                    @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt",
                    $"[REF-FINAL] Method completed - bootstrap returns {finalOpeningParams.Count} opening parameters{Environment.NewLine}");
                // =======================================
            }
            catch (Exception ex)
            {
                LoggingConfiguration.ConditionalAppendAllText(
                    @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\logger_debug.txt",
                    $"[LIVE-DIAG-ERROR] {ex.GetType().Name}: {ex.Message}{Environment.NewLine}");
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_UI] Error creating specific parameter rows: {ex.Message}");
            }
        }

        /// <summary>
        /// Sets the default opening parameter selection based on the MEP parameter
        /// </summary>
        private void SetDefaultOpeningParameterSelection(System.Windows.Forms.ComboBox openingCombo, string mepParameter)
        {
            try
            {
                // Define mapping from MEP parameters to opening parameters
                var parameterMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    { "Size", "MEP Size" },
                    { "Width", "MEP Size" },
                    { "Height", "MEP Size" },
                    { "Diameter", "MEP Size" },
                    { "System Type", "MEP System Type" },
                    { "System Abbreviation", "MEP System Abbreviation" },
                    { "Service Type", "MEP System Type" }, // For cable trays
                    { "Reference Level", "MEP Mark" } // Default fallback
                };

                // Find the corresponding opening parameter
                if (parameterMapping.TryGetValue(mepParameter, out var openingParameter))
                {
                    var index = openingCombo.Items.IndexOf(openingParameter);
                    if (index > 0) // > 0 because index 0 is "<Select Opening Parameter>"
                    {
                        openingCombo.SelectedIndex = index;
                        System.Diagnostics.Debug.WriteLine($"[DEFAULT_SELECTIONS] Set opening parameter '{openingParameter}' for MEP parameter '{mepParameter}'");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEFAULT_SELECTIONS] Error setting opening parameter selection: {ex.Message}");
            }
        }

        private List<string> GetSpecificParametersForCategory(string categoryName, List<string> allMepParameters)
        {
            var specificParameters = new List<string>();

            try
            {
                // Define specific parameters for each category (updated defaults)
                var categoryParams = categoryName.ToLower() switch
                {
                    "ducts" => new[] { "Size", "System Type", "System Abbreviation" },
                    "duct accessories" => new[] { "Size", "System Type", "System Abbreviation" },
                    "cable trays" => new[] { "Size", "Service Type", "System Abbreviation" },
                    "pipes" => new[] { "Size", "System Type", "System Abbreviation" },
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
