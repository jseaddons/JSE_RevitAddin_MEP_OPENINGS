using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Views;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// Command for opening the Parameter Transfer dialog
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ParameterTransferCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiApp = commandData.Application;
                var uiDoc = uiApp.ActiveUIDocument;
                var doc = uiDoc.Document;
                
                // Check if document is available
                if (doc == null)
                {
                    TaskDialog.Show("Error", "No active document found.");
                    return Result.Failed;
                }
                
                // Open parameter transfer dialog
                using (var dialog = new ParameterTransferDialog(doc))
                {
                    if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        var configuration = dialog.GetConfiguration();
                        
                        // Get selected openings
                        var selectedOpeningIds = GetSelectedOpenings(uiDoc);
                        if (selectedOpeningIds.Count == 0)
                        {
                            selectedOpeningIds = GetAllOpenings(doc);
                        }
                        
                        if (selectedOpeningIds.Count == 0)
                        {
                            TaskDialog.Show("No Openings", "No openings found in the project.");
                            return Result.Cancelled;
                        }
                        
                        // Execute parameter transfer using command-owned transaction (preferred)
                        var transferService = new ParameterTransferService();
                        ParameterTransferResult result;

                        using (var tx = new Transaction(doc, "Execute Parameter Transfer Configuration"))
                        {
                            tx.Start();

                            // If a WarningSwallower or failures preprocessor helper is available in the project,
                            // it should be set here. We try to find and use it reflectively to avoid hard dependency.
                            try
                            {
                                var fpType = typeof(Autodesk.Revit.DB.IFailuresPreprocessor);
                                // Project-specific WarningSwallower is optional - ignored if not present
                                var wsType = System.Type.GetType("JSE_RevitAddin_MEP_OPENINGS.Services.WarningSwallower, JSE_RevitAddin_MEP_OPENINGS");
                                if (wsType != null && fpType.IsAssignableFrom(wsType))
                                {
                                    var wsInstance = Activator.CreateInstance(wsType) as IFailuresPreprocessor;
                                    if (wsInstance != null)
                                    {
                                        // Failure handling options are set on the Transaction, not the Document
                                        var fho = tx.GetFailureHandlingOptions();
                                        fho.SetFailuresPreprocessor(wsInstance);
                                        tx.SetFailureHandlingOptions(fho);
                                    }
                                }
                            }
                            catch { /* ignore if not present */ }

                            // Call the in-transaction implementation
                            result = transferService.ExecuteTransferConfigurationInTransaction(doc, selectedOpeningIds, configuration);

                            tx.Commit();
                        }
                        
                        // Show result
                        if (result.Success)
                        {
                            var messageText = $"Parameter transfer completed successfully!\n\n" +
                                            $"Transferred: {result.TransferredCount}\n" +
                                            $"Failed: {result.FailedCount}\n" +
                                            $"Warnings: {result.Warnings.Count}";
                            
                            if (result.Warnings.Count > 0)
                            {
                                messageText += "\n\nWarnings:\n" + string.Join("\n", result.Warnings.Take(5));
                                if (result.Warnings.Count > 5)
                                {
                                    messageText += $"\n... and {result.Warnings.Count - 5} more warnings.";
                                }
                            }
                            
                            TaskDialog.Show("Success", messageText);
                        }
                        else
                        {
                            var errorText = $"Parameter transfer failed!\n\n{result.Message}";
                            
                            if (result.Errors.Count > 0)
                            {
                                errorText += "\n\nErrors:\n" + string.Join("\n", result.Errors.Take(5));
                                if (result.Errors.Count > 5)
                                {
                                    errorText += $"\n... and {result.Errors.Count - 5} more errors.";
                                }
                            }
                            
                            TaskDialog.Show("Error", errorText);
                        }
                    }
                }
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error executing parameter transfer command: {ex.Message}";
                return Result.Failed;
            }
        }
        
        private List<ElementId> GetSelectedOpenings(UIDocument uiDoc)
        {
            var selectedIds = new List<ElementId>();
            
            try
            {
                var selection = uiDoc.Selection;
                var selectedElementIds = selection.GetElementIds();
                
                foreach (var elementId in selectedElementIds)
                {
                    var element = uiDoc.Document.GetElement(elementId);
                    if (IsOpeningElement(element))
                    {
                        selectedIds.Add(elementId);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting selected openings: {ex.Message}");
            }
            
            return selectedIds;
        }
        
        private List<ElementId> GetAllOpenings(Document doc)
        {
            var openingIds = new List<ElementId>();
            
            try
            {
                // Get opening categories
                var openingCategories = new List<BuiltInCategory>
                {
                    BuiltInCategory.OST_GenericModel, // Generic openings
                    BuiltInCategory.OST_GenericAnnotation, // OST_WallOpening not available, use GenericAnnotation instead   // Wall openings
                    BuiltInCategory.OST_FloorOpening,  // Floor openings
                    BuiltInCategory.OST_CeilingOpening // Ceiling openings
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
                System.Diagnostics.Debug.WriteLine($"Error getting all openings: {ex.Message}");
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
    }
    
    /// <summary>
    /// Command for opening the Parameter Renaming dialog
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ParameterRenamingCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Open parameter renaming dialog
                using (var dialog = new ParameterRenamingDialog())
                {
                    dialog.ShowDialog();
                }
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error executing parameter renaming command: {ex.Message}";
                return Result.Failed;
            }
        }
    }
    
    /// <summary>
    /// Command for executing parameter transfer with predefined configuration
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class QuickParameterTransferCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiApp = commandData.Application;
                var uiDoc = uiApp.ActiveUIDocument;
                var doc = uiDoc.Document;
                
                // Check if document is available
                if (doc == null)
                {
                    TaskDialog.Show("Error", "No active document found.");
                    return Result.Failed;
                }
                
                // Get selected openings
                var selectedOpeningIds = GetSelectedOpenings(uiDoc);
                if (selectedOpeningIds.Count == 0)
                {
                    selectedOpeningIds = GetAllOpenings(doc);
                }
                
                if (selectedOpeningIds.Count == 0)
                {
                    TaskDialog.Show("No Openings", "No openings found in the project.");
                    return Result.Cancelled;
                }
                
                // Create predefined configuration
                var configuration = CreatePredefinedConfiguration();
                
                // Execute parameter transfer
                var transferService = new ParameterTransferService();
                var result = transferService.ExecuteTransferConfiguration(doc, selectedOpeningIds, configuration);
                
                // Show result
                if (result.Success)
                {
                    var messageText = $"Quick parameter transfer completed!\n\n" +
                                    $"Transferred: {result.TransferredCount}\n" +
                                    $"Failed: {result.FailedCount}";
                    
                    TaskDialog.Show("Success", messageText);
                }
                else
                {
                    TaskDialog.Show("Error", $"Quick parameter transfer failed: {result.Message}");
                }
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error executing quick parameter transfer command: {ex.Message}";
                return Result.Failed;
            }
        }
        
        private List<ElementId> GetSelectedOpenings(UIDocument uiDoc)
        {
            var selectedIds = new List<ElementId>();
            
            try
            {
                var selection = uiDoc.Selection;
                var selectedElementIds = selection.GetElementIds();
                
                foreach (var elementId in selectedElementIds)
                {
                    var element = uiDoc.Document.GetElement(elementId);
                    if (IsOpeningElement(element))
                    {
                        selectedIds.Add(elementId);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting selected openings: {ex.Message}");
            }
            
            return selectedIds;
        }
        
        private List<ElementId> GetAllOpenings(Document doc)
        {
            var openingIds = new List<ElementId>();
            
            try
            {
                // Get opening categories
                var openingCategories = new List<BuiltInCategory>
                {
                    BuiltInCategory.OST_GenericModel, // Generic openings
                    BuiltInCategory.OST_GenericAnnotation, // OST_WallOpening not available, use GenericAnnotation instead   // Wall openings
                    BuiltInCategory.OST_FloorOpening,  // Floor openings
                    BuiltInCategory.OST_CeilingOpening // Ceiling openings
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
                System.Diagnostics.Debug.WriteLine($"Error getting all openings: {ex.Message}");
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
        
        private ParameterTransferConfiguration CreatePredefinedConfiguration()
        {
            var configuration = new ParameterTransferConfiguration();
            
            // Add predefined mappings
            var mappingService = new ParameterMappingService();
            var predefinedMappings = mappingService.GetPredefinedMappings();
            configuration.Mappings.AddRange(predefinedMappings);
            
            // Enable model name transfer
            configuration.TransferModelNames = true;
            configuration.ModelNameParameter = "Model_Name";
            
            // Load predefined renaming conditions
            var renamingService = new ParameterRenamingService();
            renamingService.LoadPredefinedRenamingConditions();
            configuration.RenamingConditions = renamingService.GetAllRenamingConditions();
            
            return configuration;
        }
    }
}
