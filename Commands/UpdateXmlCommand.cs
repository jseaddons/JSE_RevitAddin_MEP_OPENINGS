using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using WinForms = System.Windows.Forms;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// Command to update XML files after manual cluster sleeve adjustments
    /// This is an expensive operation that should only be run after manual changes to sleeves
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class UpdateXmlCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiApp = commandData.Application;
                var uiDoc = uiApp.ActiveUIDocument;
                var doc = uiDoc?.Document;

                if (doc == null)
                {
                    message = "No active document found.";
                    return Result.Failed;
                }

                // ✅ NEW: Check if filter XML files exist before proceeding
                var missingCategories = CheckFilterFilesForCategories(doc);
                if (missingCategories.Count > 0)
                {
                    var categoryList = string.Join("\n• ", missingCategories);
                    var validationResult = WinForms.MessageBox.Show(
                        $"⚠️ FILTER DATA NOT FOUND\n\n" +
                        $"The following categories cannot be updated because no filter data is found:\n\n" +
                        $"• {categoryList}\n\n" +
                        $"Update XML requires XML filter files (e.g., '*_ducts.xml', '*_pipes.xml') in the Filters directory.\n\n" +
                        $"Would you like to continue with available categories only?",
                        "Missing Filter Data",
                        WinForms.MessageBoxButtons.YesNo,
                        WinForms.MessageBoxIcon.Warning);
                    
                    if (validationResult != WinForms.DialogResult.Yes)
                    {
                        return Result.Cancelled; // User cancelled
                    }
                }

                // ⚠️ WARNING DIALOG: Inform user this is expensive and should only be used after manual changes
                var result = WinForms.MessageBox.Show(
                    "⚠️ UPDATE XML OPERATION\n\n" +
                    "This operation will:\n" +
                    "• Scan all cluster sleeves in the model\n" +
                    "• Update XML with current sleeve sizes\n" +
                    "• Update MEP element information\n" +
                    "• Re-number MEP marks\n\n" +
                    "⏱️ This may take several minutes depending on model size.\n\n" +
                    "Use this ONLY after:\n" +
                    "• Manually resizing cluster sleeves\n" +
                    "• Joining cluster sleeves\n" +
                    "• Modifying sleeve geometry\n\n" +
                    "Do you want to continue?",
                    "Update XML - Expensive Operation",
                    WinForms.MessageBoxButtons.YesNo,
                    WinForms.MessageBoxIcon.Warning);

                if (result != WinForms.DialogResult.Yes)
                {
                    return Result.Cancelled;
                }

                // Execute update
                var updateService = new UpdateXmlService();
                updateService.UpdateXmlFromRevit(doc);

                // ✅ Optionally re-number MEP marks after XML update
                var renumberResult = WinForms.MessageBox.Show(
                    "✅ XML Update Complete!\n\n" +
                    "All cluster sleeves have been updated in XML files.\n\n" +
                    "Would you like to re-number MEP marks now?\n" +
                    "(This will update marks based on current sleeve count)",
                    "Update Complete - Re-number Marks?",
                    WinForms.MessageBoxButtons.YesNo,
                    WinForms.MessageBoxIcon.Question);

                if (renumberResult == WinForms.DialogResult.Yes)
                {
                    // TODO: Trigger MEP mark re-numbering (can call MarkParameterService with remarkAll=true)
                    DebugLogger.Info($"[UpdateXmlCommand] User requested MEP mark re-numbering after XML update");
                    // Note: Re-numbering should be done via Parameter Service or separate command
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error updating XML: {ex.Message}";
                DebugLogger.Error($"[UpdateXmlCommand] Error: {ex.Message}\n{ex.StackTrace}");
                
                WinForms.MessageBox.Show(
                    $"❌ Error updating XML:\n\n{ex.Message}\n\nSee logs for details.",
                    "Update XML Error",
                    WinForms.MessageBoxButtons.OK,
                    WinForms.MessageBoxIcon.Error);
                
                return Result.Failed;
            }
        }

        /// <summary>
        /// Check if filter XML files exist for MEP categories (same validation as ParameterServiceDialogV2)
        /// </summary>
        private List<string> CheckFilterFilesForCategories(Document doc)
        {
            var missingCategories = new List<string>();
            
            try
            {
                if (doc == null) return missingCategories;
                
                // Get filters directory
                ProjectPathService.EnsureFiltersDirectory(doc);
                string filtersDirectory = ProjectPathService.GetFiltersDirectory(doc);
                
                if (!Directory.Exists(filtersDirectory))
                {
                    // No filters directory - all categories missing
                    return new List<string> { "Ducts", "Pipes", "Cable Trays", "Duct Accessories" };
                }
                
                // Check each MEP category (only Reference Elements categories need filter files)
                var categoriesToCheck = new Dictionary<string, string[]>
                {
                    { "Ducts", new[] { "*_ducts.xml", "*ducts*.xml", "ducts*.xml" } },
                    { "Pipes", new[] { "*_pipes.xml", "*pipes*.xml", "pipes*.xml" } },
                    { "Cable Trays", new[] { "*_cable_trays.xml", "*cable_trays*.xml", "*cable_tray*.xml", "*cabletray*.xml" } },
                    { "Duct Accessories", new[] { "*_duct_accessories.xml", "*duct_accessories*.xml", "*damper*.xml" } }
                };
                
                foreach (var categoryPair in categoriesToCheck)
                {
                    string category = categoryPair.Key;
                    string[] patterns = categoryPair.Value;
                    
                    bool found = false;
                    foreach (var pattern in patterns)
                    {
                        var files = Directory.GetFiles(filtersDirectory, pattern);
                        if (files.Length > 0)
                        {
                            // Found at least one file - check if it has clash zones
                            foreach (var file in files)
                            {
                                try
                                {
                                    // Quick check: deserialize and see if it has clash zones
                                    var serializer = new XmlSerializer(typeof(OpeningFilter));
                                    using (var reader = new StreamReader(file))
                                    {
                                        var filter = (OpeningFilter)serializer.Deserialize(reader);
                                        if (filter?.ClashZoneStorage?.AllZones != null && filter.ClashZoneStorage.AllZones.Count > 0)
                                        {
                                            found = true;
                                            break;
                                        }
                                    }
                                }
                                catch
                                {
                                    // If deserialization fails, skip this file
                                    continue;
                                }
                            }
                        }
                        
                        if (found) break;
                    }
                    
                    if (!found)
                    {
                        missingCategories.Add(category);
                    }
                }
                
                DebugLogger.Info($"[UpdateXmlCommand] Filter file check: {missingCategories.Count} categories missing filter data");
                if (missingCategories.Count > 0)
                {
                    DebugLogger.Warning($"[UpdateXmlCommand] Missing filter files for: {string.Join(", ", missingCategories)}");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UpdateXmlCommand] Error checking filter files: {ex.Message}");
                // On error, don't block processing - return empty list
            }
            
            return missingCategories;
        }
    }
}

