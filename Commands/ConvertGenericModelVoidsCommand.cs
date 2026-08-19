using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// Command to convert existing Generic Model openings into cuttable openings
    /// in linked Architecture and Structure files.
    /// 
    /// WORKFLOW:
    /// 1. Collects all Generic Model openings from MEP document
    /// 2. For each Generic Model opening:
    ///    - Extracts dimensions, location, and orientation
    ///    - Finds linked architecture and structure files
    ///    - Creates cuttable openings (Wall Opening families or Void profiles)
    ///    - Updates Generic Model with reference info
    /// 3. Reports results and statistics
    /// 
    /// ADVANTAGES:
    /// - No need to re-place existing Generic Models
    /// - Preserves all MEP opening data and parameters
    /// - Creates actual cuttable geometry in architecture/structure
    /// - Supports coordinate transforms for misaligned models
    /// 
    /// REQUIREMENTS:
    /// - Generic Model openings must be placed in MEP document
    /// - Linked architecture and structure files must be open
    /// - Wall Opening families should be loaded in linked files (optional fallback: void profiles)
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class ConvertGenericModelVoidsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uiDoc = commandData.Application.ActiveUIDocument;
                Document doc = uiDoc.Document;

                DebugLogger.InitLogFile("ConvertGenericModelVoidsCommand");
                DebugLogger.Log("ConvertGenericModelVoidsCommand: Execute started");

                // Initialize converter
                var converter = new GenericModelVoidConverter(doc);

                // Collect Generic Model openings
                var openings = converter.CollectGenericModelOpenings();
                if (openings.Count == 0)
                {
                    TaskDialog.Show("Info", "No Generic Model openings found in the document.");
                    return Result.Succeeded;
                }

                DebugLogger.Log($"Found {openings.Count} Generic Model openings to convert");

                // Find linked documents
                var linkedDocs = GetLinkedDocuments(doc);
                if (linkedDocs.Count == 0)
                {
                    TaskDialog.Show("Warning", "No linked architecture or structure documents found.\n" +
                                              "Please ensure architecture and structure files are linked.");
                    return Result.Failed;
                }

                DebugLogger.Log($"Found {linkedDocs.Count} linked documents to process");

                // Process each opening and create cuttable versions in linked files
                int successCount = 0;
                int failureCount = 0;
                var resultsByFile = new Dictionary<string, (int success, int failed)>();

                foreach (var linkedDoc in linkedDocs)
                {
                    string docTitle = linkedDoc.Title ?? linkedDoc.PathName ?? "Unknown";
                    var (success, failed) = ProcessOpeningsForLinkedFile(doc, linkedDoc, openings, converter);
                    
                    successCount += success;
                    failureCount += failed;
                    resultsByFile[docTitle] = (success, failed);

                    DebugLogger.Log($"[{docTitle}] Created {success} openings, {failed} failures");
                }

                // Show results
                ShowResultsDialog(resultsByFile, successCount, failureCount);

                DebugLogger.Log("ConvertGenericModelVoidsCommand: Execute completed successfully");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"ConvertGenericModelVoidsCommand: Error - {ex.Message}\n{ex.StackTrace}");
                TaskDialog.Show("Error", $"Error converting Generic Model voids:\n{ex.Message}");
                return Result.Failed;
            }
        }

        /// <summary>
        /// Gets all linked documents from the active document
        /// Filters by architecture and structure models
        /// </summary>
        private List<Document> GetLinkedDocuments(Document mainDoc)
        {
            var linkedDocs = new List<Document>();

            try
            {
                var linkedInstances = new FilteredElementCollector(mainDoc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>()
                    .ToList();

                foreach (var linkInstance in linkedInstances)
                {
                    try
                    {
                        var linkedDoc = linkInstance.GetLinkDocument();
                        if (linkedDoc != null)
                        {
                            var docTitle = linkedDoc.Title ?? "";
                            // Filter for architecture and structure models
                            if (docTitle.Contains("Arch", StringComparison.OrdinalIgnoreCase) ||
                                docTitle.Contains("Structure", StringComparison.OrdinalIgnoreCase) ||
                                docTitle.Contains("Struct", StringComparison.OrdinalIgnoreCase))
                            {
                                linkedDocs.Add(linkedDoc);
                                DebugLogger.Log($"Added linked document: {docTitle}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Log($"Error getting linked document: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Error collecting linked documents: {ex.Message}");
            }

            return linkedDocs;
        }

        /// <summary>
        /// Processes Generic Model openings and creates cuttable versions in a linked file
        /// </summary>
        private (int success, int failed) ProcessOpeningsForLinkedFile(
            Document mepDoc,
            Document linkedDoc,
            List<FamilyInstance> openings,
            GenericModelVoidConverter converter)
        {
            int successCount = 0;
            int failureCount = 0;

            try
            {
                // Calculate coordinate transform
                var transform = converter.CalculateCoordinateTransform(linkedDoc);

                foreach (var opening in openings)
                {
                    try
                    {
                        // Extract opening data
                        var data = converter.ExtractOpeningData(opening);

                        // Create cuttable opening in linked file
                        var resultId = converter.CreateCuttableOpeningInLinkedFile(linkedDoc, data, transform);

                        if (resultId != ElementId.InvalidElementId)
                        {
                            // Update Generic Model with reference info
                            UpdateOpeningReference(opening, linkedDoc, resultId);
                            successCount++;
                        }
                        else
                        {
                            failureCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Log($"Error processing opening {opening.Id}: {ex.Message}");
                        failureCount++;
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Error processing openings for linked file: {ex.Message}");
            }

            return (successCount, failureCount);
        }

        /// <summary>
        /// Updates Generic Model opening with reference to created cuttable opening
        /// </summary>
        private void UpdateOpeningReference(FamilyInstance opening, Document linkedDoc, ElementId resultId)
        {
            try
            {
                using (var tx = new Transaction(opening.Document, "Update Opening Reference"))
                {
                    tx.Start();

                    var linkedFileParam = opening.LookupParameter("Linked File")
                                        ?? opening.LookupParameter("Source File");
                    if (linkedFileParam != null && !linkedFileParam.IsReadOnly && linkedFileParam.StorageType == StorageType.String)
                    {
                        linkedFileParam.Set(linkedDoc.Title ?? "Linked Model");
                    }

                    var referenceIdParam = opening.LookupParameter("Reference Opening ID")
                                         ?? opening.LookupParameter("Linked Element ID");
                    if (referenceIdParam != null && !referenceIdParam.IsReadOnly && referenceIdParam.StorageType == StorageType.String)
                    {
                        referenceIdParam.Set(resultId.ToLong().ToString());
                    }

                    var statusParam = opening.LookupParameter("Opening Status")
                                    ?? opening.LookupParameter("Status");
                    if (statusParam != null && !statusParam.IsReadOnly && statusParam.StorageType == StorageType.String)
                    {
                        statusParam.Set("Converted to Cuttable Opening");
                    }

                    tx.Commit();
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Error updating opening reference: {ex.Message}");
            }
        }

        /// <summary>
        /// Shows results dialog with statistics
        /// </summary>
        private void ShowResultsDialog(Dictionary<string, (int success, int failed)> resultsByFile, int totalSuccess, int totalFailed)
        {
            string message = "Generic Model Void Conversion Results\n" +
                           "=====================================\n\n";

            foreach (var kvp in resultsByFile)
            {
                message += $"{kvp.Key}:\n" +
                         $"  ✓ Created: {kvp.Value.success}\n" +
                         $"  ✗ Failed: {kvp.Value.failed}\n\n";
            }

            message += "=====================================\n" +
                      $"Total Successfully Created: {totalSuccess}\n" +
                      $"Total Failed: {totalFailed}\n";

            if (totalSuccess > 0)
            {
                message += "\n✓ Conversion completed. Generic Model openings are now referenced in linked files.\n" +
                         "  Architecture and structure walls/floors should now cut these openings.";
            }
            else if (totalFailed > 0)
            {
                message += "\n✗ No openings were successfully created. Check logs for details.";
            }

            TaskDialog.Show("Conversion Results", message);
        }
    }
}
