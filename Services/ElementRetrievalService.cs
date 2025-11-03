using System;
using System.Linq;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Centralized service for retrieving elements from documents and linked documents.
    /// This eliminates code duplication across RefreshService, ThreePointValidator, and other services.
    /// </summary>
    public static class ElementRetrievalService
    {
        /// <summary>
        /// Gets an element from the document or linked documents.
        /// Searches the host document first, then all linked documents.
        /// </summary>
        /// <param name="document">The host document to search</param>
        /// <param name="elementId">The element ID to find</param>
        /// <param name="enableLogging">If true, logs search progress (default: false)</param>
        /// <returns>The element if found, null otherwise</returns>
        public static Element GetElementFromDocumentOrLinked(Document document, ElementId elementId, bool enableLogging = false)
        {
            if (elementId == null || elementId == ElementId.InvalidElementId)
            {
                return null;
            }

            if (document == null)
            {
                if (enableLogging && !DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning("[ElementRetrievalService] Document is null");
                return null;
            }

            try
            {
                if (enableLogging && !DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[ElementRetrievalService] Looking for element {elementId} in document '{document.Title}'");

                // First try host document
                var element = document.GetElement(elementId);
                if (element != null)
                {
                    if (enableLogging && !DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[ElementRetrievalService] Found element {elementId} in host document - Category: {element.Category?.Name ?? "Unknown"}");
                    return element;
                }

                if (enableLogging && !DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[ElementRetrievalService] Element {elementId} not found in host document, searching linked documents...");

                // If not found, search linked documents
                var linkInstances = new FilteredElementCollector(document)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>();

                if (enableLogging && !DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[ElementRetrievalService] Found {linkInstances.Count()} linked documents to search");

                foreach (var link in linkInstances)
                {
                    var linkDoc = link.GetLinkDocument();
                    if (linkDoc != null)
                    {
                        if (enableLogging && !DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[ElementRetrievalService] Searching in linked document: {linkDoc.Title}");
                        
                        element = linkDoc.GetElement(elementId);
                        if (element != null)
                        {
                            if (enableLogging && !DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[ElementRetrievalService] Found element {elementId} in linked document '{linkDoc.Title}' - Category: {element.Category?.Name ?? "Unknown"}");
                            return element;
                        }
                    }
                }

                if (enableLogging && !DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[ElementRetrievalService] Element {elementId} not found in host document or any linked documents");
                
                return null;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[ElementRetrievalService] Error getting element {elementId}: {ex.Message}");
                return null;
            }
        }
    }
}

