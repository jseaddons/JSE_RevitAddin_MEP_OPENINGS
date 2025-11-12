using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
// Note: RefreshServiceRefactored is in JSE_RevitAddin_MEP_OPENINGS.Services namespace (same as this file)
// Refresh helper classes are in JSE_RevitAddin_MEP_OPENINGS.Services.Refresh namespace

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Factory for creating RefreshService instances with feature flag support
    /// Allows switching between legacy and refactored refresh service implementations
    /// </summary>
    public static class RefreshServiceFactory
    {
        /// <summary>
        /// Creates the appropriate RefreshService implementation based on feature flag
        /// </summary>
        /// <param name="document">Revit document</param>
        /// <param name="uiDocument">Revit UI document</param>
        /// <param name="appProfileService">Application profile service</param>
        /// <returns>IRefreshService implementation (legacy or refactored)</returns>
        public static IRefreshService Create(
            Document document,
            UIDocument uiDocument,
            ApplicationProfileService appProfileService)
        {
            // Check feature flag from settings
            var settings = appProfileService?.GetCurrentSettings();
            bool useRefactoredRefresh = settings?.UseRefactoredRefreshService ?? false;

            if (useRefactoredRefresh)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info("[RefreshServiceFactory] Using REFACTORED RefreshService");
                
                return new RefreshServiceRefactoredWrapper(
                    new RefreshServiceRefactored(document, uiDocument, appProfileService));
            }
            else
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info("[RefreshServiceFactory] Using LEGACY RefreshService");
                
                return new RefreshServiceLegacyWrapper(
                    new RefreshService(document, uiDocument, appProfileService));
            }
        }
    }

    /// <summary>
    /// Common interface for refresh service implementations
    /// </summary>
    public interface IRefreshService
    {
        void SetUIReferences(
            System.Windows.Forms.Label statusLabel,
            System.Windows.Forms.ProgressBar progressBar,
            System.Windows.Forms.Button refreshButton);

        void ExecuteRefresh(
            List<string> selectedFilterItems,
            List<string> selectedMepCategories,
            List<string> selectedReferenceFiles,
            List<string> selectedHostFiles,
            Dictionary<string, double> clearanceSettings);

        void LoadExistingClashZoneData();
    }

    /// <summary>
    /// Wrapper for legacy RefreshService to implement IRefreshService interface
    /// </summary>
    internal class RefreshServiceLegacyWrapper : IRefreshService
    {
        private readonly RefreshService _service;

        public RefreshServiceLegacyWrapper(RefreshService service)
        {
            _service = service ?? throw new ArgumentNullException(nameof(service));
        }

        public void SetUIReferences(
            System.Windows.Forms.Label statusLabel,
            System.Windows.Forms.ProgressBar progressBar,
            System.Windows.Forms.Button refreshButton)
        {
            _service.SetUIReferences(statusLabel, progressBar, refreshButton);
        }

        public void ExecuteRefresh(
            List<string> selectedFilterItems,
            List<string> selectedMepCategories,
            List<string> selectedReferenceFiles,
            List<string> selectedHostFiles,
            Dictionary<string, double> clearanceSettings)
        {
            _service.ExecuteRefresh(
                selectedFilterItems,
                selectedMepCategories,
                selectedReferenceFiles,
                selectedHostFiles,
                clearanceSettings);
        }

        public void LoadExistingClashZoneData()
        {
            _service.LoadExistingClashZoneData();
        }
    }

    /// <summary>
    /// Wrapper for refactored RefreshServiceRefactored to implement IRefreshService interface
    /// </summary>
    internal class RefreshServiceRefactoredWrapper : IRefreshService
    {
        private readonly RefreshServiceRefactored _service;

        public RefreshServiceRefactoredWrapper(RefreshServiceRefactored service)
        {
            _service = service ?? throw new ArgumentNullException(nameof(service));
        }

        public void SetUIReferences(
            System.Windows.Forms.Label statusLabel,
            System.Windows.Forms.ProgressBar progressBar,
            System.Windows.Forms.Button refreshButton)
        {
            _service.SetUIReferences(statusLabel, progressBar, refreshButton);
        }

        public void ExecuteRefresh(
            List<string> selectedFilterItems,
            List<string> selectedMepCategories,
            List<string> selectedReferenceFiles,
            List<string> selectedHostFiles,
            Dictionary<string, double> clearanceSettings)
        {
            // Refactored service returns Result, but we need void for compatibility
            // Convert Result to void (errors are logged internally)
            var result = _service.ExecuteRefresh(
                selectedFilterItems,
                selectedMepCategories,
                selectedReferenceFiles,
                selectedHostFiles,
                clearanceSettings);

            if (result != Autodesk.Revit.UI.Result.Succeeded && !DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Warning($"[RefreshServiceRefactoredWrapper] Refresh completed with result: {result}");
            }
        }

        public void LoadExistingClashZoneData()
        {
            // Refactored service loads existing clash zones internally during ExecuteRefresh
            // This method is a no-op for compatibility with legacy interface
            if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info("[RefreshServiceRefactoredWrapper] LoadExistingClashZoneData called - refactored service loads data internally");
        }
    }
}

