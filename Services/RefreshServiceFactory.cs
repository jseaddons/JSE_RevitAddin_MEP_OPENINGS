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
    /// Factory for creating RefreshService instances
    /// ✅ PHASE 2: Legacy RefreshService removed - only refactored service is used
    /// </summary>
    public static class RefreshServiceFactory
    {
        /// <summary>
        /// Creates the refactored RefreshService implementation
        /// </summary>
        /// <param name="document">Revit document</param>
        /// <param name="uiDocument">Revit UI document</param>
        /// <param name="appProfileService">Application profile service</param>
        /// <returns>IRefreshService implementation (refactored)</returns>
        public static IRefreshService Create(
            Document document,
            UIDocument uiDocument,
            ApplicationProfileService appProfileService)
        {
            if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info("[RefreshServiceFactory] Creating REFACTORED RefreshService");
            
            return new RefreshServiceRefactoredWrapper(
                new RefreshServiceRefactored(document, uiDocument, appProfileService));
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
            try
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info("[RefreshServiceRefactoredWrapper] Starting refactored refresh service...");
                
                // ✅ CRITICAL: Verify existing sleeves and reset flags for deleted sleeves
                // This ensures that sleeves deleted in Revit are re-detected and can be re-placed
                try
                {
                    using (var dbContext = new Data.SleeveDbContext(_service.GetDocument()))
                    {
                        var repo = new Data.Repositories.ClashZoneRepository(dbContext, msg => 
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info(msg);
                        });
                        
                        int resetCount = repo.VerifyExistingSleevesAndResetFlags(
                            _service.GetDocument(), 
                            new List<string>(), // All filters
                            selectedMepCategories ?? new List<string>());
                            
                        if (resetCount > 0)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[RefreshServiceRefactoredWrapper] 🔄 Reset flags for {resetCount} deleted sleeves");
                        }
                    }
                }
                catch (Exception verifyEx)
                {
                    // Log but don't fail - refresh can continue even if flag reset fails
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[RefreshServiceRefactoredWrapper] ⚠️ Sleeve verification failed (continuing): {verifyEx.Message}");
                }
                
                // Refactored service returns Result, but we need void for compatibility
                // Convert Result to void (errors are logged internally)
                var result = _service.ExecuteRefresh(
                    selectedFilterItems,
                    selectedMepCategories,
                    selectedReferenceFiles,
                    selectedHostFiles,
                    clearanceSettings);

                if (result == Autodesk.Revit.UI.Result.Succeeded)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info("[RefreshServiceRefactoredWrapper] ✅ Refresh completed successfully");
                }
                else if (result == Autodesk.Revit.UI.Result.Cancelled)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning("[RefreshServiceRefactoredWrapper] ⚠️ Refresh was cancelled (missing selections)");
                }
                else
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[RefreshServiceRefactoredWrapper] ❌ Refresh failed with result: {result}");
                    
                    // Show error to user
                    System.Windows.Forms.MessageBox.Show(
                        $"Refresh failed. Check logs for details.\nResult: {result}",
                        "Refresh Error",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                // ✅ CRITICAL: Catch and log ALL exceptions
                var errorMsg = $"Refactored refresh service threw exception: {ex.Message}";
                DebugLogger.Error($"[RefreshServiceRefactoredWrapper] ❌ {errorMsg}\n{ex.StackTrace}");
                
                // Show error to user
                System.Windows.Forms.MessageBox.Show(
                    $"{errorMsg}\n\nCheck debug logs for full details.",
                    "Refresh Exception",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
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

