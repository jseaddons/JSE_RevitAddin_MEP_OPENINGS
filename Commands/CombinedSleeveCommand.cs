using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.UI;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Repository;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase3And4.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.BoundingBox;
using JSE_RevitAddin_MEP_OPENINGS.Services.Combined;
using System.Collections.Generic;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CombinedSleeveCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Feature 20: Crash Safe Execution
                var crashSafeExecutor = new CrashSafeExecutor();
                string resultMsg = "";
                var result = crashSafeExecutor.ExecuteWithTimeout(() =>
                {
                    try
                    {
                        // Initialize Document and Services
                        var uidoc = commandData.Application.ActiveUIDocument;
                        var doc = uidoc.Document;

                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info("[CombinedSleeveCommand] Starting Combined Sleeve UI...");

                        // Create Event Handler
                        var handler = new CombinedSleeveRequestHandler();
                        // IMPORTANT: ExternalEvent can fail if not in valid context, but we are in IExternalCommand here.
                        var externalEvent = Autodesk.Revit.UI.ExternalEvent.Create(handler);

                        if (externalEvent == null)
                        {
                            TaskDialog.Show("Error", "ExternalEvent creation failed!");
                            return Result.Failed;
                        }

                        // Composition Root: Instantiate Dependencies
                        
                        // 1. Data Access
                        var dbContext = new SleeveDbContext(doc);
                        var repo = new ClashZoneRepository(dbContext);
                        var combinedRepo = new CombinedClusterRepository(repo);

                        // 2. Core Services
                        var discoveryService = new CombinedClusterDiscoveryService(combinedRepo);
                        var formationService = new JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase3And4.Services.CombinedClusterFormationService(combinedRepo);
                        var paramService = new ParameterAggregatorService();
                        var persistService = new CombinedClusterPersistenceService(repo);

                        // 3. Manual Calculation Components (Adapter + Helpers)
                        var getClashZoneFunc = new Func<int, string, Models.ClashZone>((id, xml) => {
                            var zones = repo.GetClashZonesBySleeveIds(new[] { id });
                            return (zones != null && zones.Count > 0) ? zones[0] : null; 
                        });
                 
                        var rotationService = new ClusterRotationService(getClashZoneFunc);
                        
                        var fallbackFunc = new Func<List<FamilyInstance>, double, (double, double, double, XYZ)>((sleeves, angle) => {
                             if (sleeves == null || sleeves.Count == 0) return (0,0,0, XYZ.Zero);
                             var bbox = sleeves[0].get_BoundingBox(null);
                             if (bbox == null) return (0,0,0, XYZ.Zero);
                             return (bbox.Max.X - bbox.Min.X, bbox.Max.Z - bbox.Min.Z, bbox.Max.Y - bbox.Min.Y, (bbox.Min + bbox.Max)/2);
                        });
                        
                        var bboxCalculator = new RotatedBoundingBoxCalculator(getClashZoneFunc, fallbackFunc);
                        
                        var manualCalculator = new ManualClusterCalculationAdapter(rotationService, bboxCalculator, repo);

                        // 4. Inject into ViewModel
                        // Initialize ViewModel with Dependencies
                        var viewModel = new CombinedSleeveViewModel(
                            uidoc, 
                            externalEvent, 
                            handler,
                            repo,
                            discoveryService, 
                            formationService, 
                            paramService, 
                            persistService,
                            manualCalculator);

                        // Create and Show Window
                        var window = new CombinedSleeveWindow(viewModel);
                        
                        // FIX: Set Owner logic to prevent window hiding behind Revit
                        try 
                        {
                            var handle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
                            new System.Windows.Interop.WindowInteropHelper(window).Owner = handle;

                        }
                        catch (Exception winEx)
                        {
                             if (!DeploymentConfiguration.DeploymentMode)
                                 DebugLogger.Warning($"Could not set window owner: {winEx.Message}");
                        }

                        window.Show();

                        return Result.Succeeded;
                    }
                    catch (Exception ex)
                    {
                        // VITAL: Show error to user since command fails silently otherwise
                        TaskDialog.Show("Error", $"Command Failed: {ex.Message}\n{ex.StackTrace}");
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error($"[CombinedSleeveCommand] Failed to launch UI: {ex.Message}");
                        resultMsg = ex.Message;
                        return Result.Failed;
                    }
                }, "CombinedSleeveCommand");
                
                if (result == Result.Failed && !string.IsNullOrEmpty(resultMsg))
                {
                    message = resultMsg;
                }
                return result;
        }
    }
}
