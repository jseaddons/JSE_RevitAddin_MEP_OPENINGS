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
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.BoundingBox;
using JSE_RevitAddin_MEP_OPENINGS.Services.Combined;
using JSE_RevitAddin_MEP_OPENINGS.Services.Geometry;
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

                        // Ensure profile/settings context is set to the active project folder
                        ApplicationProfileService.Instance.UpdateForCurrentDocument(doc);

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
                        var combinedSleeveRepo = new CombinedSleeveRepository(dbContext);
                        var combinedRepo = new CombinedClusterRepository(repo);

                        // 2. Core Services - NEW Agent A/B Architecture
                        var discoveryService = new CombinedClusterDiscoveryService(combinedRepo);
                        var proximityService = new CrossCategoryProximityService(); // NEW: Agent B
                        var cornerService = new SleeveCornerCalculationService(); // For corner calculation
                        var placementService = new CombinedSleevePlacementService(doc, combinedSleeveRepo, proximityService, cornerService); // NEW: Agent A

                        // 3. Manual Calculation Components (Adapter + Helpers)
                        var getClashZoneFunc = new Func<long, string, Models.ClashZone>((id, xml) => {
                            var zones = repo.GetClashZonesBySleeveIds(new[] { id });
                            return (zones != null && zones.Count > 0) ? zones[0] : null; 
                        });
                 
                        var rotationService = new ClusterRotationService(getClashZoneFunc);
                        
                        var fallbackFunc = new Func<List<FamilyInstance>, double, (double, double, double, XYZ)>((sleeves, angle) => {
                             if (sleeves == null || sleeves.Count == 0) return (0,0,0, XYZ.Zero);
                             
                             // DEBUG TRACE FOR USER
                             if (!DeploymentConfiguration.DeploymentMode) DebugLogger.Info($"[CombinedSleeveCommand] Fallback Func called with {sleeves.Count} sleeves, Angle={angle:F4}rad");

                             // If angle is effectively zero, use standard bbox
                             if (Math.Abs(angle) < 1e-6)
                             {
                                 var bbox = sleeves[0].get_BoundingBox(null);
                                 if (bbox == null) return (0,0,0, XYZ.Zero);
                                 
                                 // Union all bboxes
                                 double minX = bbox.Min.X, minY = bbox.Min.Y, minZ = bbox.Min.Z;
                                 double maxX = bbox.Max.X, maxY = bbox.Max.Y, maxZ = bbox.Max.Z;
                                 
                                 for(int i=1; i<sleeves.Count; i++)
                                 {
                                     var b = sleeves[i].get_BoundingBox(null);
                                     if(b != null)
                                     {
                                        minX = Math.Min(minX, b.Min.X); minY = Math.Min(minY, b.Min.Y); minZ = Math.Min(minZ, b.Min.Z);
                                        maxX = Math.Max(maxX, b.Max.X); maxY = Math.Max(maxY, b.Max.Y); maxZ = Math.Max(maxZ, b.Max.Z);
                                     }
                                 }
                                 return (maxX - minX, maxZ - minZ, maxY - minY, new XYZ((minX+maxX)/2, (minY+maxY)/2, (minZ+maxZ)/2));
                             }
                             
                             // For rotated clusters, we must rotate corners to local space first
                             double rMinX = double.MaxValue, rMinY = double.MaxValue, rMinZ = double.MaxValue;
                             double rMaxX = double.MinValue, rMaxY = double.MinValue, rMaxZ = double.MinValue;
                             
                             // Pre-calc rotation terms (inverse rotation to allow alignment to axes)
                             double cos = Math.Cos(-angle);
                             double sin = Math.Sin(-angle);
                             
                             foreach(var s in sleeves)
                             {
                                 var bb = s.get_BoundingBox(null);
                                 if (bb == null) continue;
                                 
                                 // Get 8 corners of Axis-Aligned BBox (approximation of element geometry)
                                 // Note: For better accuracy we should use geometry extraction, but bbox corners is a safe fallback
                                 var pts = new List<XYZ> {
                                     bb.Min, new XYZ(bb.Max.X, bb.Min.Y, bb.Min.Z), new XYZ(bb.Min.X, bb.Max.Y, bb.Min.Z), new XYZ(bb.Max.X, bb.Max.Y, bb.Min.Z),
                                     new XYZ(bb.Min.X, bb.Min.Y, bb.Max.Z), new XYZ(bb.Max.X, bb.Min.Y, bb.Max.Z), new XYZ(bb.Min.X, bb.Max.Y, bb.Max.Z), bb.Max
                                 };
                                 
                                 foreach(var p in pts)
                                 {
                                     // Rotate point to local space
                                     double rx = p.X * cos - p.Y * sin;
                                     double ry = p.X * sin + p.Y * cos;
                                     
                                     rMinX = Math.Min(rMinX, rx); rMaxX = Math.Max(rMaxX, rx);
                                     rMinY = Math.Min(rMinY, ry); rMaxY = Math.Max(rMaxY, ry);
                                     rMinZ = Math.Min(rMinZ, p.Z); rMaxZ = Math.Max(rMaxZ, p.Z);
                                 }
                             }
                             
                             double w = rMaxX - rMinX;
                             double h = rMaxZ - rMinZ; // Z is Height
                             double d = rMaxY - rMinY; // Y is Depth (Thickness)
                             
                             // Calculate center in Local Space
                             double cx = (rMinX + rMaxX) / 2.0;
                             double cy = (rMinY + rMaxY) / 2.0;
                             double cz = (rMinZ + rMaxZ) / 2.0;
                             
                             // Rotate center back to World Space
                             double cosBack = Math.Cos(angle);
                             double sinBack = Math.Sin(angle);
                             double wx = cx * cosBack - cy * sinBack;
                             double wy = cx * sinBack + cy * cosBack;
                             
                             return (w, h, d, new XYZ(wx, wy, cz));
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
                            proximityService, 
                            placementService,
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
