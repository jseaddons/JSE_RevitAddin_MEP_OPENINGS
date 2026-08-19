using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// Phase 2 Service: Handles placement of both Cluster and Individual sleeves from Database.
    /// Supports Bulk and Sequential transaction modes.
    /// </summary>
    public class UnifiedPlacementService
    {
        private readonly BatchClusterPlacementService _clusterPlacementService;
        private readonly IClashZoneRepository _repository;
        private readonly SettingsModel _settings;

        public UnifiedPlacementService(
            BatchClusterPlacementService clusterPlacementService,
            IClashZoneRepository repository,
            SettingsModel settings)
        {
            _clusterPlacementService = clusterPlacementService ?? throw new ArgumentNullException(nameof(clusterPlacementService));
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        }

        public void PlaceAllPending(Document doc, bool useSingleTransaction = true)
        {
            // 1. Place Clusters First (Priority)
            try
            {
                _clusterPlacementService.PlaceFromDatabase(doc, null, useSingleTransaction); // Place all pending
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("unified_placement_errors.log", $"Error placing clusters: {ex.Message}\n");
            }

            // 2. Place Individual Sleeves
            try
            {
                PlacePendingIndividualSleeves(doc, useSingleTransaction);
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("unified_placement_errors.log", $"Error placing individual sleeves: {ex.Message}\n");
            }
        }

                /// <summary>
        /// Places ALL individual sleeves marked as 'Pending', ignoring any clustering flags.
        /// Phase 1b of the Hybrid Workflow.
        /// </summary>
        public void PlaceAllIndividuals(Document doc)
        {
            try
            {
                // Fetch ALL pending zones regardless of cluster status
                // We reuse the logic but disable the filter for clusters
                var categories = _repository.GetDistinctCategories();
                var allPending = new List<ClashZone>();

                foreach (var cat in categories)
                {
                    var zones = _repository.GetClashZonesByCategory(cat);
                    // FILTER: Only pick those that are calculated and pending
                    allPending.AddRange(zones.Where(z =>
                       z.PlacementStatus == "Pending" &&
                       !z.IsResolved && 
                       !z.IsCombinedResolved
                       // We DO NOT filter out 'IsClusterResolved' or 'MarkedForCluster' here 
                       // because in Phase 1b we place EVERYTHING individually first.
                    ));
                }

                if (!allPending.Any()) return;

                using (Transaction t = new Transaction(doc, "Place All Individuals (Phase 1b)"))
                {
                    t.Start();
                    var failureOptions = t.GetFailureHandlingOptions();
                    failureOptions.SetFailuresPreprocessor(new WarningSwallower());
                    t.SetFailureHandlingOptions(failureOptions);

                    int successCount = 0;
                    var placedZones = new List<ClashZone>();
                    foreach (var zone in allPending)
                    {
                        try
                        {
                            if (PlaceSingleSleeve(doc, zone))
                            {
                                successCount++;
                                placedZones.Add(zone);
                            }
                        }
                        catch (Exception ex)
                        {
                            SafeFileLogger.SafeAppendText("placement_errors.log", $"Failed Phase 1b place for {zone.Id}: {ex.Message}");
                        }
                    }
                    t.Commit();
                    
                    if (placedZones.Any())
                    {
                        _repository.SaveSleeveSnapshotsForPlacedSleeves(-1, placedZones);
                    }
                    SafeFileLogger.SafeAppendText("placement.log", $"Phase 1b: Placed {successCount} individual sleeves.");
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("placement_errors.log", $"CRITICAL Phase 1b Error: {ex.Message}");
            }
        }

        public void PlacePendingIndividualSleeves(Document doc, bool useSingleTransaction)
        {
            // Fetch pending zones that are NOT resolved by clusters/combined
            var categories = _repository.GetDistinctCategories();
            var allPending = new List<ClashZone>();

            foreach(var cat in categories)
            {
                  var zones = _repository.GetClashZonesByCategory(cat);
                  allPending.AddRange(zones.Where(z => 
                     z.PlacementStatus == "Pending" && 
                     !z.IsClusterResolved && 
                     !z.IsResolved &&
                     !z.IsCombinedResolved
                  ));
            }

            if (!allPending.Any()) return;

            if (useSingleTransaction)
            {
                using (Transaction t = new Transaction(doc, "Place Individual Sleeves (Batch)"))
                {
                    t.Start();
                    // Disable warnings/failures
                    var failureOptions = t.GetFailureHandlingOptions();
                    failureOptions.SetFailuresPreprocessor(new WarningSwallower());
                    t.SetFailureHandlingOptions(failureOptions);

                    var placedZones = new List<ClashZone>();
                    foreach (var zone in allPending)
                    {
                        if (PlaceSingleSleeve(doc, zone))
                        {
                            placedZones.Add(zone);
                        }
                    }
                    t.Commit();
                    
                    // ✅ CRITICAL FIX: Save snapshots for all placed individual sleeves
                    // This ensures "Self-Healing Retrieval" works when these are later clustered.
                    if (placedZones.Any())
                    {
                        _repository.SaveSleeveSnapshotsForPlacedSleeves(-1, placedZones);
                        SafeFileLogger.SafeAppendText("placement.log", $"[{DateTime.Now:HH:mm:ss}] 📸 Snapshot saved for {placedZones.Count} individual sleeves.\n");
                    }
                }
            }
            else
            {
                var placedZones = new List<ClashZone>();
                foreach (var zone in allPending)
                {
                    using (Transaction t = new Transaction(doc, $"Place Sleeve {zone.Id}"))
                    {
                        t.Start();
                        // Disable warnings
                        var failureOptions = t.GetFailureHandlingOptions();
                        failureOptions.SetFailuresPreprocessor(new WarningSwallower());
                        t.SetFailureHandlingOptions(failureOptions);

                        if (PlaceSingleSleeve(doc, zone))
                        {
                            placedZones.Add(zone);
                            t.Commit();
                        }
                        else
                        {
                            t.RollBack();
                        }
                    }
                }
                
                // ✅ CRITICAL FIX: Save snapshots for all placed individual sleeves (even in sequential mode)
                if (placedZones.Any())
                {
                    _repository.SaveSleeveSnapshotsForPlacedSleeves(-1, placedZones);
                    SafeFileLogger.SafeAppendText("placement.log", $"[{DateTime.Now:HH:mm:ss}] 📸 Snapshot saved for {placedZones.Count} individual sleeves (Sequential).\n");
                }
            }
        }



        private bool PlaceSingleSleeve(Document doc, ClashZone zone)
        {
            try
            {
                // 1. Validation
                if (string.IsNullOrEmpty(zone.CalculatedFamilyName)) return false;

                // 2. Load Family
                FamilySymbol symbol = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(x => x.Name.Equals(zone.CalculatedFamilyName));
                
                if (symbol == null)
                {
                    // Log error: family not found
                    return false;
                }
                if (!symbol.IsActive) symbol.Activate();

                // 3. Create Instance
                // Use structural host or level?
                // NewSleevePlacer uses host if possible.
                Element host = null;
                if (zone.StructuralElementIdValue > 0)
                {
#if REVIT2024_OR_GREATER
                    try { host = doc.GetElement(new ElementId(zone.StructuralElementIdValue)); } catch {}
#else
                    try { host = doc.GetElement(new ElementId((int)zone.StructuralElementIdValue)); } catch {}
#endif
                }

                bool hasWallCenterline = Math.Abs(zone.WallCenterlinePointX) > 1e-6 || Math.Abs(zone.WallCenterlinePointY) > 1e-6 || Math.Abs(zone.WallCenterlinePointZ) > 1e-6;
                XYZ placementPoint = hasWallCenterline 
                    ? new XYZ(zone.WallCenterlinePointX, zone.WallCenterlinePointY, zone.WallCenterlinePointZ)
                    : new XYZ(zone.IntersectionPointX, zone.IntersectionPointY, zone.IntersectionPointZ); 
                
                // DIAGNOSTIC
                if (!DeploymentConfiguration.DeploymentMode && hasWallCenterline)
                {
                    SafeFileLogger.SafeAppendText("placement_debug.log", $"[UnifiedPlacement] Using centerline for zone {zone.Id}: {placementPoint}\n");
                }                
                FamilyInstance instance = null;
                Autodesk.Revit.DB.Structure.StructuralType st = Autodesk.Revit.DB.Structure.StructuralType.NonStructural;
                
                if (host != null && (host is Wall || host is Floor || host is Ceiling))
                {
                     instance = doc.Create.NewFamilyInstance(placementPoint, symbol, host, st);
                }
                else
                {
                     instance = doc.Create.NewFamilyInstance(placementPoint, symbol, st);
                }

                if (instance == null) return false;

                // 4. Set Parameters (Width, Height, Depth taken from Calculated fields)
                SetParam(instance, "V", zone.CalculatedSleeveHeight); 
                SetParam(instance, "H", zone.CalculatedSleeveWidth);
                SetParam(instance, "D", zone.CalculatedSleeveDepth);
                SetParam(instance, "Nominal Diameter", zone.CalculatedSleeveWidth);

                // Rotation
                if (Math.Abs(zone.CalculatedRotation) > 0.001)
                {
                    ElementTransformUtils.RotateElement(doc, instance.Id, Line.CreateBound(placementPoint, placementPoint + XYZ.BasisZ), zone.CalculatedRotation);
                }

                // 5. Update DB Status
                _repository.UpdateSleevePlacement(zone.Id, instance.Id.GetIntegerValue(), 
                    zone.CalculatedSleeveWidth, zone.CalculatedSleeveHeight, 0,
                    placementPoint.X, placementPoint.Y, placementPoint.Z,
                    placementPoint.X, placementPoint.Y, placementPoint.Z,
                    zone.CalculatedRotation, zone.CalculatedFamilyName, 
                    markedForClusterProcess: false); 
                    
                _repository.UpdateCalculatedSleeveData(zone.Id, 
                    zone.CalculatedSleeveWidth, zone.CalculatedSleeveHeight, zone.CalculatedSleeveDepth,
                    zone.CalculatedRotation, zone.CalculatedFamilyName,
                    "Placed", zone.CalculationBatchId);
                    
                // ✅ UPDATE LOCAL OBJECT: Crucial for snapshot saving later
                zone.SleeveInstanceId = instance.Id.GetIntegerValue();
                
                return true;
            }
            catch (Exception ex)
            {
               SafeFileLogger.SafeAppendText("placement_errors.log", $"Failed to place zone {zone.Id}: {ex.Message}\n");
               return false;
            }
        }

        private void SetParam(Element e, string name, double val)
        {
            Parameter p = e.LookupParameter(name);
            if (p != null && !p.IsReadOnly)
            {
                p.Set(val);
            }
        }
    }
    
    // Simple helper if not exists
    public class WarningSwallower : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
        {
            failuresAccessor.DeleteAllWarnings();
            return FailureProcessingResult.Continue;
        }
    }
}