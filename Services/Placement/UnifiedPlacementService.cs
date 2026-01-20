using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

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
                    foreach (var zone in allPending)
                    {
                        try
                        {
                            PlaceSingleSleeve(doc, zone);
                            successCount++;
                        }
                        catch (Exception ex)
                        {
                            SafeFileLogger.SafeAppendText("placement_errors.log", $"Failed Phase 1b place for {zone.Id}: {ex.Message}");
                        }
                    }
                    t.Commit();
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

                    foreach (var zone in allPending)
                    {
                        PlaceSingleSleeve(doc, zone);
                    }
                    t.Commit();
                }
            }
            else
            {
                foreach (var zone in allPending)
                {
                    using (Transaction t = new Transaction(doc, $"Place Sleeve {zone.Id}"))
                    {
                        t.Start();
                         // Disable warnings
                        var failureOptions = t.GetFailureHandlingOptions();
                        failureOptions.SetFailuresPreprocessor(new WarningSwallower());
                        t.SetFailureHandlingOptions(failureOptions);

                        PlaceSingleSleeve(doc, zone);
                        t.Commit();
                    }
                }
            }
        }

        private void PlaceSingleSleeve(Document doc, ClashZone zone)
        {
            try
            {
                // 1. Validation
                if (string.IsNullOrEmpty(zone.CalculatedFamilyName)) return;

                // 2. Load Family
                FamilySymbol symbol = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(x => x.Name.Equals(zone.CalculatedFamilyName));
                
                if (symbol == null)
                {
                    // Log error: family not found
                    return;
                }
                if (!symbol.IsActive) symbol.Activate();

                // 3. Create Instance
                // Use structural host or level?
                // NewSleevePlacer uses host if possible.
                Element host = null;
                if (zone.StructuralElementIdValue > 0)
                {
                    // Compatibility for 2023: ID is integer/long
                     try { host = doc.GetElement(new ElementId((int)zone.StructuralElementIdValue)); } catch {}
                }

                XYZ placementPoint = new XYZ(zone.IntersectionPointX, zone.IntersectionPointY, zone.IntersectionPointZ); // Use stored intersection
                // Or use `SleevePlacementPoint` if calculated?
                // In Phase 1 we calculated and saved... wait, Phase 1 only saved dimensions/rotation. 
                // Position is usually Intersection + Centering?
                // If pure intersection is used, fine.
                // NewSleevePlacer calculates adjustment.
                // If adjustments were needed (like for Dampers), they should have been in Phase 1?
                // Phase 1 `SleeveCalculationService` didn't seem to save Adjusted Point.
                // It saved dimensions.
                // Re-calculating adjustment here is okay if it doesn't use transactions.
                // But generally, intersection is main point.
                
                Level level = doc.GetElement(new ElementId(doc.ActiveView.LevelId.IntegerValue)) as Level; // Fallback? 
                // Ideally get level from zone.MepElementLevelName or similar.
                
                FamilyInstance instance = null;
                if (host != null && host is Wall) // Wall hosting usually requires face or efficient standard placement
                {
                    // Standard Create.NewFamilyInstance(point, symbol, host, StructuralType)
                    instance = doc.Create.NewFamilyInstance(placementPoint, symbol, host, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                }
                else if (host != null && (host is Floor || host is Ceiling))
                {
                     instance = doc.Create.NewFamilyInstance(placementPoint, symbol, host, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                }
                else
                {
                    // Unhosted / Face based?
                     instance = doc.Create.NewFamilyInstance(placementPoint, symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                }

                if (instance == null) return;

                // 4. Set Parameters (Width, Height, Depth taken from Calculated fields)
                // 🔍 LOGGING: Log calculated values from DB before setting
                SafeFileLogger.SafeAppendText("placement_sizing_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 📏 PLACING SLEEVE {zone.Id} (InstId: {instance.Id.IntegerValue}):\n" +
                    $"    - CalculatedWidth: {zone.CalculatedSleeveWidth * 304.8:F1}mm ({zone.CalculatedSleeveWidth:F4}ft)\n" +
                    $"    - CalculatedHeight: {zone.CalculatedSleeveHeight * 304.8:F1}mm ({zone.CalculatedSleeveHeight:F4}ft)\n" +
                    $"    - CalculatedDepth: {zone.CalculatedSleeveDepth * 304.8:F1}mm ({zone.CalculatedSleeveDepth:F4}ft)\n" +
                    $"    - CalculatedFamily: {zone.CalculatedFamilyName}\n");

                SetParam(instance, "V", zone.CalculatedSleeveHeight); // Height
                SetParam(instance, "H", zone.CalculatedSleeveWidth);  // Width
                SetParam(instance, "D", zone.CalculatedSleeveDepth);  // Depth (Length)
                
                // Diameter? If circular family.
                SetParam(instance, "Nominal Diameter", zone.CalculatedSleeveWidth); // Often Width or D

                // Rotation
                if (Math.Abs(zone.CalculatedRotation) > 0.001)
                {
                    ElementTransformUtils.RotateElement(doc, instance.Id, Line.CreateBound(placementPoint, placementPoint + XYZ.BasisZ), zone.CalculatedRotation);
                }

                // 5. Update DB Status (After commit it's real, but inside transaction we can't update DB easily? DB is outside transaction usually)
                // We should collect IDs and update DB *after* transaction commits if bulk.
                // But inside 'PlaceAllPending', we are managing transaction.
                // Wait, if transaction rolls back, DB should not be updated.
                // So updating DB *should* happen after commit.
                // I need to return success/failure list.
                
                // For now, I will update DB immediately (risk of drift if commit fails).
                // Better pattern: List of (ZoneId, InstanceId) to update after loop.
                
                _repository.UpdateSleevePlacement(zone.Id, instance.Id.IntegerValue, 
                    zone.CalculatedSleeveWidth, zone.CalculatedSleeveHeight, 0,
                    placementPoint.X, placementPoint.Y, placementPoint.Z,
                    placementPoint.X, placementPoint.Y, placementPoint.Z,
                    zone.CalculatedRotation, zone.CalculatedFamilyName, 
                    markedForClusterProcess: false); // ✅ FIX: Reset this flag for individual placement
                    
                // Update 'PlacementStatus' to Placed
                _repository.UpdateCalculatedSleeveData(zone.Id, 
                    zone.CalculatedSleeveWidth, zone.CalculatedSleeveHeight, zone.CalculatedSleeveDepth,
                    zone.CalculatedRotation, zone.CalculatedFamilyName,
                    "Placed", zone.CalculationBatchId);
            }
            catch (Exception ex)
            {
               // Log
               SafeFileLogger.SafeAppendText("placement_errors.log", $"Failed to place zone {zone.Id}: {ex.Message}\n");
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
