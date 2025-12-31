using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Entities;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Services.Refresh;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service to update SQLite database from the current Revit model.
    /// Handles manual resizing, moving, and creation of sleeves.
    /// </summary>
    public class UpdateDbService
    {
        private readonly Action<string> _logger;

        public UpdateDbService(Action<string> logger = null)
        {
            _logger = logger ?? (_ => { });
        }

        public int UpdateDbFromRevit(Document doc)
        {
            int updatedCount = 0;
            int newCount = 0;

            try
            {
                using (var context = new SleeveDbContext(doc, _logger))
                {
                    var repo = new ClashZoneRepository(context, _logger);
                    
                    // 1. Collect all Sleeve Instances (FamilyInstances)
                    // Assumption: Sleeves have specific family names or parameters.
                    // Better to filter by parameter existence "ClashZoneGuid" OR family name convention.
                    // Let's grab all FamilyInstances and check for typical parameters.
                    var sleeveCollector = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilyInstance))
                        .WhereElementIsNotElementType()
                        .ToElements()
                        .Cast<FamilyInstance>()
                        .Where(IsSleeveFamily)
                        .ToList();

                    var zonesToUpdate = new List<ClashZone>();
                    var zonesToInsert = new List<ClashZone>();

                    foreach (var sleeve in sleeveCollector)
                    {
                        try
                        {
                            // Get Geometry / location
                            var bbox = sleeve.get_BoundingBox(null);
                            if (bbox == null) continue;

                            var locationPoint = (sleeve.Location as LocationPoint)?.Point;
                            if (locationPoint == null) locationPoint = (bbox.Min + bbox.Max) * 0.5;

                            // Transform to World Coords if necessary? 
                            // LocationPoint is in Model Coords (Active). Database stores World?
                            // Phase 2 stores "PlacementX/Y/Z" (World) and "PlacementActiveX/Y/Z" (Active).
                            // If base point is not 0,0,0, they differ.
                            // We should use ProjectLocation for transform.
                            // For now, assuming Active Document Coords is what we get from API.
                            
                            // Dimensions
                            double width = 0, height = 0, diameter = 0, length = 0;
                            var wParam = sleeve.LookupParameter("Width");
                            var hParam = sleeve.LookupParameter("Height");
                            var diaParam = sleeve.LookupParameter("Diameter");
                            var lenParam = sleeve.LookupParameter("Depth") ?? sleeve.LookupParameter("Length"); // Depth often used for length/thickness

                            if (wParam != null) width = wParam.AsDouble();
                            if (hParam != null) height = hParam.AsDouble();
                            if (diaParam != null) diameter = diaParam.AsDouble();
                            if (lenParam != null) length = lenParam.AsDouble();

                            // Rotation
                            double rotation = (sleeve.Location as LocationPoint)?.Rotation ?? 0;

                            // 2. Check for GUID
                            string guidStr = sleeve.LookupParameter("ClashZoneGuid")?.AsString();
                            Guid zoneGuid;

                            if (Guid.TryParse(guidStr, out zoneGuid))
                            {
                                // EXISTING SLEEVE - UPDATE
                                var existingZone = repo.GetClashZonesByGuids(new[] { zoneGuid }).FirstOrDefault();
                                if (existingZone != null)
                                {
                                    // Update properties
                                    existingZone.SleeveWidth = width;
                                    existingZone.SleeveHeight = height;
                                    existingZone.SleeveDiameter = diameter;
                                    existingZone.SleeveLength = length;
                                    existingZone.PlacementActiveX = locationPoint.X;
                                    existingZone.PlacementActiveY = locationPoint.Y;
                                    existingZone.PlacementActiveZ = locationPoint.Z;
                                    existingZone.PlacementX = locationPoint.X; // Assuming no shared coords transform for now or same
                                    existingZone.PlacementY = locationPoint.Y;
                                    existingZone.PlacementZ = locationPoint.Z;
                                    existingZone.RotationAngle = rotation;
                                    existingZone.LastUpdated = DateTime.Now;
                                    // existingZone.IsManual = true; // Mark as manually touched if we have such flag?
                                    // existingZone.IsResolved = true; // If it exists, it is resolved? Or keep current status?
                                    // Usually Update DB implies keeping it in sync.

                                    zonesToUpdate.Add(existingZone);
                                    updatedCount++;
                                }
                                else
                                {
                                    // GUID exists on element but not in DB? (Deleted from DB?). Re-insert?
                                    // Treat as NEW manual sleeve.
                                    var newZone = CreateZoneFromSleeve(doc, sleeve, zoneGuid);
                                    if (newZone != null) 
                                    {
                                        zonesToInsert.Add(newZone);
                                        newCount++;
                                    }
                                }
                            }
                            else
                            {
                                // NEW MANUAL SLEEVE (No GUID)
                                // Generate GUID
                                zoneGuid = Guid.NewGuid();
                                
                                // Write GUID to Element Parameter (crucial for next time)
                                var pGuid = sleeve.LookupParameter("ClashZoneGuid");
                                if (pGuid != null && !pGuid.IsReadOnly)
                                {
                                    pGuid.Set(zoneGuid.ToString());
                                }

                                var newZone = CreateZoneFromSleeve(doc, sleeve, zoneGuid);
                                if (newZone != null)
                                {
                                    zonesToInsert.Add(newZone);
                                    newCount++;
                                }
                            }
                        }
                        catch (Exception ex) 
                        {
                            _logger($"Error processing sleeve {sleeve.Id}: {ex.Message}");
                        }
                    }

                    // Batch Commits
                    if (zonesToUpdate.Count > 0)
                    {
                        repo.InsertOrUpdateClashZonesBulk(zonesToUpdate, "ManualUpdate");
                    }
                    if (zonesToInsert.Count > 0)
                    {
                        repo.InsertOrUpdateClashZonesBulk(zonesToInsert, "ManualUpdate");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger($"UpdateDB Failed: {ex.Message}");
                throw;
            }

            return updatedCount + newCount;
        }

        private bool IsSleeveFamily(FamilyInstance fi)
        {
            // Naive check: Parameter "ClashZoneGuid" exists? Or naming convention?
            // Checking parameter is safest.
            return fi.LookupParameter("ClashZoneGuid") != null;
        }

        private ClashZone CreateZoneFromSleeve(Document doc, FamilyInstance sleeve, Guid guid)
        {
            // Identify intersecting MEP to fill required fields
            // Use BoundingBox intersection
            var bbox = sleeve.get_BoundingBox(null);
            if (bbox == null) return null;

            var outline = new Outline(bbox.Min, bbox.Max);
            var filter = new BoundingBoxIntersectsFilter(outline);
            
            // Categories: Ducts, Pipes, Cable Trays, Conduits
            var categories = new List<BuiltInCategory> { 
                BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_PipeCurves, 
                BuiltInCategory.OST_CableTray, BuiltInCategory.OST_Conduit 
            };
            var catFilter = new ElementMulticategoryFilter(categories);
            var finalFilter = new LogicalAndFilter(filter, catFilter);

            var mepElement = new FilteredElementCollector(doc)
                .WherePasses(finalFilter)
                .WhereElementIsNotElementType()
                .FirstOrDefault(); // Take first intersection

            // if (mepElement == null) return null; // Cannot create valid zone without MEP? 
            // Actually, for manual sleeves we might want to allow it even without distinct MEP (Placeholder?).
            // But DB schema constraints might fail. Let's assume user placed it ON a pipe.

            int mepId = mepElement?.Id.IntegerValue ?? -1;
            string mepCat = mepElement?.Category?.Name ?? "Manual";

            // Dimensions
            double width = sleeve.LookupParameter("Width")?.AsDouble() ?? 0;
            double height = sleeve.LookupParameter("Height")?.AsDouble() ?? 0;
            double diameter = sleeve.LookupParameter("Diameter")?.AsDouble() ?? 0;
            double length = sleeve.LookupParameter("Depth")?.AsDouble() ?? 0;
            
            var loc = (sleeve.Location as LocationPoint)?.Point ?? (bbox.Min + bbox.Max)*0.5;

            return new ClashZone
            {
                ClashZoneGuid = guid,
                MepElementIdValue = mepId,
                MepElementCategory = mepCat,
                HostElementIdValue = -1, // Manual - unknown host? Or search hosts too? Expensive.
                HostElementCategory = "Manual",
                CenterX = loc.X,
                CenterY = loc.Y,
                CenterZ = loc.Z,
                SleeveWidth = width,
                SleeveHeight = height,
                SleeveDiameter = diameter,
                SleeveLength = length,
                PlacementActiveX = loc.X,
                PlacementActiveY = loc.Y,
                PlacementActiveZ = loc.Z,
                PlacementX = loc.X,
                PlacementY = loc.Y,
                PlacementZ = loc.Z,
                RotationAngle = (sleeve.Location as LocationPoint)?.Rotation ?? 0,
                IsResolved = true, // It is placed, so resolved.
                SleeveInstanceId = sleeve.Id.IntegerValue,
                IsManual = true, // Assuming column exists or we just track it
                CreatedDate = DateTime.Now,
                LastUpdated = DateTime.Now
            };
        }
    }
}
