using System;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Structure;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Debug service for placing fire damper/duct accessory sleeves.
    /// Computes accurate center from accessory geometry and projects onto wall.
    /// </summary>
    public class FireDamperSleevePlacerDebug
    {
        private readonly Document _doc;

        public FireDamperSleevePlacerDebug(Document doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        }        /// <summary>
                 /// Places a sleeve for a fire damper accessory with debug logging.
                 /// </summary>
                 /// <returns>True if sleeve was placed successfully, false if skipped or failed</returns>
        public bool PlaceFireDamperSleeveDebug(
            FamilyInstance accessory,
            FamilySymbol sleeveSymbol,
            Wall linkedWall)
        {
            DamperLogger.Log("FireDamperSleevePlacerDebug: Starting placement.");

            try
            {
                // Get bounding boxes for the damper and the wall
                var damperBBox = accessory.get_BoundingBox(null);
                var wallBBox = linkedWall.get_BoundingBox(null);

                if (damperBBox == null || wallBBox == null)
                {
                    DamperLogger.Log("Bounding box unavailable for damper or wall. Skipping placement.");
                    return false;
                }

                // Check if bounding boxes intersect
                if (!BoundingBoxesIntersect(damperBBox, wallBBox))
                {
                    DamperLogger.Log("Damper bounding box does not intersect wall bounding box. Skipping placement.");
                    return false;
                }

                // Retrieve damper dimensions using correct parameter names
                var widthParam = accessory.LookupParameter("Damper Width");
                var heightParam = accessory.LookupParameter("Damper Height");

                if (widthParam == null || heightParam == null)
                {
                    DamperLogger.Log("Damper parameters 'Damper Width' or 'Damper Height' not found. Skipping placement.");
                    return false;
                }

                double damperWidth = widthParam.AsDouble();
                double damperHeight = heightParam.AsDouble();

                DamperLogger.Log($"Retrieved damper dimensions: Width={damperWidth}, Height={damperHeight}");

                // Log damper dimensions in millimeters
                DamperLogger.Log($"Damper dimensions (mm): Width={UnitUtils.ConvertFromInternalUnits(damperWidth, UnitTypeId.Millimeters)}, Height={UnitUtils.ConvertFromInternalUnits(damperHeight, UnitTypeId.Millimeters)}");

                // Get symbol type name and log
                string familyTypeName = accessory.Symbol.Name;
                DamperLogger.Log($"Damper family type name: {familyTypeName}");

                // Apply clearance logic based on damper type
                double clearance = UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters);
                double rightClearance = 0.0;  // declared here for later offset calculation
                double sleeveWidth;
                double sleeveHeight = damperHeight + 2 * clearance;

                bool isMSFD = familyTypeName.IndexOf("MSFD", StringComparison.OrdinalIgnoreCase) >= 0;
                if (isMSFD)
                {
                    rightClearance = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
                    sleeveWidth = damperWidth + clearance + rightClearance;
                    DamperLogger.Log("MSFD detected by type name; applying 50mm left and 100mm right clearance.");
                    DamperLogger.Log($"MSFD sleeve dims (internal): Width={sleeveWidth}, Height={sleeveHeight}");
                    DamperLogger.Log($"MSFD sleeve dims (mm): Width={UnitUtils.ConvertFromInternalUnits(sleeveWidth, UnitTypeId.Millimeters)}, Height={UnitUtils.ConvertFromInternalUnits(sleeveHeight, UnitTypeId.Millimeters)}");
                }
                else
                {
                    sleeveWidth = damperWidth + 2 * clearance;
                    DamperLogger.Log("Standard detected; applying 50mm clearance each side.");
                    DamperLogger.Log($"Standard sleeve dims (internal): Width={sleeveWidth}, Height={sleeveHeight}");
                    DamperLogger.Log($"Standard sleeve dims (mm): Width={UnitUtils.ConvertFromInternalUnits(sleeveWidth, UnitTypeId.Millimeters)}, Height={UnitUtils.ConvertFromInternalUnits(sleeveHeight, UnitTypeId.Millimeters)}");
                }

                // Activate sleeve symbol if not already active
                if (!sleeveSymbol.IsActive)
                {
                    sleeveSymbol.Activate();
                    DamperLogger.Log("Activated sleeve family symbol.");
                }

                // Retrieve a level for placement
                var level = new FilteredElementCollector(_doc)
                                .OfClass(typeof(Level))
                                .Cast<Level>()
                                .FirstOrDefault();
                if (level == null)
                {
                    DamperLogger.Log("No Level found; aborting placement.");
                    return false;
                }

                // Determine raw damper insertion point (MEP connector or bounding box center)
                XYZ rawCenter = (accessory.MEPModel?.ConnectorManager.Connectors.Cast<Connector>().FirstOrDefault()?.Origin)
                                ?? (damperBBox.Min + damperBBox.Max) * 0.5;
                DamperLogger.Log($"Raw damper center before projection: {rawCenter}");

                // Project raw center onto wall plane (wall face intersection)
                XYZ wallFaceOrigin = (linkedWall.Location as LocationCurve)?.Curve.GetEndPoint(0) ?? rawCenter;
                XYZ wallNormal = linkedWall.Orientation;  // wall face normal
                double distance = (rawCenter - wallFaceOrigin).DotProduct(wallNormal);
                XYZ projectedCenter = rawCenter - wallNormal.Multiply(distance);
                DamperLogger.Log($"Projected damper-wall intersection center: {projectedCenter}");

                // Place the sleeve at the wall/damper intersection center
                var instance = _doc.Create.NewFamilyInstance(
                    projectedCenter,  // intersection center
                    sleeveSymbol,
                    level,
                    StructuralType.NonStructural);
                // Explicitly set the Level parameter for schedule consistency
                Parameter levelParam = instance.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM);
                if (levelParam != null && !levelParam.IsReadOnly)
                {
                    levelParam.Set(level.Id);
                }
                else
                {
                    var levelByName = instance.LookupParameter("Level");
                    if (levelByName != null && !levelByName.IsReadOnly)
                        levelByName.Set(level.Id);
                }
                DamperLogger.Log("Created sleeve at wall/damper intersection center.");

                // Set sleeve dimensions and depth
                instance.LookupParameter("Width")?.Set(sleeveWidth);
                instance.LookupParameter("Height")?.Set(sleeveHeight);
                instance.LookupParameter("Depth")?.Set(linkedWall.Width);
                DamperLogger.Log("Set sleeve dimensions and depth.");

                // Apply MSFD offset so clearances are correct: shift towards right side
                if (isMSFD)
                {
                    double offset = (rightClearance - clearance) / 2.0;
                    XYZ rightDir;
                    // Try duct direction to compute right side vector
                    var ductHost = accessory.Host as Duct;
                    if (ductHost != null)
                    {
                        var locCurve = ductHost.Location as LocationCurve;
                        if (locCurve?.Curve is Line line)
                        {
                            var ductDir = line.Direction.Normalize();
                            rightDir = ductDir.CrossProduct(XYZ.BasisZ).Normalize();
                            DamperLogger.Log($"Computed right direction from duct: {rightDir}");
                        }
                        else
                        {
                            rightDir = accessory.GetTransform().BasisY.Normalize();
                            DamperLogger.Log($"Duct line unavailable; using BasisY: {rightDir}");
                        }
                    }
                    else
                    {
                        rightDir = accessory.GetTransform().BasisY.Normalize();
                        DamperLogger.Log($"No duct host; using BasisY: {rightDir}");
                    }
                    // Move sleeve along rightDir by offset
                    ElementTransformUtils.MoveElement(_doc, instance.Id, rightDir.Multiply(offset));
                    DamperLogger.Log($"MSFD sleeve moved {UnitUtils.ConvertFromInternalUnits(offset, UnitTypeId.Millimeters)} mm towards right side to achieve 50/100 mm clearances.");
                }

                // Determine orientation: prefer duct direction if available
                var transform = accessory.GetTransform();
                XYZ orientationDir;
                if (accessory.Host is Duct hostDuct)
                {
                    var locCurve = hostDuct.Location as LocationCurve;
                    orientationDir = (locCurve != null && locCurve.Curve is Line l) ? l.Direction.Normalize() : transform.BasisX;
                }
                else
                {
                    orientationDir = transform.BasisX;
                }
                DamperLogger.Log($"Orientation for rotation: {orientationDir}");

                // Rotate sleeve for Y-axis aligned dampers
                bool isYAxis = orientationDir.IsAlmostEqualTo(XYZ.BasisY) || orientationDir.IsAlmostEqualTo(-XYZ.BasisY);
                if (isYAxis)
                {
                    DamperLogger.Log("Aligned with Y-axis; rotating sleeve 90° around Z.");
                    var axis = Line.CreateBound(projectedCenter, projectedCenter + XYZ.BasisZ);
                    ElementTransformUtils.RotateElement(_doc, instance.Id, axis, Math.PI / 2);
                }
                else
                {
                    DamperLogger.Log("Not aligned with Y-axis; no rotation applied.");
                }

                // Get reference level from host element (damper)
                Level refLevel = JSE_RevitAddin_MEP_OPENINGS.Helpers.HostLevelHelper.GetHostReferenceLevel(_doc, accessory);
                if (refLevel != null)
                {
                    Parameter schedLevelParam = instance.LookupParameter("Schedule Level");
                    if (schedLevelParam != null && !schedLevelParam.IsReadOnly)
                    {
                        schedLevelParam.Set(refLevel.Id);
                        DamperLogger.Log($"[FireDamperSleevePlacerDebug] Set Schedule Level to {refLevel.Name} for damper {accessory.Id.IntegerValue}");
                    }
                }

                DamperLogger.Log("Set sleeve dimensions and completed placement.");
                return true;
            }
            catch (Exception ex)
            {
                DamperLogger.Log($"Exception in PlaceFireDamperSleeveDebug: {ex.Message}");
                return false;
            }
        }

        private bool BoundingBoxesIntersect(BoundingBoxXYZ bbox1, BoundingBoxXYZ bbox2)
        {
            return !(bbox1.Max.X < bbox2.Min.X || bbox1.Min.X > bbox2.Max.X ||
                     bbox1.Max.Y < bbox2.Min.Y || bbox1.Min.Y > bbox2.Max.Y ||
                     bbox1.Max.Z < bbox2.Min.Z || bbox1.Min.Z > bbox2.Max.Z);
        }
    }
}
