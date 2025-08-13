using System;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Structure;




namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Unified service for placing fire damper/duct accessory sleeves.
    /// Supports both production and debug modes.
    /// </summary>
    public class FireDamperSleevePlacerService
    {
        private readonly Document _doc;
        private readonly bool _enableDebugLogging;

        public FireDamperSleevePlacerService(Document doc, bool enableDebugLogging = false)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _enableDebugLogging = enableDebugLogging;
        }

        private void Log(string message)
        {
            if (_enableDebugLogging)
            {
                DamperLogger.Log(message);
            }
        }

        /// <summary>
        /// Gets the connector side using world coordinates (for MSFD dampers).
        /// </summary>
        public static string GetConnectorSideWorld(FamilyInstance damper, out Connector? connector)
        {
            connector = null;

            var cm = damper.MEPModel?.ConnectorManager;
            if (cm == null) return "Right";

            // Pick the connector with the largest absolute component
            Connector? best = null;
            double maxAbs = 0;

            foreach (Connector c in cm.Connectors)
            {
                var dir = c.CoordinateSystem.BasisX;  // world direction
                double maxComp = Math.Max(Math.Max(Math.Abs(dir.X), Math.Abs(dir.Y)), Math.Abs(dir.Z));
                if (maxComp > maxAbs)
                {
                    maxAbs = maxComp;
                    best = c;
                }
            }

            if (best == null) return "Right";
            connector = best;

            var d = best.CoordinateSystem.BasisX;

            // Precedence: Z for top/bottom, then X for left/right
            if (Math.Abs(d.Z) >= 0.9)
                return d.Z > 0 ? "Top" : "Bottom";
            if (Math.Abs(d.X) >= 0.9)
                return d.X > 0 ? "Right" : "Left";

            return "Right";
        }

        /// <summary>
        /// Gets the connector side using local coordinates (for standard dampers).
        /// </summary>
        private string GetConnectorSideLocal(FamilyInstance damper, out Connector? connector)
        {
            connector = null;
            var cons = (damper.MEPModel as MechanicalFitting)?.ConnectorManager?.Connectors
                    ?? damper.MEPModel?.ConnectorManager?.Connectors;

            if (cons == null || cons.Size == 0)
            {
                Log("[GetConnectorSideLocal] No connectors found. Returning 'Unknown'.");
                return "Unknown";
            }

            // Debug logging for connectors
            if (_enableDebugLogging)
            {
                foreach (Connector c in cons)
                {
                    Log($"[GetConnectorSideLocal] Connector at {c.Origin}, BasisX={c.CoordinateSystem.BasisX}, BasisY={c.CoordinateSystem.BasisY}, BasisZ={c.CoordinateSystem.BasisZ}");
                }
            }

            // Pick the connector closest to origin
            connector = cons.Cast<Connector>()
                             .OrderBy(c => c.Origin.DistanceTo(damper.GetTransform().Origin))
                             .FirstOrDefault();

            if (connector == null)
            {
                Log("[GetConnectorSideLocal] Closest connector is null. Returning 'Unknown'.");
                return "Unknown";
            }

            var damperT = damper.GetTransform();
            var connDir = connector.CoordinateSystem.BasisX; // connector normal

            // Dot products with damper local axes
            double dotX = connDir.DotProduct(damperT.BasisX);
            double dotY = connDir.DotProduct(damperT.BasisY);
            double dotZ = connDir.DotProduct(damperT.BasisZ);

            const double tol = 0.9;

            if (Math.Abs(dotX) >= tol) return dotX > 0 ? "Right" : "Left";
            if (Math.Abs(dotZ) >= tol) return dotZ > 0 ? "Top" : "Bottom";
            
            Log("[GetConnectorSideLocal] Could not determine side, returning 'Unknown'.");
            return "Unknown";
        }

        /// <summary>
        /// Determines if a wall is oriented along the Y-axis.
        /// Y-axis wall: wall runs parallel to Y-axis, so its normal is in X direction.
        /// </summary>
        private static bool IsYAxisWall(Wall wall)
        {
            XYZ wallOrientation = wall.Orientation;
            return Math.Abs(wallOrientation.Y) < 0.1 && Math.Abs(wallOrientation.X) > 0.9;
        }

        /// <summary>
        /// Gets the correct world offset direction based on connector side and wall orientation.
        /// </summary>
        private XYZ GetOffsetDirection(string side, bool isYAxisWall)
        {
            if (isYAxisWall)
            {
                // Y-axis wall: Right/Left are in Y direction, Top/Bottom in Z direction
                return side switch
                {
                    "Right" => XYZ.BasisY,   // Positive Y direction
                    "Left" => -XYZ.BasisY,   // Negative Y direction
                    "Top" => XYZ.BasisZ,     // Positive Z direction
                    "Bottom" => -XYZ.BasisZ, // Negative Z direction
                    _ => XYZ.BasisX          // Fallback
                };
            }
            else
            {
                // X-axis wall: Right/Left are in X direction, Top/Bottom in Z direction
                return side switch
                {
                    "Right" => XYZ.BasisX,   // Positive X direction
                    "Left" => -XYZ.BasisX,   // Negative X direction
                    "Top" => XYZ.BasisZ,     // Positive Z direction
                    "Bottom" => -XYZ.BasisZ, // Negative Z direction
                    _ => XYZ.BasisX          // Fallback
                };
            }
        }

        /// <summary>
        /// Places a sleeve for a fire damper accessory.
        /// </summary>
        /// <param name="accessory">The fire damper element</param>
        /// <param name="sleeveSymbol">The sleeve family symbol to place</param>
        /// <param name="linkedWall">The wall to place the sleeve in</param>
        /// <returns>True if sleeve was placed successfully, false if skipped or failed</returns>
        public bool PlaceFireDamperSleeve(
            FamilyInstance accessory,
            FamilySymbol sleeveSymbol,
            Wall linkedWall)
        {
            Log("FireDamperSleevePlacerService: Starting placement.");

            try
            {
                // Get bounding boxes for the damper and the wall
                var damperBBox = accessory.get_BoundingBox(null);
                var wallBBox = linkedWall.get_BoundingBox(null);

                if (damperBBox == null || wallBBox == null)
                {
                    Log("Bounding box unavailable for damper or wall. Skipping placement.");
                    return false;
                }

                // Check if bounding boxes intersect
                if (!BoundingBoxesIntersect(damperBBox, wallBBox))
                {
                    Log("Damper bounding box does not intersect wall bounding box. Skipping placement.");
                    return false;
                }

                // Retrieve damper dimensions using correct parameter names
                var widthParam = accessory.LookupParameter("Damper Width");
                var heightParam = accessory.LookupParameter("Damper Height");

                if (widthParam == null || heightParam == null)
                {
                    Log("Damper parameters 'Damper Width' or 'Damper Height' not found. Skipping placement.");
                    return false;
                }

                double damperWidth = widthParam.AsDouble();
                double damperHeight = heightParam.AsDouble();

                Log($"Retrieved damper dimensions: Width={damperWidth}, Height={damperHeight}");
                Log($"Damper dimensions (mm): Width={UnitUtils.ConvertFromInternalUnits(damperWidth, UnitTypeId.Millimeters)}, Height={UnitUtils.ConvertFromInternalUnits(damperHeight, UnitTypeId.Millimeters)}");

                // Get symbol type name and determine damper type
                string familyTypeName = accessory.Symbol.Name;
                Log($"Damper family type name: {familyTypeName}");

                // Clearance values
                double clearance50 = UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters);
                double clearance100 = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);

                // Determine damper type (MSFD or Standard)
                string typeNameUpper = familyTypeName?.Trim().ToUpperInvariant() ?? "";
                bool isMSFD = typeNameUpper.Contains("MSFD");
                Log($"Damper type name: '{familyTypeName}', isMSFD: {isMSFD}");

                // Get connector side based on damper type
                Connector? conn;
                string side;
                if (isMSFD)
                {
                    side = GetConnectorSideWorld(accessory, out conn);
                    Log($"[DEBUG] Used GetConnectorSideWorld for MSFD. Detected connector side: {side}");
                    
                    // Debug: Log the actual connector direction for investigation
                    if (_enableDebugLogging && conn != null)
                    {
                        var connDir = conn.CoordinateSystem.BasisX;
                        Log($"[DEBUG] Connector BasisX direction: ({connDir.X:F3}, {connDir.Y:F3}, {connDir.Z:F3})");
                        Log($"[DEBUG] Damper ID: {accessory.Id.Value}");
                    }
                }
                else
                {
                    side = GetConnectorSideLocal(accessory, out conn);
                    Log($"[DEBUG] Used GetConnectorSideLocal for non-MSFD. Detected connector side: {side}");
                }

                // Set clearances for each side
                double left = clearance50, right = clearance50, top = clearance50, bottom = clearance50;
                if (isMSFD)
                {
                    // If connector side is unknown, fallback to Right
                    string effectiveSide = side;
                    if (side == "Unknown")
                    {
                        effectiveSide = "Right";
                        Log("[MSFD CLEARANCE] Connector side unknown, defaulting to 'Right' for clearance.");
                    }
                    
                    // Apply 100mm clearance to the connector side, 50mm elsewhere
                    switch (effectiveSide)
                    {
                        case "Left": left = clearance100; break;
                        case "Right": right = clearance100; break;
                        case "Top": top = clearance100; break;
                        case "Bottom": bottom = clearance100; break;
                    }
                    
                    Log($"[MSFD CLEARANCE] Applied 100mm clearance to {effectiveSide}, 50mm elsewhere. left={UnitUtils.ConvertFromInternalUnits(left, UnitTypeId.Millimeters)}, right={UnitUtils.ConvertFromInternalUnits(right, UnitTypeId.Millimeters)}, top={UnitUtils.ConvertFromInternalUnits(top, UnitTypeId.Millimeters)}, bottom={UnitUtils.ConvertFromInternalUnits(bottom, UnitTypeId.Millimeters)}");
                }
                else
                {
                    Log($"[STANDARD CLEARANCE] 50mm all sides. left={UnitUtils.ConvertFromInternalUnits(left, UnitTypeId.Millimeters)}, right={UnitUtils.ConvertFromInternalUnits(right, UnitTypeId.Millimeters)}, top={UnitUtils.ConvertFromInternalUnits(top, UnitTypeId.Millimeters)}, bottom={UnitUtils.ConvertFromInternalUnits(bottom, UnitTypeId.Millimeters)}");
                }

                // Compute sleeve dimensions
                double sleeveWidth = damperWidth + left + right;
                double sleeveHeight = damperHeight + top + bottom;
                Log($"Sleeve dims (internal): Width={sleeveWidth}, Height={sleeveHeight}");
                Log($"Sleeve dims (mm): Width={UnitUtils.ConvertFromInternalUnits(sleeveWidth, UnitTypeId.Millimeters)}, Height={UnitUtils.ConvertFromInternalUnits(sleeveHeight, UnitTypeId.Millimeters)}");

                // Activate sleeve symbol if not already active
                if (!sleeveSymbol.IsActive)
                {
                    sleeveSymbol.Activate();
                    Log("Activated sleeve family symbol.");
                }

                // Retrieve a level for placement
                var level = new FilteredElementCollector(_doc)
                                .OfClass(typeof(Level))
                                .Cast<Level>()
                                .FirstOrDefault();
                if (level == null)
                {
                    Log("No Level found; aborting placement.");
                    return false;
                }

                // Determine raw damper insertion point (MEP connector or bounding box center)
                XYZ rawCenter = (accessory.MEPModel?.ConnectorManager.Connectors.Cast<Connector>().FirstOrDefault()?.Origin)
                                ?? (damperBBox.Min + damperBBox.Max) * 0.5;
                Log($"Raw damper center before projection: {rawCenter}");

                // Project raw center onto wall plane (wall face intersection)
                XYZ wallFaceOrigin = (linkedWall.Location as LocationCurve)?.Curve.GetEndPoint(0) ?? rawCenter;
                XYZ wallNormal = linkedWall.Orientation;  // wall face normal
                double distance = (rawCenter - wallFaceOrigin).DotProduct(wallNormal);
                XYZ projectedCenter = rawCenter - wallNormal.Multiply(distance);
                Log($"Projected damper-wall intersection center: {projectedCenter}");

                // Calculate offset for MSFD dampers (25mm toward connector direction)
                XYZ offset = XYZ.Zero;
                double offset25 = UnitUtils.ConvertToInternalUnits(25.0, UnitTypeId.Millimeters);
                if (isMSFD && conn != null)
                {
                    // Determine wall orientation and get correct offset direction
                    bool isYAxisWall = IsYAxisWall(linkedWall);
                    Log($"[OFFSET] Wall orientation: ({wallNormal.X:F3}, {wallNormal.Y:F3}, {wallNormal.Z:F3}), isYAxisWall: {isYAxisWall}");
                    
                    XYZ worldOffsetDir = GetOffsetDirection(side, isYAxisWall);
                    offset = worldOffsetDir * offset25;
                    Log($"[OFFSET] Side: {side}, wall type: {(isYAxisWall ? "Y-axis" : "X-axis")}, using world direction: ({worldOffsetDir.X:F3}, {worldOffsetDir.Y:F3}, {worldOffsetDir.Z:F3})");
                }
                
                XYZ sleeveCenter = projectedCenter + offset;
                Log($"[CONNECTOR OFFSET] Sleeve placement center after offset: {sleeveCenter}");

                // DUPLICATION SUPPRESSION: Check if a sleeve already exists at this location
                double sleeveCheckRadius = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters); // 100mm tolerance
                bool duplicateExists = false;
                try
                {
                    duplicateExists = OpeningDuplicationChecker.IsAnySleeveAtLocationEnhanced(_doc, sleeveCenter, sleeveCheckRadius);
                }
                catch (Exception ex)
                {
                    Log($"[ERROR] Duplication suppression check failed: {ex.Message}");
                }
                if (duplicateExists)
                {
                    Log("Duplicate damper sleeve detected at this location. Skipping placement.");
                    return false;
                }

                // Place the sleeve at the computed center
                var instance = _doc.Create.NewFamilyInstance(
                    sleeveCenter,
                    sleeveSymbol,
                    level,
                    StructuralType.NonStructural);
                
                    
                // Set the Level parameter for schedule consistency
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
                Log("Created sleeve at computed center.");

                // Set sleeve dimensions and depth
                instance.LookupParameter("Width")?.Set(sleeveWidth);
                instance.LookupParameter("Height")?.Set(sleeveHeight);
                instance.LookupParameter("Depth")?.Set(linkedWall.Width);
                Log("Set sleeve dimensions and depth.");

                // Rotate sleeve for Y-axis dampers only (simplified rotation logic)
                try
                {
                    var damperDir = accessory.GetTotalTransform().BasisX; // damper flow direction
                    bool isYAxisDamper = Math.Abs(damperDir.Y) > Math.Abs(damperDir.X);
                    
                    if (isYAxisDamper)
                    {
                        Log("[ROTATE] Y-axis damper detected, applying 90° rotation");
                        _doc.Regenerate(); // Force Revit to update the instance location
                        LocationPoint? loc = instance.Location as LocationPoint;
                        if (loc != null)
                        {
                            XYZ rotationPoint = loc.Point;
                            Line rotationAxis = Line.CreateBound(rotationPoint, rotationPoint.Add(XYZ.BasisZ));
                            double rotationAngle = Math.PI / 2; // 90 degrees
                            ElementTransformUtils.RotateElement(_doc, instance.Id, rotationAxis, rotationAngle);
                            Log($"[ROTATE] Rotated Y-axis damper sleeve 90° at {rotationPoint}");
                        }
                        else
                        {
                            Log("[ROTATE ERROR] Instance location is not a LocationPoint!");
                        }
                    }
                    else
                    {
                        Log("[ROTATE] X-axis damper - no rotation needed");
                    }
                }
                catch (Exception rex)
                {
                    Log($"[ERROR] Failed to rotate sleeve: {rex.Message}");
                }

                // Set the Schedule Level parameter
                Level? refLevel = JSE_RevitAddin_MEP_OPENINGS.Helpers.HostLevelHelper.GetHostReferenceLevel(_doc, accessory);
                if (refLevel != null)
                {
                    Parameter schedLevelParam = instance.LookupParameter("Schedule Level");
                    if (schedLevelParam != null && !schedLevelParam.IsReadOnly)
                    {
                        schedLevelParam.Set(refLevel.Id);
                        Log($"[FireDamperSleevePlacerService] Set Schedule Level to {refLevel.Name} for damper {accessory.Id.Value}");
                    }
                }

                Log("Set sleeve dimensions and completed placement.");
                return true;
            }
            catch (Exception ex)
            {
                Log($"Exception in PlaceFireDamperSleeve: {ex.Message}");
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
