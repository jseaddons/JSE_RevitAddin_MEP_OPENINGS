using System;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.DamperDetection
{
    /// <summary>
    /// Detects MEP connector presence and direction for dampers.
    /// Follows Single Responsibility Principle - only responsible for connector detection.
    /// </summary>
    public class DamperConnectorDetector : IDamperConnectorDetector
    {
        /// <summary>
        /// Checks if a damper has at least one MEP connector.
        /// </summary>
        public bool HasMepConnector(FamilyInstance damper)
        {
            if (damper == null)
                return false;

            var cm = damper.MEPModel?.ConnectorManager;
            if (cm == null)
                return false;

            return cm.Connectors != null && cm.Connectors.Size > 0;
        }

        /// <summary>
        /// Detects if a damper has an MEP connector and returns its direction.
        /// </summary>
        /// <param name="damper">The damper FamilyInstance to check</param>
        /// <param name="useWorldCoordinates">If true, uses world coordinates (for non-standard dampers). If false, uses local coordinates (for standard dampers)</param>
        /// <param name="connector">Output parameter containing the detected connector, or null if not found</param>
        /// <param name="wallOrientation">Optional wall orientation ("X" or "Y") to prioritize wall width axis during detection</param>
        /// <returns>Connector direction: World coords ("+X", "-X", "+Y", "-Y", "+Z", "-Z") or Local coords ("Left", "Right", "Top", "Bottom") or empty string if not detected</returns>
        public string DetectConnectorSide(FamilyInstance damper, bool useWorldCoordinates, out Connector connector, string wallOrientation = null)
        {
            connector = null;

            if (damper == null)
                return string.Empty;

            if (useWorldCoordinates)
            {
                return DetectConnectorSideWorld(damper, out connector, wallOrientation);
            }
            else
            {
                return DetectConnectorSideLocal(damper, out connector);
            }
        }

        /// <summary>
        /// Gets the connector side using world coordinates (for non-standard dampers like MSFD).
        /// Returns world coordinate direction: "+X", "-X", "+Y", "-Y", "+Z", "-Z"
        /// This accounts for damper rotation/flip within the linked file (connector direction already reflects this).
        /// </summary>
        /// <param name="wallOrientation">Optional wall orientation ("X" or "Y") to prioritize wall width axis</param>
        private string DetectConnectorSideWorld(FamilyInstance damper, out Connector connector, string wallOrientation = null)
        {
            connector = null;
            
            // Runtime build stamp (assembly file timestamp + version)
            try
            {
                var asm = typeof(DamperConnectorDetector).Assembly;
                string asmLoc = asm.Location;
                DateTime asmWrite = File.Exists(asmLoc) ? File.GetLastWriteTime(asmLoc) : DateTime.MinValue;
                string asmVer = asm.GetName().Version?.ToString() ?? "unknown";
                SafeFileLogger.SafeAppendText("damper_connector_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [BUILD-STAMP] AssemblyLastWrite={asmWrite:yyyy-MM-dd HH:mm:ss}, Version={asmVer}, Assembly={asmLoc}\n");
            }
            catch { }

            var cm = damper.MEPModel?.ConnectorManager;
            if (cm == null)
            {
                SafeFileLogger.SafeAppendText("damper_connector_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss.fff}] [DetectConnectorSideWorld] Damper ID={damper?.Id?.IntegerValue ?? -1}: No ConnectorManager, returning '+X' (fallback)\n");
                return "+X"; // Default fallback
            }

            // Get damper transform for coordinate conversion (CRASH-PROOF: null safety)
            Transform damperTransform;
            try
            {
                damperTransform = damper.GetTransform();
                if (damperTransform == null)
                {
                    SafeFileLogger.SafeAppendText("damper_connector_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss.fff}] [DetectConnectorSideWorld] Damper ID={damper?.Id?.IntegerValue ?? -1}: Null transform, returning '+X' (fallback)\n");
                    return "+X";
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("damper_connector_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss.fff}] [DetectConnectorSideWorld] Damper ID={damper?.Id?.IntegerValue ?? -1}: Exception getting transform: {ex.Message}, returning '+X' (fallback)\n");
                return "+X";
            }
            
            // Get bounding box center for reference (CRASH-PROOF: try-catch + null safety)
            XYZ damperCenter;
            try
            {
                var bbox = damper.get_BoundingBox(null);
                if (bbox != null && bbox.Min != null && bbox.Max != null)
                {
                    damperCenter = (bbox.Min + bbox.Max) / 2.0;
                }
                else
                {
                    damperCenter = damperTransform.Origin ?? XYZ.Zero;
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("damper_connector_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss.fff}] [DetectConnectorSideWorld] Damper ID={damper?.Id?.IntegerValue ?? -1}: Exception getting bounding box: {ex.Message}, using transform origin\n");
                damperCenter = damperTransform.Origin ?? XYZ.Zero;
            }
            
            // Get flip state for logging (CRASH-PROOF: try-catch)
            bool isFacingFlipped = false;
            bool isHandFlipped = false;
            try
            {
                isFacingFlipped = damper.FacingFlipped;
                isHandFlipped = damper.HandFlipped;
            }
            catch { /* Flip state not critical, continue */ }
            
            // Find the MEP connector - prefer the one furthest from center (CRASH-PROOF: safe iteration)
            int connectorCount = 0;
            Connector best = null;
            double maxDistance = 0;
            
            try
            {
                connectorCount = cm.Connectors?.Size ?? 0;
                if (cm.Connectors != null)
                {
                    foreach (Connector c in cm.Connectors)
                    {
                        if (c == null || c.Origin == null) continue; // CRASH-PROOF: skip null connectors
                        
                        try
                        {
                            double distance = c.Origin.DistanceTo(damperCenter);
                            if (!double.IsNaN(distance) && !double.IsInfinity(distance) && distance > maxDistance) // CRASH-PROOF: validate numeric
                            {
                                maxDistance = distance;
                                best = c;
                            }
                        }
                        catch { /* Skip invalid connector */ }
                    }
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("damper_connector_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss.fff}] [DetectConnectorSideWorld] Damper ID={damper?.Id?.IntegerValue ?? -1}: Exception iterating connectors: {ex.Message}\n");
            }

            if (best == null)
            {
                SafeFileLogger.SafeAppendText("damper_connector_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss.fff}] [DetectConnectorSideWorld] Damper ID={damper?.Id?.IntegerValue ?? -1}: No connectors found, returning '+X' (fallback)\n");
                return "+X";
            }

            connector = best;
            XYZ connectorOrigin;
            XYZ connectorBasisX;
            XYZ positionVector;
            
            // CRASH-PROOF: Safe access to connector properties with validation
            try
            {
                connectorOrigin = best.Origin ?? XYZ.Zero;
                
                // ✅ KEY INSIGHT: The connector's BasisX in Revit points INWARD (toward the damper body)
                // NOT outward. We need to NEGATE it to get the side the connector is on.
                // 
                // The connector.CoordinateSystem is already in WORLD coordinates for placed instances.
                // We don't need to transform it - Revit gives us the world-space direction directly.
                // But we DO need to negate it because BasisX points inward, not outward.
                
                var coordSystem = best.CoordinateSystem;
                if (coordSystem == null || coordSystem.BasisX == null)
                {
                    SafeFileLogger.SafeAppendText("damper_connector_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss.fff}] [DetectConnectorSideWorld] Damper ID={damper?.Id?.IntegerValue ?? -1}: Null CoordinateSystem, returning '+X' (fallback)\n");
                    return "+X";
                }
                
                connectorBasisX = -coordSystem.BasisX; // Already in world coordinates, negated to get outward direction
                
                // For comparison/debugging, also get position-based direction
                positionVector = connectorOrigin - damperCenter;
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("damper_connector_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss.fff}] [DetectConnectorSideWorld] Damper ID={damper?.Id?.IntegerValue ?? -1}: Exception accessing connector properties: {ex.Message}, returning '+X' (fallback)\n");
                return "+X";
            }
            
            // Calculate absolute values (CRASH-PROOF: validate numeric components)
            double absX = 0, absY = 0, absZ = 0;
            double posAbsX = 0, posAbsY = 0, posAbsZ = 0;
            
            try
            {
                absX = double.IsNaN(connectorBasisX.X) ? 0 : Math.Abs(connectorBasisX.X);
                absY = double.IsNaN(connectorBasisX.Y) ? 0 : Math.Abs(connectorBasisX.Y);
                absZ = double.IsNaN(connectorBasisX.Z) ? 0 : Math.Abs(connectorBasisX.Z);
                
                // Position vector absolutes for logging
                posAbsX = double.IsNaN(positionVector.X) ? 0 : Math.Abs(positionVector.X);
                posAbsY = double.IsNaN(positionVector.Y) ? 0 : Math.Abs(positionVector.Y);
                posAbsZ = double.IsNaN(positionVector.Z) ? 0 : Math.Abs(positionVector.Z);
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("damper_connector_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss.fff}] [DetectConnectorSideWorld] Damper ID={damper?.Id?.IntegerValue ?? -1}: Exception calculating absolutes: {ex.Message}\n");
            }
            
            // ✅ DETERMINE DIRECTION: Prefer position along wall width axis; fallback to BasisX
            // For walls, using the connector's offset from damper center along the wall width axis
            // is robust against BasisX orientation ambiguity. Fallback to BasisX sign when needed.
            string detectedDirection;
            const double axisPosThreshold = 0.02; // ~6mm in feet
            
            // Check Z first (vertical)
            if (absZ >= absX && absZ >= absY && absZ > 0.5)
            {
                detectedDirection = connectorBasisX.Z > 0 ? "+Z" : "-Z";
            }
            // For Y-walls, prioritize position along Y (fallback to BasisX if near zero)
            // NOTE: positionVector = connector - center, so negate to get "side connector is on"
            else if (!string.IsNullOrEmpty(wallOrientation) && 
                     string.Equals(wallOrientation, "Y", StringComparison.OrdinalIgnoreCase))
            {
                if (Math.Abs(positionVector.Y) >= axisPosThreshold)
                    detectedDirection = positionVector.Y > 0 ? "-Y" : "+Y"; // Flipped: connector offset +Y means damper is on -Y side
                else
                    detectedDirection = connectorBasisX.Y > 0 ? "+Y" : "-Y";
            }
            // For X-walls, prioritize position along X (fallback to BasisX if near zero)
            else if (!string.IsNullOrEmpty(wallOrientation) && 
                     string.Equals(wallOrientation, "X", StringComparison.OrdinalIgnoreCase))
            {
                if (Math.Abs(positionVector.X) >= axisPosThreshold)
                    detectedDirection = positionVector.X > 0 ? "-X" : "+X"; // Flipped: connector offset +X means damper is on -X side
                else
                    detectedDirection = connectorBasisX.X > 0 ? "+X" : "-X";
            }
            // Default: use dominant axis
            else if (absY >= absX && absY > 0.5)
            {
                detectedDirection = connectorBasisX.Y > 0 ? "+Y" : "-Y";
            }
            else if (absX > 0.5)
            {
                detectedDirection = connectorBasisX.X > 0 ? "+X" : "-X";
            }
            else
            {
                // Fallback
                detectedDirection = "+X";
            }
            
            // ✅ COMPREHENSIVE LOGGING
            string logMessage = $"[{DateTime.Now:HH:mm:ss.fff}] [DetectConnectorSideWorld] " +
                $"Damper ID={damper?.Id?.IntegerValue ?? -1}, " +
                $"Family='{damper?.Symbol?.Family?.Name ?? "Unknown"}', " +
                $"Type='{damper?.Symbol?.Name ?? "Unknown"}', " +
                $"FacingFlipped={isFacingFlipped}, HandFlipped={isHandFlipped}, " +
                $"TotalConnectors={connectorCount}, " +
                $"DamperCenter=({damperCenter.X:F4}, {damperCenter.Y:F4}, {damperCenter.Z:F4}), " +
                $"ConnectorOrigin=({connectorOrigin.X:F4}, {connectorOrigin.Y:F4}, {connectorOrigin.Z:F4}), " +
                $"ConnectorBasisX=({connectorBasisX.X:F4}, {connectorBasisX.Y:F4}, {connectorBasisX.Z:F4}) [WORLD COORDS - connector facing direction], " +
                $"BasisXAbs: X={absX:F4}, Y={absY:F4}, Z={absZ:F4}, " +
                $"PositionVector=({positionVector.X:F4}, {positionVector.Y:F4}, {positionVector.Z:F4}), " +
                $"PosAbs: X={posAbsX:F4}, Y={posAbsY:F4}, Z={posAbsZ:F4}, " +
                $"WallOrientation='{wallOrientation ?? "null"}', " +
                $"DetectedDirection='{detectedDirection}'\n";
            
            SafeFileLogger.SafeAppendText("damper_connector_debug.log", logMessage);
            System.Diagnostics.Debug.WriteLine($"[DamperConnectorDetector] {logMessage.Trim()}");
            
            return detectedDirection;
        }

        /// <summary>
        /// Gets the connector side using local coordinates (for standard dampers).
        /// </summary>
        private string DetectConnectorSideLocal(FamilyInstance damper, out Connector connector)
        {
            connector = null;

            var cons = (damper.MEPModel as MechanicalFitting)?.ConnectorManager?.Connectors
                    ?? damper.MEPModel?.ConnectorManager?.Connectors;

            if (cons == null || cons.Size == 0)
                return "Unknown";

            // Pick the connector closest to origin
            connector = cons.Cast<Connector>()
                             .OrderBy(c => c.Origin.DistanceTo(damper.GetTransform().Origin))
                             .FirstOrDefault();

            if (connector == null)
                return "Unknown";

            var damperT = damper.GetTransform();
            var connDir = connector.CoordinateSystem.BasisX; // connector normal

            // Dot products with damper local axes
            double dotX = connDir.DotProduct(damperT.BasisX);
            double dotY = connDir.DotProduct(damperT.BasisY);
            double dotZ = connDir.DotProduct(damperT.BasisZ);

            const double tol = 0.9;

            if (Math.Abs(dotX) >= tol)
                return dotX > 0 ? "Right" : "Left";
            if (Math.Abs(dotZ) >= tol)
                return dotZ > 0 ? "Top" : "Bottom";

            return "Unknown";
        }
    }
}

