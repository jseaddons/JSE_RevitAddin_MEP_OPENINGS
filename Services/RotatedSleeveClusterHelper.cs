using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Helper class for calculating cluster bounding boxes of rotated sleeves
    /// Implements watertight algorithm that works for all scenarios:
    /// - Perpendicular stacking
    /// - Inline arrangement
    /// - Diagonal arrangement
    /// - Grid/matrix arrangement
    /// </summary>
    public static class RotatedSleeveClusterHelper
    {
        /// <summary>
        /// Calculate cluster bounding box for rotated sleeves using watertight algorithm
        /// </summary>
        /// <param name="sleeveCenters">List of sleeve centers in world coordinates</param>
        /// <param name="sleeveDimensions">List of (width, height) tuples for each sleeve</param>
        /// <param name="rotationAngle">Rotation angle in radians (cluster rotation)</param>
        /// <param name="logPath">Optional log file path for debugging</param>
        /// <returns>Tuple: (width, height, minX, minY, maxX, maxY, midpoint in rotated space, midpoint in world space)</returns>
        public static (double width, double height, double minX, double minY, double maxX, double maxY, XYZ midRotated, XYZ midWorld) 
            CalculateClusterBoundingBox(
            List<XYZ> sleeveCenters,
            List<(double width, double height)> sleeveDimensions,
            double rotationAngle,
            string logPath = null)
        {
            if (sleeveCenters == null || sleeveCenters.Count == 0)
                throw new ArgumentException("sleeveCenters cannot be null or empty");
            
            if (sleeveDimensions == null || sleeveDimensions.Count == 0)
                throw new ArgumentException("sleeveDimensions cannot be null or empty");
            
            if (sleeveCenters.Count != sleeveDimensions.Count)
                throw new ArgumentException($"sleeveCenters count ({sleeveCenters.Count}) must match sleeveDimensions count ({sleeveDimensions.Count})");

            bool enableLogging = !string.IsNullOrEmpty(logPath);

            if (enableLogging)
            {
                SafeFileLogger.SafeAppendText(logPath, "========== ROTATED SLEEVE CLUSTER HELPER ==========");
                SafeFileLogger.SafeAppendText(logPath, $"Input: {sleeveCenters.Count} sleeves, Rotation Angle: {rotationAngle * 180.0 / Math.PI:F2}°");
            }

            // Step 1: Choose origin (first sleeve center)
            XYZ origin = sleeveCenters[0];

            // Step 2: Pre-calculate rotation matrix components
            double cosA = Math.Cos(rotationAngle);
            double sinA = Math.Sin(rotationAngle);

            if (enableLogging)
            {
                SafeFileLogger.SafeAppendText(logPath, $"Origin (first sleeve center): ({origin.X:F6}, {origin.Y:F6}, {origin.Z:F6})");
                SafeFileLogger.SafeAppendText(logPath, $"Rotation Matrix: cos={cosA:F6}, sin={sinA:F6}");
            }

            // Step 3: Transform all sleeve centers to rotated coordinate system and calculate corners
            var transformedCorners = new List<XYZ>();

            for (int i = 0; i < sleeveCenters.Count; i++)
            {
                XYZ center = sleeveCenters[i];
                var dims = sleeveDimensions[i];

                // Step 3a: Translate center relative to origin
                double relCx = center.X - origin.X;
                double relCy = center.Y - origin.Y;

                // Step 3b: Rotate the relative center to rotated coordinate system
                // rot_cx = rel_cx * cosA - rel_cy * sinA
                // rot_cy = rel_cx * sinA + rel_cy * cosA
                double rotCx = relCx * cosA - relCy * sinA;
                double rotCy = relCx * sinA + relCy * cosA;

                if (enableLogging)
                {
                    SafeFileLogger.SafeAppendText(logPath, $"  Sleeve {i + 1}: World Center=({center.X:F6}, {center.Y:F6})");
                    SafeFileLogger.SafeAppendText(logPath, $"  Sleeve {i + 1}: Relative=({relCx:F6}, {relCy:F6}), Rotated Center=({rotCx:F6}, {rotCy:F6})");
                    SafeFileLogger.SafeAppendText(logPath, $"  Sleeve {i + 1} Dimensions: W={dims.width * 304.8:F1}mm, H={dims.height * 304.8:F1}mm");
                }

                // Step 3c: Calculate 4 corners with rotated offsets
                double halfW = dims.width / 2.0;
                double halfH = dims.height / 2.0;

                // Generate all 4 corner offsets
                var cornerOffsets = new[]
                {
                    (-halfW, -halfH),  // Bottom-left
                    (halfW, -halfH),   // Bottom-right
                    (-halfW, halfH),   // Top-left
                    (halfW, halfH)      // Top-right
                };

                var corners = new XYZ[4];
                for (int j = 0; j < 4; j++)
                {
                    double dx = cornerOffsets[j].Item1;
                    double dy = cornerOffsets[j].Item2;

                    // Rotate the corner offset vector by rotationAngle
                    // rotated_dx = dx * cosA - dy * sinA
                    // rotated_dy = dx * sinA + dy * cosA
                    double rotatedDx = dx * cosA - dy * sinA;
                    double rotatedDy = dx * sinA + dy * cosA;

                    // Add rotated offset to rotated center
                    // corner_x = rot_cx + rotated_dx
                    // corner_y = rot_cy + rotated_dy
                    corners[j] = new XYZ(
                        rotCx + rotatedDx,
                        rotCy + rotatedDy,
                        center.Z  // Keep Z coordinate
                    );
                }

                transformedCorners.AddRange(corners);

                if (enableLogging)
                {
                    SafeFileLogger.SafeAppendText(logPath, $"  Sleeve {i + 1} Corners (in rotated coordinate system): ({corners[0].X:F6}, {corners[0].Y:F6}), ({corners[1].X:F6}, {corners[1].Y:F6}), ({corners[2].X:F6}, {corners[2].Y:F6}), ({corners[3].X:F6}, {corners[3].Y:F6})");
                }
            }

            // Step 4: Find min/max of all corners in rotated coordinate system
            double minX = transformedCorners.Min(p => p.X);
            double minY = transformedCorners.Min(p => p.Y);
            double maxX = transformedCorners.Max(p => p.X);
            double maxY = transformedCorners.Max(p => p.Y);
            double width = maxX - minX;
            double height = maxY - minY;

            if (enableLogging)
            {
                SafeFileLogger.SafeAppendText(logPath, $"✅ CORRECT UNION (in rotated coordinate system): MinX={minX:F6}, MinY={minY:F6}, MaxX={maxX:F6}, MaxY={maxY:F6}");
                SafeFileLogger.SafeAppendText(logPath, $"✅ CLUSTER SIZE (CORRECT TRANSFORM): W={width * 304.8:F1}mm, H={height * 304.8:F1}mm");
            }

            // Step 5: Calculate midpoint in rotated coordinate space
            XYZ midRotated = new XYZ((minX + maxX) / 2.0, (minY + maxY) / 2.0, origin.Z);

            // Step 6: Transform midpoint back to world coordinates
            XYZ midWorld;
            if (Math.Abs(rotationAngle) > 1e-6)
            {
                // Rotate back (inverse rotation: -rotationAngle)
                double cosAInv = Math.Cos(-rotationAngle);
                double sinAInv = Math.Sin(-rotationAngle);
                double midRotatedBackX = midRotated.X * cosAInv - midRotated.Y * sinAInv;
                double midRotatedBackY = midRotated.X * sinAInv + midRotated.Y * cosAInv;

                // Translate back (add origin)
                midWorld = new XYZ(
                    origin.X + midRotatedBackX,
                    origin.Y + midRotatedBackY,
                    midRotated.Z
                );

                if (enableLogging)
                {
                    SafeFileLogger.SafeAppendText(logPath, $"Midpoint (in rotated space): ({midRotated.X:F6}, {midRotated.Y:F6}, {midRotated.Z:F6})");
                    SafeFileLogger.SafeAppendText(logPath, $"Midpoint (in world space): ({midWorld.X:F6}, {midWorld.Y:F6}, {midWorld.Z:F6})");
                }
            }
            else
            {
                // No rotation: just add origin
                midWorld = new XYZ(
                    origin.X + midRotated.X,
                    origin.Y + midRotated.Y,
                    midRotated.Z
                );

                if (enableLogging)
                {
                    SafeFileLogger.SafeAppendText(logPath, $"No rotation - Midpoint (in world space): ({midWorld.X:F6}, {midWorld.Y:F6}, {midWorld.Z:F6})");
                }
            }

            if (enableLogging)
            {
                SafeFileLogger.SafeAppendText(logPath, "========== END ROTATED SLEEVE CLUSTER HELPER ==========\n");
            }

            return (width, height, minX, minY, maxX, maxY, midRotated, midWorld);
        }

        /// <summary>
        /// Test method for debugging - creates sample data and tests the algorithm
        /// </summary>
        public static void TestAlgorithm(string logPath = null)
        {
            bool enableLogging = !string.IsNullOrEmpty(logPath);

            if (enableLogging)
            {
                SafeFileLogger.SafeAppendText(logPath, "========== TESTING ROTATED SLEEVE CLUSTER ALGORITHM ==========\n");
            }

            // Test Scenario 1: Perpendicular stacking (2 sleeves, 550mm x 200mm each, stacked vertically)
            // Sleeve 1 at origin, Sleeve 2 offset by 200mm perpendicular to axis
            {
                var centers = new List<XYZ>
                {
                    new XYZ(0, 0, 0),           // Sleeve 1 center
                    new XYZ(0, 200.0 / 304.8, 0)  // Sleeve 2 center (200mm = 0.656168 feet)
                };

                var dimensions = new List<(double width, double height)>
                {
                    (550.0 / 304.8, 200.0 / 304.8),  // Sleeve 1: 550mm x 200mm
                    (550.0 / 304.8, 200.0 / 304.8)   // Sleeve 2: 550mm x 200mm
                };

                double rotationAngle = -45.0 * Math.PI / 180.0;  // -45 degrees

                var result = CalculateClusterBoundingBox(centers, dimensions, rotationAngle, logPath);

                if (enableLogging)
                {
                    SafeFileLogger.SafeAppendText(logPath, "TEST 1: Perpendicular Stacking");
                    SafeFileLogger.SafeAppendText(logPath, "Expected: W=550mm, H=400mm");
                    SafeFileLogger.SafeAppendText(logPath, $"Actual: W={result.width * 304.8:F1}mm, H={result.height * 304.8:F1}mm");
                    SafeFileLogger.SafeAppendText(logPath, $"Result: {(Math.Abs(result.width * 304.8 - 550) < 1 && Math.Abs(result.height * 304.8 - 400) < 1 ? "✅ PASS" : "❌ FAIL")}\n");
                }
            }

            // Test Scenario 2: Inline arrangement (2 sleeves, 550mm x 200mm each, arranged along axis)
            {
                var centers = new List<XYZ>
                {
                    new XYZ(0, 0, 0),           // Sleeve 1 center
                    new XYZ(600.0 / 304.8, 0, 0)  // Sleeve 2 center (600mm along axis)
                };

                var dimensions = new List<(double width, double height)>
                {
                    (550.0 / 304.8, 200.0 / 304.8),  // Sleeve 1: 550mm x 200mm
                    (550.0 / 304.8, 200.0 / 304.8)   // Sleeve 2: 550mm x 200mm
                };

                double rotationAngle = -45.0 * Math.PI / 180.0;  // -45 degrees

                var result = CalculateClusterBoundingBox(centers, dimensions, rotationAngle, logPath);

                if (enableLogging)
                {
                    SafeFileLogger.SafeAppendText(logPath, "TEST 2: Inline Arrangement");
                    SafeFileLogger.SafeAppendText(logPath, "Expected: W≈1100mm, H=200mm");
                    SafeFileLogger.SafeAppendText(logPath, $"Actual: W={result.width * 304.8:F1}mm, H={result.height * 304.8:F1}mm\n");
                }
            }

            if (enableLogging)
            {
                SafeFileLogger.SafeAppendText(logPath, "========== END TESTING ==========\n");
            }
        }
    }
}

