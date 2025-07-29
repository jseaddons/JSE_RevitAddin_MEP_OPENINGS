using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Handles specialized logging for sleeve placement with counts and missing elements
    /// </summary>
    public static class SleeveLogManager
    {
        // Log file paths
        private static readonly string SummaryLogPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SleeveManager_Summary.log");
        private static readonly string PipeLogPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SleeveManager_Pipes.log");
        private static readonly string DuctLogPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SleeveManager_Ducts.log");
        private static readonly string CableTrayLogPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SleeveManager_CableTrays.log");
        private static readonly string DebugLogPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SleeveManager_Debug.log");

        // Statistics tracking
        private static int _totalPipesWithWalls = 0;
        private static int _totalPipeSleevesPLaced = 0;
        private static int _totalDuctsWithWalls = 0;
        private static int _totalDuctSleevesPLaced = 0;
        private static int _totalCableTraysWithWalls = 0;
        private static int _totalCableTraySleevesPLaced = 0;
        private static List<string> _missingPipes = new List<string>();
        private static List<string> _missingDucts = new List<string>();
        private static List<string> _missingCableTrays = new List<string>();
        private static List<string> _warnings = new List<string>();
        private static HashSet<int> _processedElements = new HashSet<int>();

        // Element tracking for cross-checking
        private static HashSet<int> _allModelPipes = new HashSet<int>();
        private static HashSet<int> _allModelDucts = new HashSet<int>();
        private static HashSet<int> _allModelCableTrays = new HashSet<int>();

        /// <summary>
        /// Initialize or reset all logs
        /// </summary>
        public static void InitializeLogs()
        {
            try
            {
                // Create header for each log
                string header = $"===== JSE MEP SLEEVE PLACEMENT LOG {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====\r\n";

                // Reset statistics
                _totalPipesWithWalls = 0;
                _totalPipeSleevesPLaced = 0;
                _totalDuctsWithWalls = 0;
                _totalDuctSleevesPLaced = 0;
                _totalCableTraysWithWalls = 0;
                _totalCableTraySleevesPLaced = 0;
                _missingPipes.Clear();
                _missingDucts.Clear();
                _missingCableTrays.Clear();
                _warnings.Clear();
                _processedElements.Clear();
                _allModelPipes.Clear();
                _allModelDucts.Clear();
                _allModelCableTrays.Clear();

                // Create/overwrite each log file
                File.WriteAllText(SummaryLogPath, header + "SUMMARY LOG\r\n\r\n");
                File.WriteAllText(PipeLogPath, header + "PIPE SLEEVE LOG\r\n\r\n");
                File.WriteAllText(DuctLogPath, header + "DUCT SLEEVE LOG\r\n\r\n");
                File.WriteAllText(CableTrayLogPath, header + "CABLE TRAY SLEEVE LOG\r\n\r\n");
                File.WriteAllText(DebugLogPath, header + "DETAILED DEBUG LOG\r\n\r\n");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to initialize log files: {ex.Message}");
            }
        }

        /// <summary>
        /// Register all elements from the model for later verification
        /// </summary>
        public static void RegisterModelElements(IEnumerable<int> pipeIds, IEnumerable<int> ductIds, IEnumerable<int> cableTrayIds)
        {
            foreach (var id in pipeIds) _allModelPipes.Add(id);
            foreach (var id in ductIds) _allModelDucts.Add(id);
            foreach (var id in cableTrayIds) _allModelCableTrays.Add(id);

            AppendToLog(DebugLogPath, $"Registered model elements - Pipes: {_allModelPipes.Count}, Ducts: {_allModelDucts.Count}, CableTrays: {_allModelCableTrays.Count}");

            // Log all pipe IDs for debugging
            StringBuilder pipeIds_str = new StringBuilder("Registered pipe IDs: ");
            foreach (var id in _allModelPipes)
            {
                pipeIds_str.Append($"{id}, ");
            }
            AppendToLog(DebugLogPath, pipeIds_str.ToString());
        }
        /// <summary>
        /// Log a found pipe wall intersection
        /// </summary>
        public static void LogPipeWallIntersection(int elementId, XYZ location, double diameter, int wallId = 0, XYZ? wallOrientation = null)
        {
            _totalPipesWithWalls++;
            _processedElements.Add(elementId);

            string wallInfo = wallId > 0 ? $"wall {wallId}" : "wall";
            string message = $"Found pipe {elementId} intersecting {wallInfo} at {FormatXYZ(location)}, diameter: {FormatMM(diameter)}mm";
            AppendToLog(PipeLogPath, message);

            // More detailed info in debug log
            AppendToLog(DebugLogPath, $"PIPE-WALL INTERSECTION - Element: {elementId}, Wall: {(wallId > 0 ? wallId.ToString() : "unknown")}");
            AppendToLog(DebugLogPath, $"  Location: {FormatXYZ(location)}");
            if (wallOrientation != null)
            {
                AppendToLog(DebugLogPath, $"  Wall orientation: {FormatXYZ(wallOrientation)}");
            }
            AppendToLog(DebugLogPath, $"  Pipe diameter: {FormatMM(diameter)}mm");
        }

        /// <summary>
        /// Log a successful pipe sleeve placement
        /// </summary>
        public static void LogPipeSleeveSuccess(int elementId, int sleeveId, double sleeveDiameter, XYZ sleevePosition)
        {
            _totalPipeSleevesPLaced++;

            string message = $"SUCCESS: Placed sleeve {sleeveId} for pipe {elementId}, sleeve diameter: {FormatMM(sleeveDiameter)}mm";
            AppendToLog(PipeLogPath, message);

            // More detailed info in debug log
            AppendToLog(DebugLogPath, $"PIPE SLEEVE PLACED - Pipe: {elementId}, Sleeve: {sleeveId}");
            AppendToLog(DebugLogPath, $"  Final position: {FormatXYZ(sleevePosition)}");
            AppendToLog(DebugLogPath, $"  Sleeve diameter: {FormatMM(sleeveDiameter)}mm");
        }

        /// <summary>
        /// Log a failed pipe sleeve placement
        /// </summary>
        public static void LogPipeSleeveFailure(int elementId, string reason)
        {
            _processedElements.Add(elementId);
            string message = $"FAILED: Could not place sleeve for pipe {elementId}. Reason: {reason}";
            AppendToLog(PipeLogPath, message);
            _missingPipes.Add($"Pipe {elementId}: {reason}");

            // More detailed info in debug log
            AppendToLog(DebugLogPath, $"PIPE SLEEVE FAILED - Element: {elementId}");
            AppendToLog(DebugLogPath, $"  Reason: {reason}");
        }

        /// <summary>
        /// Log a found duct wall intersection
        /// </summary>
        public static void LogDuctWallIntersection(int elementId, XYZ location, double width, double height, int wallId = 0, XYZ? wallOrientation = null)
        {
            _totalDuctsWithWalls++;
            _processedElements.Add(elementId);
            string wallInfo = wallId > 0 ? $"wall {wallId}" : "wall";
            string message = $"Found duct {elementId} intersecting {wallInfo} at {FormatXYZ(location)}, size: {FormatMM(width)}mm × {FormatMM(height)}mm";
            AppendToLog(DuctLogPath, message);

            // More detailed info in debug log
            AppendToLog(DebugLogPath, $"DUCT-WALL INTERSECTION - Element: {elementId}, Wall: {(wallId > 0 ? wallId.ToString() : "unknown")}");
            AppendToLog(DebugLogPath, $"  Location: {FormatXYZ(location)}");
            if (wallOrientation != null)
            {
                AppendToLog(DebugLogPath, $"  Wall orientation: {FormatXYZ(wallOrientation)}");
            }
            AppendToLog(DebugLogPath, $"  Duct size: {FormatMM(width)}mm × {FormatMM(height)}mm");
        }

        /// <summary>
        /// Log a successful duct sleeve placement
        /// </summary>
        public static void LogDuctSleeveSuccess(int elementId, int sleeveId, double sleeveWidth, double sleeveHeight, XYZ sleevePosition)
        {
            _totalDuctSleevesPLaced++;

            string message = $"SUCCESS: Placed sleeve {sleeveId} for duct {elementId}, sleeve size: {FormatMM(sleeveWidth)}mm × {FormatMM(sleeveHeight)}mm";
            AppendToLog(DuctLogPath, message);

            // More detailed info in debug log
            AppendToLog(DebugLogPath, $"DUCT SLEEVE PLACED - Duct: {elementId}, Sleeve: {sleeveId}");
            AppendToLog(DebugLogPath, $"  Final position: {FormatXYZ(sleevePosition)}");
            AppendToLog(DebugLogPath, $"  Sleeve size: {FormatMM(sleeveWidth)}mm × {FormatMM(sleeveHeight)}mm");
        }

        /// <summary>
        /// Log a failed duct sleeve placement
        /// </summary>
        public static void LogDuctSleeveFailure(int elementId, string reason)
        {
            _processedElements.Add(elementId);
            string message = $"FAILED: Could not place sleeve for duct {elementId}. Reason: {reason}";
            AppendToLog(DuctLogPath, message);
            _missingDucts.Add($"Duct {elementId}: {reason}");

            // More detailed info in debug log
            AppendToLog(DebugLogPath, $"DUCT SLEEVE FAILED - Element: {elementId}");
            AppendToLog(DebugLogPath, $"  Reason: {reason}");
        }

        /// <summary>
        /// Log a found cable tray wall intersection
        /// </summary>
        public static void LogCableTrayWallIntersection(int elementId, XYZ location, double width, double height, int wallId = 0, XYZ? wallOrientation = null)
        {
            _totalCableTraysWithWalls++;
            _processedElements.Add(elementId);

            string wallInfo = wallId > 0 ? $"wall {wallId}" : "wall";
            string message = $"Found cable tray {elementId} intersecting {wallInfo} at {FormatXYZ(location)}, size: {FormatMM(width)}mm × {FormatMM(height)}mm";
            AppendToLog(CableTrayLogPath, message);

            // More detailed info in debug log
            AppendToLog(DebugLogPath, $"CABLETRAY-WALL INTERSECTION - Element: {elementId}, Wall: {(wallId > 0 ? wallId.ToString() : "unknown")}");
            AppendToLog(DebugLogPath, $"  Location: {FormatXYZ(location)}");
            if (wallOrientation != null)
            {
                AppendToLog(DebugLogPath, $"  Wall orientation: {FormatXYZ(wallOrientation)}");
            }
            AppendToLog(DebugLogPath, $"  Cable tray size: {FormatMM(width)}mm × {FormatMM(height)}mm");
        }

        /// <summary>
        /// Log a successful cable tray sleeve placement
        /// </summary>
        public static void LogCableTraySleeveSuccess(int elementId, int sleeveId, double sleeveWidth, double sleeveHeight, XYZ sleevePosition)
        {
            _totalCableTraySleevesPLaced++;

            string message = $"SUCCESS: Placed sleeve {sleeveId} for cable tray {elementId}, sleeve size: {FormatMM(sleeveWidth)}mm × {FormatMM(sleeveHeight)}mm";
            AppendToLog(CableTrayLogPath, message);

            // More detailed info in debug log
            AppendToLog(DebugLogPath, $"CABLETRAY SLEEVE PLACED - Cable Tray: {elementId}, Sleeve: {sleeveId}");
            AppendToLog(DebugLogPath, $"  Final position: {FormatXYZ(sleevePosition)}");
            AppendToLog(DebugLogPath, $"  Sleeve size: {FormatMM(sleeveWidth)}mm × {FormatMM(sleeveHeight)}mm");
        }

        /// <summary>
        /// Log a failed cable tray sleeve placement
        /// </summary>
        public static void LogCableTraySleeveFailure(int elementId, string reason)
        {
            _processedElements.Add(elementId);
            string message = $"FAILED: Could not place sleeve for cable tray {elementId}. Reason: {reason}";
            AppendToLog(CableTrayLogPath, message);
            _missingCableTrays.Add($"Cable Tray {elementId}: {reason}");

            // More detailed info in debug log
            AppendToLog(DebugLogPath, $"CABLETRAY SLEEVE FAILED - Element: {elementId}");
            AppendToLog(DebugLogPath, $"  Reason: {reason}");
        }

        /// <summary>
        /// Log info-level messages for cable tray sleeve parameter actions (e.g., HostOrientation)
        /// </summary>
        public static void LogCableTraySleeveInfo(int elementId, string info)
        {
            string message = $"INFO: Cable tray {elementId}: {info}";
            AppendToLog(CableTrayLogPath, message);
            // Also log to debug for traceability
            AppendToLog(DebugLogPath, $"CABLETRAY SLEEVE INFO - Element: {elementId}");
            AppendToLog(DebugLogPath, $"  Info: {info}");
        }

        /// <summary>
        /// Log a warning message to all logs
        /// </summary>
        public static void LogWarning(string message)
        {
            _warnings.Add(message);
            string formattedMessage = $"WARNING: {message}";
            AppendToLog(SummaryLogPath, formattedMessage);
            AppendToLog(PipeLogPath, formattedMessage);
            AppendToLog(DuctLogPath, formattedMessage);
            AppendToLog(CableTrayLogPath, formattedMessage);
            AppendToLog(DebugLogPath, formattedMessage);
        }

        /// <summary>
        /// Log a pipe inspection message for terminal connections or fitting issues
        /// </summary>
        public static void LogPipeInspection(int elementId, string message)
        {
            _processedElements.Add(elementId);
            string logMessage = $"INSPECT: Pipe {elementId} - {message}";
            AppendToLog(PipeLogPath, logMessage);

            // More detailed info in debug log
            AppendToLog(DebugLogPath, $"PIPE INSPECTION - Element: {elementId}");
            AppendToLog(DebugLogPath, $"  Note: {message}");
        }

        /// <summary>
        /// Log a duct inspection message for terminal connections or fitting issues
        /// </summary>
        public static void LogDuctInspection(int elementId, string message)
        {
            _processedElements.Add(elementId);
            string logMessage = $"INSPECT: Duct {elementId} - {message}";
            AppendToLog(DuctLogPath, logMessage);

            // More detailed info in debug log
            AppendToLog(DebugLogPath, $"DUCT INSPECTION - Element: {elementId}");
            AppendToLog(DebugLogPath, $"  Note: {message}");
        }

        /// <summary>
        /// Log a detailed geometric analysis message about element-wall proximity
        /// </summary>
        public static void LogGeometryAnalysis(string elementType, int elementId, int wallId, double distance, XYZ elementPoint, XYZ wallPoint)
        {
            string message = $"GEOMETRY: {elementType} {elementId} is {FormatMM(distance)}mm from wall {wallId}";

            // Log to both element-specific and debug logs
            if (elementType.Equals("Pipe", StringComparison.OrdinalIgnoreCase))
            {
                AppendToLog(PipeLogPath, message);
            }
            else if (elementType.Equals("Duct", StringComparison.OrdinalIgnoreCase))
            {
                AppendToLog(DuctLogPath, message);
            }
            else if (elementType.Equals("CableTray", StringComparison.OrdinalIgnoreCase))
            {
                AppendToLog(CableTrayLogPath, message);
            }

            // Detailed debug info
            AppendToLog(DebugLogPath, $"WALL PROXIMITY - {elementType}: {elementId}, Wall: {wallId}");
            AppendToLog(DebugLogPath, $"  Distance: {FormatMM(distance)}mm");
            AppendToLog(DebugLogPath, $"  Element point: {FormatXYZ(elementPoint)}");
            AppendToLog(DebugLogPath, $"  Wall point: {FormatXYZ(wallPoint)}");
        }

        /// <summary>
        /// Log diagnostic information to debug log only
        /// </summary>
        public static void LogDebug(string message)
        {
            AppendToLog(DebugLogPath, $"DEBUG: {message}");
        }

        /// <summary>
        /// Finalize logs with summary statistics and identify any completely missing elements
        /// </summary>
        public static void FinalizeLogs()
        {
            try
            {
                // Identify elements that were in the model but never processed
                FindUnprocessedElements();

                StringBuilder summary = new StringBuilder();
                summary.AppendLine("\r\n===== PLACEMENT SUMMARY =====");
                summary.AppendLine($"Total pipes with wall intersections: {_totalPipesWithWalls}");
                summary.AppendLine($"Total pipe sleeves successfully placed: {_totalPipeSleevesPLaced}");
                summary.AppendLine($"Missing pipe sleeves: {_totalPipesWithWalls - _totalPipeSleevesPLaced}");
                summary.AppendLine();
                summary.AppendLine($"Total ducts with wall intersections: {_totalDuctsWithWalls}");
                summary.AppendLine($"Total duct sleeves successfully placed: {_totalDuctSleevesPLaced}");
                summary.AppendLine($"Missing duct sleeves: {_totalDuctsWithWalls - _totalDuctSleevesPLaced}");
                summary.AppendLine();
                summary.AppendLine($"Total cable trays with wall intersections: {_totalCableTraysWithWalls}");
                summary.AppendLine($"Total cable tray sleeves successfully placed: {_totalCableTraySleevesPLaced}");
                summary.AppendLine($"Missing cable tray sleeves: {_totalCableTraysWithWalls - _totalCableTraySleevesPLaced}");

                // Add warnings if any
                if (_warnings.Count > 0)
                {
                    summary.AppendLine("\r\n===== WARNINGS =====");
                    foreach (var warning in _warnings)
                    {
                        summary.AppendLine($"- {warning}");
                    }
                }

                // Add details about missing sleeves if any
                if (_missingPipes.Count > 0)
                {
                    summary.AppendLine("\r\n===== MISSING PIPE SLEEVES =====");
                    foreach (var missing in _missingPipes)
                    {
                        summary.AppendLine($"- {missing}");
                    }
                }

                if (_missingDucts.Count > 0)
                {
                    summary.AppendLine("\r\n===== MISSING DUCT SLEEVES =====");
                    foreach (var missing in _missingDucts)
                    {
                        summary.AppendLine($"- {missing}");
                    }
                }

                if (_missingCableTrays.Count > 0)
                {
                    summary.AppendLine("\r\n===== MISSING CABLE TRAY SLEEVES =====");
                    foreach (var missing in _missingCableTrays)
                    {
                        summary.AppendLine($"- {missing}");
                    }
                }

                // Write summary to all log files
                string summaryText = summary.ToString();
                AppendToLog(SummaryLogPath, summaryText);
                AppendToLog(PipeLogPath, summaryText);
                AppendToLog(DuctLogPath, summaryText);
                AppendToLog(CableTrayLogPath, summaryText);
                AppendToLog(DebugLogPath, summaryText);
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to finalize log files: {ex.Message}");
            }
        }

        /// <summary>
        /// Find elements that were in the model but never processed (missed entirely by the detection system)
        /// </summary>
        private static void FindUnprocessedElements()
        {
            // Check for completely unprocessed pipes
            foreach (var pipeId in _allModelPipes)
            {
                if (!_processedElements.Contains(pipeId))
                {
                    string message = $"Pipe {pipeId}: COMPLETELY MISSED - No wall intersection detected";
                    _missingPipes.Add(message);
                    AppendToLog(PipeLogPath, $"MISSED: {message}");
                    AppendToLog(DebugLogPath, $"UNPROCESSED ELEMENT: {message}");
                }
            }

            // Check for completely unprocessed ducts
            foreach (var ductId in _allModelDucts)
            {
                if (!_processedElements.Contains(ductId))
                {
                    string message = $"Duct {ductId}: COMPLETELY MISSED - No wall intersection detected";
                    _missingDucts.Add(message);
                    AppendToLog(DuctLogPath, $"MISSED: {message}");
                    AppendToLog(DebugLogPath, $"UNPROCESSED ELEMENT: {message}");
                }
            }

            // Check for completely unprocessed cable trays
            foreach (var trayId in _allModelCableTrays)
            {
                if (!_processedElements.Contains(trayId))
                {
                    string message = $"Cable Tray {trayId}: COMPLETELY MISSED - No wall intersection detected";
                    _missingCableTrays.Add(message);
                    AppendToLog(CableTrayLogPath, $"MISSED: {message}");
                    AppendToLog(DebugLogPath, $"UNPROCESSED ELEMENT: {message}");
                }
            }
        }

        // Helper methods
        private static string FormatXYZ(XYZ point)
        {
            if (point == null) return "(null)";
            return $"({FormatMM(point.X)}, {FormatMM(point.Y)}, {FormatMM(point.Z)})";
        }

        private static string FormatMM(double value)
        {
            // Convert internal units to millimeters and format to 2 decimal places
            return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Millimeters).ToString("F2");
        }

        private static void AppendToLog(string logPath, string message)
        {
            try
            {
                File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] {message}\r\n");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to write to log {Path.GetFileName(logPath)}: {ex.Message}");
            }
        }
    }
}
