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
        // Log filenames
        private static readonly string SummaryLogName = "SleeveManager_Summary.log";
        private static readonly string PipeLogName = "SleeveManager_Pipes.log";
        private static readonly string DuctLogName = "SleeveManager_Ducts.log";
        private static readonly string CableTrayLogName = "SleeveManager_CableTrays.log";
        private static readonly string DebugLogName = "SleeveManager_Debug.log";

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
            if (!DebugLogger.IsEnabled) return;
            try
            {
                // Create header for each log
                string header = $"===== JSE MEP SLEEVE PLACEMENT LOG {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====\r\n";
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
                // Create/overwrite each log file header
                SafeFileLogger.SafeAppendTextAlways(SummaryLogName, header + "SUMMARY LOG\n");
                SafeFileLogger.SafeAppendTextAlways(PipeLogName, header + "PIPE SLEEVE LOG\n");
                SafeFileLogger.SafeAppendTextAlways(DuctLogName, header + "DUCT SLEEVE LOG\n");
                SafeFileLogger.SafeAppendTextAlways(CableTrayLogName, header + "CABLE TRAY SLEEVE LOG\n");
                SafeFileLogger.SafeAppendTextAlways(DebugLogName, header + "DETAILED DEBUG LOG\n");
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"Failed to initialize log files: {ex.Message}");
            }
        }

        /// <summary>
        /// Register all elements from the model for later verification
        /// </summary>
        public static void RegisterModelElements(IEnumerable<int> pipeIds, IEnumerable<int> ductIds, IEnumerable<int> cableTrayIds)
        {
            if (!DebugLogger.IsEnabled) return;
            foreach (var id in pipeIds) _allModelPipes.Add(id);
            foreach (var id in ductIds) _allModelDucts.Add(id);
            foreach (var id in cableTrayIds) _allModelCableTrays.Add(id);

            AppendToLog(DebugLogName, $"Registered model elements - Pipes: {_allModelPipes.Count}, Ducts: {_allModelDucts.Count}, CableTrays: {_allModelCableTrays.Count}");

            // Log all pipe IDs for debugging
            StringBuilder pipeIds_str = new StringBuilder("Registered pipe IDs: ");
            foreach (var id in _allModelPipes)
            {
                pipeIds_str.Append($"{id}, ");
            }
            AppendToLog(DebugLogName, pipeIds_str.ToString());
        }
        /// <summary>
        /// Log a found pipe wall intersection
        /// </summary>
        public static void LogPipeWallIntersection(int elementId, XYZ location, double diameter, int wallId = 0, XYZ? wallOrientation = null)
        {
            if (!DebugLogger.IsEnabled) return;
            _totalPipesWithWalls++;
            _processedElements.Add(elementId);

            string wallInfo = wallId > 0 ? $"wall {wallId}" : "wall";
            string message = $"Found pipe {elementId} intersecting {wallInfo} at {FormatXYZ(location)}, diameter: {FormatMM(diameter)}mm";
            AppendToLog(PipeLogName, message);

            // More detailed info in debug log
            AppendToLog(DebugLogName, $"PIPE-WALL INTERSECTION - Element: {elementId}, Wall: {(wallId > 0 ? wallId.ToString() : "unknown")}");
            AppendToLog(DebugLogName, $"  Location: {FormatXYZ(location)}");
            if (wallOrientation != null)
            {
                AppendToLog(DebugLogName, $"  Wall orientation: {FormatXYZ(wallOrientation)}");
            }
            AppendToLog(DebugLogName, $"  Pipe diameter: {FormatMM(diameter)}mm");
        }

        /// <summary>
        /// Log a successful pipe sleeve placement
        /// </summary>
        public static void LogPipeSleeveSuccess(int elementId, int sleeveId, double sleeveDiameter, XYZ sleevePosition)
        {
            if (!DebugLogger.IsEnabled) return;
            _totalPipeSleevesPLaced++;

            string message = $"SUCCESS: Placed sleeve {sleeveId} for pipe {elementId}, sleeve diameter: {FormatMM(sleeveDiameter)}mm";
            AppendToLog(PipeLogName, message);

            // More detailed info in debug log
            AppendToLog(DebugLogName, $"PIPE SLEEVE PLACED - Pipe: {elementId}, Sleeve: {sleeveId}");
            AppendToLog(DebugLogName, $"  Final position: {FormatXYZ(sleevePosition)}");
            AppendToLog(DebugLogName, $"  Sleeve diameter: {FormatMM(sleeveDiameter)}mm");
        }

        /// <summary>
        /// Log a failed pipe sleeve placement
        /// </summary>
        public static void LogPipeSleeveFailure(int elementId, string reason)
        {
            if (!DebugLogger.IsEnabled) return;
            _processedElements.Add(elementId);
            string message = $"FAILED: Could not place sleeve for pipe {elementId}. Reason: {reason}";
            AppendToLog(PipeLogName, message);
            _missingPipes.Add($"Pipe {elementId}: {reason}");

            // More detailed info in debug log
            AppendToLog(DebugLogName, $"PIPE SLEEVE FAILED - Element: {elementId}");
            AppendToLog(DebugLogName, $"  Reason: {reason}");
        }

        /// <summary>
        /// Log a found duct wall intersection
        /// </summary>
        public static void LogDuctWallIntersection(int elementId, XYZ location, double width, double height, int wallId = 0, XYZ? wallOrientation = null)
        {
            if (!DebugLogger.IsEnabled) return;
            _totalDuctsWithWalls++;
            _processedElements.Add(elementId);
            string wallInfo = wallId > 0 ? $"wall {wallId}" : "wall";
            string message = $"Found duct {elementId} intersecting {wallInfo} at {FormatXYZ(location)}, size: {FormatMM(width)}mm × {FormatMM(height)}mm";
            AppendToLog(DuctLogName, message);

            // More detailed info in debug log
            AppendToLog(DebugLogName, $"DUCT-WALL INTERSECTION - Element: {elementId}, Wall: {(wallId > 0 ? wallId.ToString() : "unknown")}");
            AppendToLog(DebugLogName, $"  Location: {FormatXYZ(location)}");
            if (wallOrientation != null)
            {
                AppendToLog(DebugLogName, $"  Wall orientation: {FormatXYZ(wallOrientation)}");
            }
            AppendToLog(DebugLogName, $"  Duct size: {FormatMM(width)}mm × {FormatMM(height)}mm");
        }

        /// <summary>
        /// Log a successful duct sleeve placement
        /// </summary>
        public static void LogDuctSleeveSuccess(int elementId, int sleeveId, double sleeveWidth, double sleeveHeight, XYZ sleevePosition)
        {
            if (!DebugLogger.IsEnabled) return;
            _totalDuctSleevesPLaced++;

            string message = $"SUCCESS: Placed sleeve {sleeveId} for duct {elementId}, sleeve size: {FormatMM(sleeveWidth)}mm × {FormatMM(sleeveHeight)}mm";
            AppendToLog(DuctLogName, message);

            // More detailed info in debug log
            AppendToLog(DebugLogName, $"DUCT SLEEVE PLACED - Duct: {elementId}, Sleeve: {sleeveId}");
            AppendToLog(DebugLogName, $"  Final position: {FormatXYZ(sleevePosition)}");
            AppendToLog(DebugLogName, $"  Sleeve size: {FormatMM(sleeveWidth)}mm × {FormatMM(sleeveHeight)}mm");
        }

        /// <summary>
        /// Log a failed duct sleeve placement
        /// </summary>
        public static void LogDuctSleeveFailure(int elementId, string reason)
        {
            if (!DebugLogger.IsEnabled) return;
            _processedElements.Add(elementId);
            string message = $"FAILED: Could not place sleeve for duct {elementId}. Reason: {reason}";
            AppendToLog(DuctLogName, message);
            _missingDucts.Add($"Duct {elementId}: {reason}");

            // More detailed info in debug log
            AppendToLog(DebugLogName, $"DUCT SLEEVE FAILED - Element: {elementId}");
            AppendToLog(DebugLogName, $"  Reason: {reason}");
        }

        /// <summary>
        /// Log a found cable tray wall intersection
        /// </summary>
        public static void LogCableTrayWallIntersection(int elementId, XYZ location, double width, double height, int wallId = 0, XYZ? wallOrientation = null)
        {
            if (!DebugLogger.IsEnabled) return;
            _totalCableTraysWithWalls++;
            _processedElements.Add(elementId);

            string wallInfo = wallId > 0 ? $"wall {wallId}" : "wall";
            string message = $"Found cable tray {elementId} intersecting {wallInfo} at {FormatXYZ(location)}, size: {FormatMM(width)}mm × {FormatMM(height)}mm";
            AppendToLog(CableTrayLogName, message);

            // More detailed info in debug log
            AppendToLog(DebugLogName, $"CABLETRAY-WALL INTERSECTION - Element: {elementId}, Wall: {(wallId > 0 ? wallId.ToString() : "unknown")}");
            AppendToLog(DebugLogName, $"  Location: {FormatXYZ(location)}");
            if (wallOrientation != null)
            {
                AppendToLog(DebugLogName, $"  Wall orientation: {FormatXYZ(wallOrientation)}");
            }
            AppendToLog(DebugLogName, $"  Cable tray size: {FormatMM(width)}mm × {FormatMM(height)}mm");
        }

        /// <summary>
        /// Log a successful cable tray sleeve placement
        /// </summary>
        public static void LogCableTraySleeveSuccess(int elementId, int sleeveId, double sleeveWidth, double sleeveHeight, XYZ sleevePosition)
        {
            if (!DebugLogger.IsEnabled) return;
            _totalCableTraySleevesPLaced++;

            string message = $"SUCCESS: Placed sleeve {sleeveId} for cable tray {elementId}, sleeve size: {FormatMM(sleeveWidth)}mm × {FormatMM(sleeveHeight)}mm";
            AppendToLog(CableTrayLogName, message);

            // More detailed info in debug log
            AppendToLog(DebugLogName, $"CABLETRAY SLEEVE PLACED - Cable Tray: {elementId}, Sleeve: {sleeveId}");
            AppendToLog(DebugLogName, $"  Final position: {FormatXYZ(sleevePosition)}");
            AppendToLog(DebugLogName, $"  Sleeve size: {FormatMM(sleeveWidth)}mm × {FormatMM(sleeveHeight)}mm");
        }

        /// <summary>
        /// Log a failed cable tray sleeve placement
        /// </summary>
        public static void LogCableTraySleeveFailure(int elementId, string reason)
        {
            if (!DebugLogger.IsEnabled) return;
            _processedElements.Add(elementId);
            string message = $"FAILED: Could not place sleeve for cable tray {elementId}. Reason: {reason}";
            AppendToLog(CableTrayLogName, message);
            _missingCableTrays.Add($"Cable Tray {elementId}: {reason}");

            // More detailed info in debug log
            AppendToLog(DebugLogName, $"CABLETRAY SLEEVE FAILED - Element: {elementId}");
            AppendToLog(DebugLogName, $"  Reason: {reason}");
        }

        /// <summary>
        /// Log info-level messages for cable tray sleeve parameter actions (e.g., HostOrientation)
        /// </summary>
        public static void LogCableTraySleeveInfo(int elementId, string info)
        {
            if (!DebugLogger.IsEnabled) return;
            AppendToLog(CableTrayLogName, info);
            // Also log to debug for traceability
            AppendToLog(DebugLogName, $"CABLETRAY SLEEVE INFO - Element: {elementId}");
            AppendToLog(DebugLogName, $"  Info: {info}");
        }

        /// <summary>
        /// Log a warning message to all logs
        /// </summary>
        public static void LogWarning(string message)
        {
            if (!DebugLogger.IsEnabled) return;
            _warnings.Add(message);
            string formattedMessage = $"WARNING: {message}";
            AppendToLog(SummaryLogName, formattedMessage);
            AppendToLog(PipeLogName, formattedMessage);
            AppendToLog(DuctLogName, formattedMessage);
            AppendToLog(CableTrayLogName, formattedMessage);
            AppendToLog(DebugLogName, formattedMessage);
        }

        /// <summary>
        /// Log a pipe inspection message for terminal connections or fitting issues
        /// </summary>
        public static void LogPipeInspection(int elementId, string message)
        {
            if (!DebugLogger.IsEnabled) return;
            _processedElements.Add(elementId);
            string logMessage = $"INSPECT: Pipe {elementId} - {message}";
            AppendToLog(PipeLogName, logMessage);

            // More detailed info in debug log
            AppendToLog(DebugLogName, $"PIPE INSPECTION - Element: {elementId}");
            AppendToLog(DebugLogName, $"  Note: {message}");
        }

        /// <summary>
        /// Log a duct inspection message for terminal connections or fitting issues
        /// </summary>
        public static void LogDuctInspection(int elementId, string message)
        {
            if (!DebugLogger.IsEnabled) return;
            _processedElements.Add(elementId);
            string logMessage = $"INSPECT: Duct {elementId} - {message}";
            AppendToLog(DuctLogName, logMessage);

            // More detailed info in debug log
            AppendToLog(DebugLogName, $"DUCT INSPECTION - Element: {elementId}");
            AppendToLog(DebugLogName, $"  Note: {message}");
        }

        /// <summary>
        /// Log a detailed geometric analysis message about element-wall proximity
        /// </summary>
        public static void LogGeometryAnalysis(string elementType, int elementId, int wallId, double distance, XYZ elementPoint, XYZ wallPoint)
        {
            if (!DebugLogger.IsEnabled) return;
            string message = $"GEOMETRY: {elementType} {elementId} is {FormatMM(distance)}mm from wall {wallId}";

            // Log to both element-specific and debug logs
            if (elementType.Equals("Pipe", StringComparison.OrdinalIgnoreCase))
            {
                AppendToLog(PipeLogName, message);
            }
            else if (elementType.Equals("Duct", StringComparison.OrdinalIgnoreCase))
            {
                AppendToLog(DuctLogName, message);
            }
            else if (elementType.Equals("CableTray", StringComparison.OrdinalIgnoreCase))
            {
                AppendToLog(CableTrayLogName, message);
            }

            // Detailed debug info
            AppendToLog(DebugLogName, $"WALL PROXIMITY - {elementType}: {elementId}, Wall: {wallId}");
            AppendToLog(DebugLogName, $"  Distance: {FormatMM(distance)}mm");
            AppendToLog(DebugLogName, $"  Element point: {FormatXYZ(elementPoint)}");
            AppendToLog(DebugLogName, $"  Wall point: {FormatXYZ(wallPoint)}");
        }

        /// <summary>
        /// Log diagnostic information to debug log only
        /// </summary>
        public static void LogDebug(string message)
        {
            if (!DebugLogger.IsEnabled) return;
            AppendToLog(DebugLogName, $"DEBUG: {message}");
        }

        /// <summary>
        /// Finalize logs with summary statistics and identify any completely missing elements
        /// </summary>
        public static void FinalizeLogs()
        {
            if (!DebugLogger.IsEnabled) return;
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
                AppendToLog(SummaryLogName, summaryText);
                AppendToLog(PipeLogName, summaryText);
                AppendToLog(DuctLogName, summaryText);
                AppendToLog(CableTrayLogName, summaryText);
                AppendToLog(DebugLogName, summaryText);
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
                    AppendToLog(PipeLogName, $"MISSED: {message}");
                    AppendToLog(DebugLogName, $"UNPROCESSED ELEMENT: {message}");
                }
            }

            // Check for completely unprocessed ducts
            foreach (var ductId in _allModelDucts)
            {
                if (!_processedElements.Contains(ductId))
                {
                    string message = $"Duct {ductId}: COMPLETELY MISSED - No wall intersection detected";
                    _missingDucts.Add(message);
                    AppendToLog(DuctLogName, $"MISSED: {message}");
                    AppendToLog(DebugLogName, $"UNPROCESSED ELEMENT: {message}");
                }
            }

            // Check for completely unprocessed cable trays
            foreach (var trayId in _allModelCableTrays)
            {
                if (!_processedElements.Contains(trayId))
                {
                    string message = $"Cable Tray {trayId}: COMPLETELY MISSED - No wall intersection detected";
                    _missingCableTrays.Add(message);
                    AppendToLog(CableTrayLogName, $"MISSED: {message}");
                    AppendToLog(DebugLogName, $"UNPROCESSED ELEMENT: {message}");
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

        private static void AppendToLog(string logName, string message)
        {
            try
            {
                // ✅ PERFORMANCE: Consolidate to SafeFileLogger
                SafeFileLogger.SafeAppendText(logName, message);
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"Failed to write to log {logName}: {ex.Message}");
            }
        }
    }
}
