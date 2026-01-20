using Autodesk.Revit.DB;
using System;
using System.IO;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.DamperDetection
{
    /// <summary>
    /// Orchestrates damper connector detection using type detection and connector detection services.
    /// Follows Dependency Inversion Principle - depends on abstractions (interfaces).
    /// Follows Single Responsibility Principle - coordinates detection, doesn't implement it.
    /// </summary>
    public class DamperConnectorService
    {
        private readonly IDamperTypeDetector _typeDetector;
        private readonly IDamperConnectorDetector _connectorDetector;

        /// <summary>
        /// Creates a new instance with default detectors.
        /// </summary>
        public DamperConnectorService()
            : this(new DamperTypeDetector(), new DamperConnectorDetector())
        {
        }

        /// <summary>
        /// Creates a new instance with injected dependencies (for testing).
        /// </summary>
        public DamperConnectorService(IDamperTypeDetector typeDetector, IDamperConnectorDetector connectorDetector)
        {
            _typeDetector = typeDetector ?? throw new System.ArgumentNullException(nameof(typeDetector));
            _connectorDetector = connectorDetector ?? throw new System.ArgumentNullException(nameof(connectorDetector));
        }

        /// <summary>
        /// Detects complete connector information for a damper element.
        /// Returns DamperConnectorInfo.None if the element is not a damper or has no connector.
        /// </summary>
        /// <param name="wallOrientation">Wall orientation ("X" or "Y") to prioritize wall width axis during detection</param>
        public DamperConnectorInfo DetectConnectorInfo(Element mepElement, string mepCategory, string wallOrientation = null)
        {
            // Only process duct accessories (dampers)
            if (mepCategory != "Duct Accessories")
            {
                return DamperConnectorInfo.None;
            }

            var damper = mepElement as FamilyInstance;
            if (damper == null)
            {
                return DamperConnectorInfo.None;
            }

            // Step 1: Detect damper type from family/type name
            string familyTypeName = damper.Symbol?.Name ?? "";
            string damperType = _typeDetector.DetectDamperType(familyTypeName);
            bool isStandard = _typeDetector.IsStandardDamper(damperType);
            bool requiresMepClearance = _typeDetector.RequiresMepSideClearance(damperType);

            // Step 2: Check if damper has MEP connector
            bool hasMepConnector = _connectorDetector.HasMepConnector(damper);
            string connectorSide = string.Empty;

            // Step 3: If connector exists and requires MEP clearance, detect its direction
            if (hasMepConnector && requiresMepClearance)
            {
                // Non-standard dampers (MSFD, MSD, MD, Motorized) use world coordinates
                bool useWorldCoordinates = requiresMepClearance;
                // ✅ DIAGNOSTIC: Log wall orientation being passed
                SafeFileLogger.SafeAppendText("damper_connector_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [DamperConnectorService] Calling DetectConnectorSide with wallOrientation='{wallOrientation ?? "NULL"}' for damper {damper?.Id?.IntegerValue ?? -1}\n");
                // ✅ BUILD STAMP at service level for traceability
                try
                {
                    var asm = typeof(DamperConnectorService).Assembly;
                    string asmLoc = asm.Location;
                    DateTime asmWrite = File.Exists(asmLoc) ? File.GetLastWriteTime(asmLoc) : DateTime.MinValue;
                    string asmVer = asm.GetName().Version?.ToString() ?? "unknown";
                    SafeFileLogger.SafeAppendText("damper_connector_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [BUILD-STAMP:SERVICE] AssemblyLastWrite={asmWrite:yyyy-MM-dd HH:mm:ss}, Version={asmVer}, Assembly={asmLoc}\n");
                }
                catch { }
                connectorSide = _connectorDetector.DetectConnectorSide(damper, useWorldCoordinates, out _, wallOrientation);
            }
            else if (hasMepConnector && isStandard)
            {
                // Standard dampers use local coordinates
                connectorSide = _connectorDetector.DetectConnectorSide(damper, useWorldCoordinates: false, out _);
            }

            return new DamperConnectorInfo(
                hasMepConnector: hasMepConnector,
                connectorSide: connectorSide,
                isStandardDamper: isStandard,
                damperType: damperType,
                requiresMepSideClearance: requiresMepClearance
            );
        }
    }
}
