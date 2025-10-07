using System;
using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Immutable value object containing clearance settings from UI
    /// Thread-safe and unit-testable
    /// </summary>
    public class ClearanceValues
    {
        public double DuctsNormalClearance { get; }
        public double DuctsInsulatedClearance { get; }
        public double PipesNormalClearance { get; }
        public double PipesInsulatedClearance { get; }
        public double CableTrayTopClearance { get; }
        public double CableTrayOtherClearance { get; }
        public double FireDamperStandardClearance { get; }
        public double FireDamperMsfdClearance { get; }

        public ClearanceValues(
            double ductsNormal = 50.0,
            double ductsInsulated = 25.0,
            double pipesNormal = 50.0,
            double pipesInsulated = 25.0,
            double cableTrayTop = 75.0,
            double cableTrayOther = 25.0,
            double fireDamperStandard = 50.0,
            double fireDamperMsfd = 100.0)
        {
            DuctsNormalClearance = ductsNormal;
            DuctsInsulatedClearance = ductsInsulated;
            PipesNormalClearance = pipesNormal;
            PipesInsulatedClearance = pipesInsulated;
            CableTrayTopClearance = cableTrayTop;
            CableTrayOtherClearance = cableTrayOther;
            FireDamperStandardClearance = fireDamperStandard;
            FireDamperMsfdClearance = fireDamperMsfd;
        }

        /// <summary>
        /// Create ClearanceValues from UI dictionary
        /// </summary>
        public static ClearanceValues FromDictionary(Dictionary<string, double> uiClearances)
        {
            DebugLogger.Info($"[CLEARANCE_DEBUG] ClearanceValues.FromDictionary called with {uiClearances?.Count ?? 0} values:");
            if (uiClearances != null)
            {
                foreach (var kvp in uiClearances)
                {
                    DebugLogger.Info($"[CLEARANCE_DEBUG]   {kvp.Key} = {kvp.Value}mm");
                }
            }
            
            if (uiClearances == null || uiClearances.Count == 0)
            {
                DebugLogger.Info($"[CLEARANCE_DEBUG] Using default ClearanceValues (no UI clearances provided)");
                return new ClearanceValues(); // Use defaults
            }

            var ductsNormal = GetValueOrDefault(uiClearances, "ducts_normal_clearance", 50.0);
            var ductsInsulated = GetValueOrDefault(uiClearances, "ducts_insulated_clearance", 25.0);
            var pipesNormal = GetValueOrDefault(uiClearances, "pipes_normal_clearance", 50.0);
            var pipesInsulated = GetValueOrDefault(uiClearances, "pipes_insulated_clearance", 25.0);
            var cableTrayTop = GetValueOrDefault(uiClearances, "cabletray_top_normal", 75.0);
            var cableTrayOther = GetValueOrDefault(uiClearances, "cabletray_other_normal", 25.0);
            var fireDamperStandard = GetValueOrDefault(uiClearances, "fire_damper_standard_clearance", 50.0);
            var fireDamperMsfd = GetValueOrDefault(uiClearances, "fire_damper_msfd_clearance", 100.0);
            
            DebugLogger.Info($"[CLEARANCE_DEBUG] Extracted clearance values:");
            DebugLogger.Info($"[CLEARANCE_DEBUG]   ducts_normal_clearance: {ductsNormal}mm");
            DebugLogger.Info($"[CLEARANCE_DEBUG]   ducts_insulated_clearance: {ductsInsulated}mm");
            DebugLogger.Info($"[CLEARANCE_DEBUG]   pipes_normal_clearance: {pipesNormal}mm");
            DebugLogger.Info($"[CLEARANCE_DEBUG]   pipes_insulated_clearance: {pipesInsulated}mm");
            DebugLogger.Info($"[CLEARANCE_DEBUG]   cabletray_top_normal: {cableTrayTop}mm");
            DebugLogger.Info($"[CLEARANCE_DEBUG]   cabletray_other_normal: {cableTrayOther}mm");
            DebugLogger.Info($"[CLEARANCE_DEBUG]   fire_damper_standard_clearance: {fireDamperStandard}mm");
            DebugLogger.Info($"[CLEARANCE_DEBUG]   fire_damper_msfd_clearance: {fireDamperMsfd}mm");

            return new ClearanceValues(
                ductsNormal: ductsNormal,
                ductsInsulated: ductsInsulated,
                pipesNormal: pipesNormal,
                pipesInsulated: pipesInsulated,
                cableTrayTop: cableTrayTop,
                cableTrayOther: cableTrayOther,
                fireDamperStandard: fireDamperStandard,
                fireDamperMsfd: fireDamperMsfd
            );
        }

        private static double GetValueOrDefault(Dictionary<string, double> dict, string key, double defaultValue)
        {
            return dict.TryGetValue(key, out double value) ? value : defaultValue;
        }

        /// <summary>
        /// Get clearance for ducts based on insulation status
        /// </summary>
        public double GetDuctClearance(bool isInsulated)
        {
            double clearance = isInsulated ? DuctsInsulatedClearance : DuctsNormalClearance;
            DebugLogger.Info($"[CLEARANCE_DEBUG] ClearanceValues.GetDuctClearance(isInsulated={isInsulated}) = {clearance}mm");
            return clearance;
        }

        /// <summary>
        /// Get clearance for pipes based on insulation status
        /// </summary>
        public double GetPipeClearance(bool isInsulated)
        {
            return isInsulated ? PipesInsulatedClearance : PipesNormalClearance;
        }

        /// <summary>
        /// Get clearance for cable trays based on side
        /// </summary>
        public double GetCableTrayClearance(string side)
        {
            return side?.ToLower() switch
            {
                "top" => CableTrayTopClearance,
                _ => CableTrayOtherClearance
            };
        }

        /// <summary>
        /// Get clearance for fire dampers based on type
        /// </summary>
        public double GetFireDamperClearance(bool isMsfd)
        {
            return isMsfd ? FireDamperMsfdClearance : FireDamperStandardClearance;
        }

        /// <summary>
        /// Get clearance for cable tray based on side
        /// </summary>
        public double GetCableTrayClearance(bool isTopSide = false)
        {
            return isTopSide ? CableTrayTopClearance : CableTrayOtherClearance;
        }

        public override string ToString()
        {
            return $"ClearanceValues(Ducts: {DuctsNormalClearance}/{DuctsInsulatedClearance}mm, " +
                   $"Pipes: {PipesNormalClearance}/{PipesInsulatedClearance}mm, " +
                   $"CableTray: {CableTrayTopClearance}/{CableTrayOtherClearance}mm, " +
                   $"FireDamper: {FireDamperStandardClearance}/{FireDamperMsfdClearance}mm)";
        }
    }
}
