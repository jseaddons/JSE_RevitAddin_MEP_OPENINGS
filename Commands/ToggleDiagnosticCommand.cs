using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Switch;
using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// ✅ TOGGLE DIAGNOSTIC COMMAND: Simple command to toggle diagnostic logging on/off
    /// 
    /// USAGE:
    ///   1. Add this command to your ribbon
    ///   2. Click the button to toggle diagnostic logging
    ///   3. A message box will show the current status
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ToggleDiagnosticCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Get status BEFORE toggle
                bool wasEnabled = MasterSwitch.DiagnosticLogging;
                bool deploymentModeBefore = DeploymentConfiguration.DeploymentMode;
                bool useDiagnosticModeBefore = OptimizationFlags.UseDiagnosticMode;
                
                // Toggle diagnostic logging
                MasterSwitch.DiagnosticLogging = !MasterSwitch.DiagnosticLogging;
                
                // Get status AFTER toggle
                bool isEnabled = MasterSwitch.DiagnosticLogging;
                bool deploymentModeAfter = DeploymentConfiguration.DeploymentMode;
                bool useDiagnosticModeAfter = OptimizationFlags.UseDiagnosticMode;
                
                         // Show detailed status to user
                         string status = $@"🔌 DIAGNOSTIC LOGGING TOGGLE

════════════════════════════════════════
BEFORE:
════════════════════════════════════════
DiagnosticLogging: {(wasEnabled ? "✅ ON" : "❌ OFF")}
DeploymentMode: {(deploymentModeBefore ? "❌ ON (Minimal Logging)" : "✅ OFF (Full Logging)")}
UseDiagnosticMode: {(useDiagnosticModeBefore ? "✅ ON" : "❌ OFF")}

════════════════════════════════════════
AFTER TOGGLE:
════════════════════════════════════════
DiagnosticLogging: {(isEnabled ? "✅ ON" : "❌ OFF")}
DeploymentMode: {(deploymentModeAfter ? "❌ ON (Minimal Logging)" : "✅ OFF (Full Logging)")}
UseDiagnosticMode: {(useDiagnosticModeAfter ? "✅ ON" : "❌ OFF")}

════════════════════════════════════════
CURRENT MODE:
════════════════════════════════════════
{(isEnabled ? "🔴 SLOW MODE: ~9 zones/sec (273+ logging calls)\n⚠️ WARNING: Use only for debugging!" : "🟢 FAST MODE: ~30 zones/sec (minimal logging)")}

════════════════════════════════════════
TO CHECK STATUS:
════════════════════════════════════════
Click 'Diagnostic Status' button in ribbon";
                
                TaskDialog.Show("Diagnostic Switch Toggled", status);
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error toggling diagnostic mode: {ex.Message}";
                return Result.Failed;
            }
        }
    }
}

