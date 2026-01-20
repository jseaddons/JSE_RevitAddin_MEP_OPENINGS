using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Switch;
using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// ✅ DIAGNOSTIC STATUS COMMAND: Shows current diagnostic mode status
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class DiagnosticStatusCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Get current status
                bool diagnosticLogging = MasterSwitch.DiagnosticLogging;
                bool deploymentMode = DeploymentConfiguration.DeploymentMode;
                bool useDiagnosticMode = OptimizationFlags.UseDiagnosticMode;
                bool logPerformanceMetrics = OptimizationFlags.LogPerformanceMetrics;
                
                string status = $@"🔌 DIAGNOSTIC LOGGING STATUS

════════════════════════════════════════
CURRENT STATE: {(diagnosticLogging ? "🔴 ON (SLOW MODE)" : "🟢 OFF (FAST MODE)")}
════════════════════════════════════════

MasterSwitch.DiagnosticLogging: {(diagnosticLogging ? "✅ TRUE (ON)" : "❌ FALSE (OFF)")}
DeploymentConfiguration.DeploymentMode: {(deploymentMode ? "❌ TRUE (Minimal Logging)" : "✅ FALSE (Full Logging)")}
OptimizationFlags.UseDiagnosticMode: {(useDiagnosticMode ? "✅ TRUE" : "❌ FALSE")}
OptimizationFlags.LogPerformanceMetrics: {(logPerformanceMetrics ? "✅ TRUE" : "❌ FALSE")}

════════════════════════════════════════
PERFORMANCE:
════════════════════════════════════════
{(diagnosticLogging ? "🔴 SLOW MODE: ~9 zones/sec (273+ logging calls)" : "🟢 FAST MODE: ~30 zones/sec (minimal logging)")}

⚠️ WARNING: Diagnostic mode ON = 3x slower
Use diagnostic mode ONLY when debugging.

════════════════════════════════════════
HOW TO CHANGE:
════════════════════════════════════════
1. Click 'Toggle Diagnostic' button in ribbon
2. Or set in code: MasterSwitch.DiagnosticLogging = true/false
3. Or set in Application.cs: MasterSwitch.DiagnosticLogging = true/false";
                
                TaskDialog.Show("Diagnostic Status", status);
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error checking diagnostic status: {ex.Message}";
                return Result.Failed;
            }
        }
    }
}

