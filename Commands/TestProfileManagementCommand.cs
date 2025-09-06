using System;
using System.IO;
using System.Windows;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.ViewModels;
using JSE_RevitAddin_MEP_OPENINGS.Views;
using JSE_RevitAddin_MEP_OPENINGS.Commands;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// Test command to demonstrate the profile management system
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class TestProfileManagementCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // IMMEDIATE LOGGING - Create timestamped file as soon as command starts
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            string logPath = $@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\command_started_{timestamp}.log";
            try
            {
                File.AppendAllText(logPath, $"[{DateTime.Now}] TestProfileManagementCommand.Execute() STARTED (DIRECT ADD-IN LOAD)\n");
                File.AppendAllText(logPath, $"[{DateTime.Now}] Log file: command_started_{timestamp}.log\n");
                File.AppendAllText(logPath, $"[{DateTime.Now}] commandData: {commandData != null}\n");
                File.AppendAllText(logPath, $"[{DateTime.Now}] commandData.Application: {commandData?.Application != null}\n");
                File.AppendAllText(logPath, $"[{DateTime.Now}] Loading method: DIRECT ADD-IN MANAGER\n");
            }
            catch (Exception ex)
            {
                // If even this fails, try a different location with timestamp
                try
                {
                    File.AppendAllText($@"C:\temp\revit_command_{timestamp}.log", $"[{DateTime.Now}] Command started but main log failed: {ex.Message}\n");
                }
                catch { }
                return Result.Failed;
            }

            try
            {
                var uiApp = commandData.Application;
                var uiDoc = uiApp.ActiveUIDocument;
                var doc = uiDoc.Document;

                // Get the application profile service
                var appProfileService = ApplicationProfileService.Instance;

                // CRITICAL: Update the profile service for the current document BEFORE checking IsProfileSetupRequired
                // This ensures we're looking at the correct project-specific profiles
                System.Diagnostics.Debug.WriteLine("About to call UpdateForCurrentDocument");
                File.AppendAllText(logPath, $"[{DateTime.Now}] About to call UpdateForCurrentDocument\n");
                appProfileService.UpdateForCurrentDocument(doc.PathName);
                System.Diagnostics.Debug.WriteLine("UpdateForCurrentDocument call completed");
                File.AppendAllText(logPath, $"[{DateTime.Now}] UpdateForCurrentDocument call completed\n");

                // Add basic logging to see which path we take
                System.Diagnostics.Debug.WriteLine("=== TestProfileManagementCommand STARTED ===");
                System.Diagnostics.Debug.WriteLine($"Document path: {doc.PathName}");
                System.Diagnostics.Debug.WriteLine($"IsProfileSetupRequired: {appProfileService.IsProfileSetupRequired}");

                // Log to file as well
                File.AppendAllText(logPath, $"[{DateTime.Now}] ApplicationProfileService.Instance created\n");
                File.AppendAllText(logPath, $"[{DateTime.Now}] Document path: {doc.PathName}\n");
                File.AppendAllText(logPath, $"[{DateTime.Now}] IsProfileSetupRequired: {appProfileService.IsProfileSetupRequired}\n");

                // Note: Removed WPF Application creation to prevent "Cannot create more than one System.Windows.Application instance" error

                // STEP 3: Test complete flow - Profile Management then Main Dialog
                System.Diagnostics.Debug.WriteLine("Testing complete flow - Profile Management then Main Dialog");
                File.AppendAllText(logPath, $"[{DateTime.Now}] Testing complete flow - Profile Management then Main Dialog\n");
                
                // First show profile management dialog
                using (var profileMgmtDialog = new Views.EmergencyProfileManagementDialog(appProfileService))
                {
                    var result = profileMgmtDialog.ShowDialog();
                    System.Diagnostics.Debug.WriteLine($"Profile management dialog result: {result}");
                    File.AppendAllText(logPath, $"[{DateTime.Now}] Profile management dialog result: {result}\n");
                    
                    if (result == System.Windows.Forms.DialogResult.OK && profileMgmtDialog.ShouldOpenMainDialog)
                    {
                        // Then show main dialog
                        System.Diagnostics.Debug.WriteLine("Opening main dialog after profile management");
                        File.AppendAllText(logPath, $"[{DateTime.Now}] Opening main dialog after profile management\n");
                        
                        using (var mainDialog = new Views.EmergencyMainDialog(appProfileService, doc))
                        {
                            var mainResult = mainDialog.ShowDialog();
                            System.Diagnostics.Debug.WriteLine($"Main dialog result: {mainResult}");
                            File.AppendAllText(logPath, $"[{DateTime.Now}] Main dialog result: {mainResult}\n");
                        }
                    }
                }
                
                // Log success
                File.AppendAllText(logPath, $"[{DateTime.Now}] Complete flow tested successfully\n");
                
                // Note: Not cleaning up singleton to avoid potential issues
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                // Show error in TaskDialog instead of crashing
                try
                {
                    TaskDialog.Show("Error", $"Command failed but Revit did not crash:\n\n{ex.Message}");
                }
                catch
                {
                    // If even TaskDialog fails, just set the message
                    message = $"Error: {ex.Message}";
                }
                
                return Result.Failed;
            }
        }

        /// <summary>
        /// Simple test to directly create EmergencyMainDialog
        /// </summary>
        public static void TestDirectEmergencyMainDialog(Document doc)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("=== TestDirectEmergencyMainDialog STARTED ===");
                var appProfileService = ApplicationProfileService.Instance;
                System.Diagnostics.Debug.WriteLine("Creating EmergencyMainDialog directly...");
                using (var emergencyMainDlg = new JSE_RevitAddin_MEP_OPENINGS.Views.EmergencyMainDialog(appProfileService, doc))
                {
                    emergencyMainDlg.ShowDialog(); // Modal dialog with proper disposal
                }
                System.Diagnostics.Debug.WriteLine("=== TestDirectEmergencyMainDialog COMPLETED ===");
                
                // Clean up the singleton instance
                ApplicationProfileService.CleanupInstance();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"TestDirectEmergencyMainDialog ERROR: {ex.Message}");
                
                // Show error in TaskDialog instead of crashing
                try
                {
                    var errorDialog = new TaskDialog("Static Test Error");
                    errorDialog.MainInstruction = "Static Test Failed";
                    errorDialog.MainContent = $"An error occurred in static test but Revit did not crash:\n\n{ex.Message}";
                    errorDialog.CommonButtons = TaskDialogCommonButtons.Ok;
                    errorDialog.Show();
                }
                catch { }
                
                // Clean up even if there was an error
                try
                {
                    ApplicationProfileService.CleanupInstance();
                }
                catch { }
            }
        }

        /// <summary>
        /// Static test method that can be called directly from add-in manager
        /// </summary>
        public static void StaticTest()
        {
            try
            {
                string testLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\static_test.log";
                File.AppendAllText(testLogPath, $"[{DateTime.Now}] StaticTest() called from add-in manager\n");
                File.AppendAllText(testLogPath, $"[{DateTime.Now}] Assembly loaded successfully\n");
            }
            catch (Exception ex)
            {
                try
                {
                    File.AppendAllText(@"C:\temp\static_test.log", $"[{DateTime.Now}] StaticTest failed: {ex.Message}\n");
                }
                catch { }
            }
        }
    }
}
