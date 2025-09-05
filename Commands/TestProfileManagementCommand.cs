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

                // Ensure WPF Application exists for Revit add-in compatibility
                if (System.Windows.Application.Current == null)
                {
                    new System.Windows.Application();
                }

                // Check if profile setup is required
                if (appProfileService.IsProfileSetupRequired)
                {
                    System.Diagnostics.Debug.WriteLine("Taking Profile Setup branch");
                    File.AppendAllText(logPath, $"[{DateTime.Now}] Taking Profile Setup branch\n");
                    
                    // EMERGENCY MODE: Use WinForms ProfileSetup instead of WPF to avoid crashes
                    var emergencyDialog = new Views.EmergencyProfileSetup(appProfileService.ProfileService, appProfileService.StatusManager);

                    var result = emergencyDialog.ShowDialog();
                    if (result == System.Windows.Forms.DialogResult.OK && emergencyDialog.CreatedProfile != null)
                    {
                        // Set the created profile as current
                        appProfileService.SetCurrentProfile(emergencyDialog.CreatedProfile);
                        
                        // Debug: Check if profiles are now available
                        System.Diagnostics.Debug.WriteLine($"After profile creation - IsProfileSetupRequired: {appProfileService.IsProfileSetupRequired}");
                        File.AppendAllText(logPath, $"[{DateTime.Now}] After profile creation - IsProfileSetupRequired: {appProfileService.IsProfileSetupRequired}\n");

                        // Show success message
                        TaskDialog.Show("Profile Setup", $"Profile '{emergencyDialog.CreatedProfile.Name}' created successfully!");

                        // EMERGENCY MODE: Open WinForms main dialog (crash-safe)
                        System.Diagnostics.Debug.WriteLine("Creating EmergencyMainDialog after profile setup");
                        var emergencyMainDlg = new JSE_RevitAddin_MEP_OPENINGS.Views.EmergencyMainDialog(appProfileService, doc);
                        emergencyMainDlg.Show(); // Modeless dialog
                    }
                    else
                    {
                        TaskDialog.Show("Profile Setup", "Profile setup was cancelled.");
                        return Result.Cancelled;
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("Taking Profile Management branch");
                    File.AppendAllText(logPath, $"[{DateTime.Now}] Taking Profile Management branch\n");
                    
                    // EMERGENCY MODE: Use WinForms ProfileManagement instead of WPF to prevent crashes
                    var emergencyProfileMgmt = new Views.EmergencyProfileManagementDialog(appProfileService);

                    var result = emergencyProfileMgmt.ShowDialog();
                    System.Diagnostics.Debug.WriteLine($"Profile management dialog result: {result}");
                    System.Diagnostics.Debug.WriteLine($"ShouldOpenMainDialog: {emergencyProfileMgmt.ShouldOpenMainDialog}");

                    if (result == System.Windows.Forms.DialogResult.OK && emergencyProfileMgmt.ShouldOpenMainDialog)
                    {
                        // Open WinForms main dialog after profile action
                        System.Diagnostics.Debug.WriteLine("Creating EmergencyMainDialog after profile management");
                        var emergencyMainDlg = new Views.EmergencyMainDialog(appProfileService, doc);
                        emergencyMainDlg.Show(); // Modeless
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("EmergencyMainDialog NOT created - conditions not met");
                    }
                }
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error: {ex.Message}";
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
                var emergencyMainDlg = new JSE_RevitAddin_MEP_OPENINGS.Views.EmergencyMainDialog(appProfileService, doc);
                emergencyMainDlg.Show();
                System.Diagnostics.Debug.WriteLine("=== TestDirectEmergencyMainDialog COMPLETED ===");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"TestDirectEmergencyMainDialog ERROR: {ex.Message}");
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
