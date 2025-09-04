using System;
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
            try
            {
                var uiApp = commandData.Application;
                var uiDoc = uiApp.ActiveUIDocument;
                var doc = uiDoc.Document;

                // Get the application profile service
                var appProfileService = ApplicationProfileService.Instance;

                // Ensure WPF Application exists for Revit add-in compatibility
                if (System.Windows.Application.Current == null)
                {
                    new System.Windows.Application();
                }

                // Check if profile setup is required
                if (appProfileService.IsProfileSetupRequired)
                {
                    // EMERGENCY MODE: Use WinForms instead of WPF to avoid crashes
                    var emergencyDialog = new Views.EmergencyProfileSetup(appProfileService.ProfileService, appProfileService.StatusManager);
                    
                    var result = emergencyDialog.ShowDialog();
                    if (result == System.Windows.Forms.DialogResult.OK && emergencyDialog.CreatedProfile != null)
                    {
                        // Set the created profile as current
                        appProfileService.SetCurrentProfile(emergencyDialog.CreatedProfile);
                        
                        // Show success message
                        TaskDialog.Show("Profile Setup", $"Profile '{emergencyDialog.CreatedProfile.Name}' created successfully!");

                        // EMERGENCY MODE: Open WinForms main dialog (crash-safe)
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
                    // EMERGENCY MODE: Use WinForms ProfileManagement instead of WPF to prevent crashes
                    var emergencyProfileMgmt = new Views.EmergencyProfileManagementDialog(appProfileService);
                    
                    var result = emergencyProfileMgmt.ShowDialog();
                    
                    // DEBUG: Show what happened
                    TaskDialog.Show("Debug", $"Dialog result: {result}, ShouldOpenMainDialog: {emergencyProfileMgmt.ShouldOpenMainDialog}");
                    
                    if (result == System.Windows.Forms.DialogResult.OK && emergencyProfileMgmt.ShouldOpenMainDialog)
                    {
                        // Open WinForms main dialog after profile action
                        var emergencyMainDlg = new Views.EmergencyMainDialog(appProfileService, doc);
                        emergencyMainDlg.Show(); // Modeless
                        TaskDialog.Show("Debug", "MainDialog opened!");
                    }
                    else if (result == System.Windows.Forms.DialogResult.OK)
                    {
                        TaskDialog.Show("Debug", "OK result but ShouldOpenMainDialog is false - you may have just closed the dialog without switching profiles");
                    }
                    else
                    {
                        TaskDialog.Show("Debug", $"Dialog was cancelled or had other result: {result}");
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
    }
}
