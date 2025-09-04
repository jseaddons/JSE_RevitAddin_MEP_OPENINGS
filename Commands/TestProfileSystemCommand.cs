using System;
using System.Windows;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.ViewModels;
using JSE_RevitAddin_MEP_OPENINGS.Views;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// Test command to demonstrate the profile system
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class TestProfileSystemCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiApp = commandData.Application;
                var uiDoc = uiApp.ActiveUIDocument;
                var doc = uiDoc.Document;

                // Initialize services
                var profileService = new ProfileManagementService();
                var statusManager = new StatusManager();

                // Check if profile setup is required
                if (profileService.IsProfileSetupRequired())
                {
                    // Show profile setup dialog
                    var viewModel = new ProfileSetupViewModel(profileService, statusManager);
                    var dialog = new ProfileSetupDialog(viewModel);

                    var result = dialog.ShowDialog();
                    if (result == true && dialog.CreatedProfile != null)
                    {
                        // Set the created profile as current
                        profileService.SetCurrentProfile(dialog.CreatedProfile);
                        
                        // Show success message
                        TaskDialog.Show("Profile Setup", 
                            $"Profile '{dialog.CreatedProfile.Name}' created successfully!\n\n" +
                            $"Primary Discipline: {dialog.CreatedProfile.PrimaryDiscipline?.Name}\n" +
                            $"MEP Disciplines: {string.Join(", ", dialog.CreatedProfile.MepDisciplines.Select(d => d.Name))}\n" +
                            $"Language: {dialog.CreatedProfile.Language}");
                    }
                    else
                    {
                        TaskDialog.Show("Profile Setup", "Profile setup was cancelled.");
                        return Result.Cancelled;
                    }
                }
                else
                {
                    // Show current profile info
                    var currentProfile = profileService.CurrentProfile;
                    if (currentProfile != null)
                    {
                        TaskDialog.Show("Current Profile", 
                            $"Current Profile: {currentProfile.Name}\n\n" +
                            $"Primary Discipline: {currentProfile.PrimaryDiscipline?.Name}\n" +
                            $"MEP Disciplines: {string.Join(", ", currentProfile.MepDisciplines.Select(d => d.Name))}\n" +
                            $"Language: {currentProfile.Language}\n" +
                            $"Created: {currentProfile.CreatedDate:yyyy-MM-dd HH:mm:ss}");
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
