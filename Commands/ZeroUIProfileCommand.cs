using System;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// ZERO-UI FALLBACK - NO FORMS - ONLY TASKDIALOG
    /// Complete profile management using only Revit's TaskDialog
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class ZeroUIProfileCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var appProfileService = ApplicationProfileService.Instance;
                
                // Update profile service for current document to ensure project-specific profiles
                var doc = commandData.Application.ActiveUIDocument.Document;
                var docPath = doc.PathName;
                appProfileService.UpdateForCurrentDocument(docPath);
                
                // Check if profile setup is required
                if (appProfileService.IsProfileSetupRequired)
                {
                    // Show profile creation options
                    var createResult = TaskDialog.Show("Profile Setup Required", 
                        "No profile exists. Would you like to create one?", 
                        TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);
                    
                    if (createResult == TaskDialogResult.Yes)
                    {
                        // Create a default profile using TaskDialog input
                        var profileName = GetProfileNameFromUser();
                        if (!string.IsNullOrEmpty(profileName))
                        {
                            // Create a simple default profile
                            var defaultProfile = CreateDefaultProfile(profileName);
                            appProfileService.SetCurrentProfile(defaultProfile);
                            
                            TaskDialog.Show("Success", $"Profile '{profileName}' created successfully!");
                            
                            // Show main interface options
                            ShowMainInterfaceOptions(appProfileService);
                        }
                    }
                }
                else
                {
                    // Profile exists - show options
                    ShowMainInterfaceOptions(appProfileService);
                }
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Exception: {ex.Message}");
                message = $"Error: {ex.Message}";
                return Result.Failed;
            }
        }
        
        private string GetProfileNameFromUser()
        {
            // Use TaskDialog to get profile name
            var result = TaskDialog.Show("Profile Name", 
                "Enter profile name:", 
                TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel);
            
            if (result == TaskDialogResult.Ok)
            {
                // For now, return a default name since TaskDialog doesn't have input
                return "DefaultProfile";
            }
            
            return string.Empty;
        }
        
        private Models.UserProfile CreateDefaultProfile(string name)
        {
            // Create a simple default profile
            var disciplines = new[]
            {
                new Models.Discipline { Name = "Architectural", IsSelected = true },
                new Models.Discipline { Name = "Structural", IsSelected = true },
                new Models.Discipline { Name = "Mechanical", IsSelected = true },
                new Models.Discipline { Name = "Electrical", IsSelected = true },
                new Models.Discipline { Name = "Plumbing", IsSelected = true }
            };
            
            return new Models.UserProfile
            {
                Name = name,
                Disciplines = disciplines.ToList(),
                IsActive = true
            };
        }
        
        private void ShowMainInterfaceOptions(ApplicationProfileService appProfileService)
        {
            var currentProfile = appProfileService.CurrentProfile;
            var profileInfo = currentProfile != null 
                ? $"Current Profile: {currentProfile.Name}\nDisciplines: {string.Join(", ", currentProfile.Disciplines.Select(d => d.Name))}"
                : "No current profile";
            
            TaskDialog.Show("MEP Openings - Main Interface", 
                $"{profileInfo}\n\nThis is where the main 3-panel interface would open.\n\nFor now, this is a zero-UI fallback to prevent crashes.");
        }
        
        private void ShowProfileSettings(ApplicationProfileService appProfileService)
        {
            var currentProfile = appProfileService.CurrentProfile;
            if (currentProfile != null)
            {
                var disciplines = string.Join("\n", currentProfile.Disciplines.Select(d => $"• {d.Name}: {(d.IsSelected ? "Enabled" : "Disabled")}"));
                
                TaskDialog.Show("Profile Settings", 
                    $"Profile: {currentProfile.Name}\n\nDisciplines:\n{disciplines}");
            }
            else
            {
                TaskDialog.Show("Profile Settings", "No profile is currently active.");
            }
        }
    }
}
