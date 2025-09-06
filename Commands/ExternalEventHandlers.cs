using System;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Views;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// External Event handler for UI operations to prevent crashes in Revit
    /// </summary>
    public class ShowDialogExternalEvent : IExternalEventHandler
    {
        private ApplicationProfileService? _appProfileService;
        private Document? _document;
        private DialogType _dialogType;

        public enum DialogType
        {
            MinimalTest,
            ProfileSetup,
            ProfileManagement,
            MainDialog
        }

        public void SetParameters(ApplicationProfileService appProfileService, Document? document, DialogType dialogType)
        {
            _appProfileService = appProfileService;
            _document = document;
            _dialogType = dialogType;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                switch (_dialogType)
                {
                    case DialogType.MinimalTest:
                        ShowMinimalTestDialog();
                        break;
                    case DialogType.ProfileSetup:
                        ShowProfileSetupDialog();
                        break;
                    case DialogType.ProfileManagement:
                        ShowProfileManagementDialog();
                        break;
                    case DialogType.MainDialog:
                        ShowMainDialog();
                        break;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Dialog execution failed: {ex.Message}");
            }
        }

        private void ShowMinimalTestDialog()
        {
            using (var testDialog = new System.Windows.Forms.Form())
            {
                testDialog.Text = "Minimal Test Dialog - External Event";
                testDialog.Size = new System.Drawing.Size(400, 300);
                testDialog.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                
                var label = new System.Windows.Forms.Label
                {
                    Text = "This dialog is shown via External Event.\nThis should NOT crash on second execution.\n\nClose this and run the command again to test.",
                    Dock = System.Windows.Forms.DockStyle.Fill,
                    TextAlign = System.Drawing.ContentAlignment.MiddleCenter
                };
                
                testDialog.Controls.Add(label);
                testDialog.ShowDialog();
            }
        }

        private void ShowProfileSetupDialog()
        {
            if (_appProfileService == null) return;
            
            using (var emergencyDialog = new EmergencyProfileSetup(_appProfileService.ProfileService, _appProfileService.StatusManager))
            {
                var result = emergencyDialog.ShowDialog();
                if (result == DialogResult.OK && emergencyDialog.CreatedProfile != null)
                {
                    _appProfileService.SetCurrentProfile(emergencyDialog.CreatedProfile);
                    TaskDialog.Show("Profile Setup", $"Profile '{emergencyDialog.CreatedProfile.Name}' created successfully!");
                }
            }
        }

        private void ShowProfileManagementDialog()
        {
            if (_appProfileService == null) return;
            
            using (var emergencyProfileMgmt = new EmergencyProfileManagementDialog(_appProfileService))
            {
                var result = emergencyProfileMgmt.ShowDialog();
                if (result == DialogResult.OK && emergencyProfileMgmt.ShouldOpenMainDialog)
                {
                    ShowMainDialog();
                }
            }
        }

        private void ShowMainDialog()
        {
            if (_appProfileService == null || _document == null) return;
            
            using (var emergencyMainDlg = new EmergencyMainDialog(_appProfileService, _document))
            {
                emergencyMainDlg.ShowDialog();
            }
        }

        public string GetName()
        {
            return "Show Dialog External Event";
        }
    }
}