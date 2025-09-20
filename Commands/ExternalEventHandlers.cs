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
        private UIDocument? _uiDocument;
        private DialogType _dialogType;

        public enum DialogType
        {
            MinimalTest,
            ProfileSetup,
            ProfileManagement,
            MainDialog
        }

        public void SetParameters(ApplicationProfileService appProfileService, Document? document, UIDocument? uiDocument, DialogType dialogType)
        {
            _appProfileService = appProfileService;
            _document = document;
            _uiDocument = uiDocument;
            _dialogType = dialogType;
            
            // Set logging context for external events debugging
            DebugLogger.SetServiceContext("ExternalEvents");
            DebugLogger.Info($"ExternalEventHandlers.SetParameters: UIDocument = {(uiDocument != null ? "NOT NULL" : "NULL")}");
            DebugLogger.Info($"ExternalEventHandlers.SetParameters: Document = {(document != null ? "NOT NULL" : "NULL")}");
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
            
            DebugLogger.Info($"ExternalEventHandlers.ShowMainDialog: About to create EmergencyMainDialog");
            DebugLogger.Info($"ExternalEventHandlers.ShowMainDialog: _uiDocument = {(_uiDocument != null ? "NOT NULL" : "NULL")}");
            DebugLogger.Info($"ExternalEventHandlers.ShowMainDialog: _document = {(_document != null ? "NOT NULL" : "NULL")}");
            
            // CRITICAL FIX: Get UIDocument from current Revit context if not available
            var uiDocument = _uiDocument ?? GetCurrentUIDocument();
            DebugLogger.Info($"ExternalEventHandlers.ShowMainDialog: Final UIDocument = {(uiDocument != null ? "NOT NULL" : "NULL")}");
            
            using (var emergencyMainDlg = new EmergencyMainDialog(_appProfileService, _document, uiDocument))
            {
                DebugLogger.Info($"ExternalEventHandlers.ShowMainDialog: EmergencyMainDialog created successfully");
                emergencyMainDlg.ShowDialog();
            }
        }
        
        /// <summary>
        /// Gets the current UIDocument from Revit context
        /// </summary>
        private UIDocument? GetCurrentUIDocument()
        {
            try
            {
                // In ExternalEventHandlers, we don't have direct access to UIApplication
                // The UIDocument should be passed through the SetParameters method
                DebugLogger.Warning("ExternalEventHandlers.GetCurrentUIDocument: Cannot get UIDocument from Revit context in ExternalEventHandlers");
                return null;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"ExternalEventHandlers.GetCurrentUIDocument: Error getting UIDocument: {ex.Message}");
                return null;
            }
        }

        public string GetName()
        {
            return "Show Dialog External Event";
        }
    }
}