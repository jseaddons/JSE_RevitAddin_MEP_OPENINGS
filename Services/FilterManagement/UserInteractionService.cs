using System;
using System.Windows.Forms;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.FilterManagement
{
    /// <summary>
    /// Team J: User interaction service implementation.
    /// SOLID: Single Responsibility - user interaction only.
    /// </summary>
    public class UserInteractionService : IUserInteractionService
    {
        /// <summary>
        /// Gets filter name from user via input dialog.
        /// ✅ REUSE: Existing InputDialog implementation from FilterManagementService.
        /// </summary>
        public string GetFilterNameFromUser(string title, string prompt, string defaultValue)
        {
            var inputDialog = new InputDialog(title, prompt, defaultValue);
            try
            {
                // Make sure dialog appears on top and is modal
                inputDialog.TopMost = true;
                var result = inputDialog.ShowDialog();
                if (result == DialogResult.OK)
                {
                    var filterName = inputDialog.InputText?.Trim();
                    if (string.IsNullOrEmpty(filterName))
                    {
                        ShowWarning("Filter name cannot be empty.");
                        return null;
                    }
                    return filterName;
                }
            }
            finally
            {
                inputDialog.Dispose();
            }
            return null;
        }
        
        /// <summary>
        /// Shows error message to user.
        /// </summary>
        public void ShowError(string message)
        {
            MessageBox.Show(message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        
        /// <summary>
        /// Shows warning message to user.
        /// </summary>
        public void ShowWarning(string message)
        {
            MessageBox.Show(message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        
        /// <summary>
        /// Shows confirmation dialog to user.
        /// </summary>
        public bool ShowConfirmation(string message, string title)
        {
            return MessageBox.Show(message, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
        }
    }
    
    /// <summary>
    /// Simple input dialog for getting user input.
    /// ✅ REUSE: Migrated from FilterManagementService.
    /// </summary>
    public class InputDialog : System.Windows.Forms.Form
    {
        private TextBox _inputTextBox;
        private Button _okButton;
        private Button _cancelButton;

        public string InputText => _inputTextBox?.Text;

        public InputDialog(string title, string prompt, string defaultValue = "")
        {
            InitializeComponent(title, prompt, defaultValue);
        }

        private void InitializeComponent(string title, string prompt, string defaultValue)
        {
            Text = title;
            Size = new System.Drawing.Size(400, 150);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            TopMost = true; // Keep dialog on top
            ShowInTaskbar = false; // Don't show in taskbar

            var promptLabel = new Label
            {
                Text = prompt,
                Location = new System.Drawing.Point(12, 12),
                Size = new System.Drawing.Size(360, 20),
                AutoSize = true
            };
            Controls.Add(promptLabel);

            _inputTextBox = new TextBox
            {
                Text = defaultValue,
                Location = new System.Drawing.Point(12, 40),
                Size = new System.Drawing.Size(360, 20)
            };
            Controls.Add(_inputTextBox);

            _okButton = new Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(216, 70),
                Size = new System.Drawing.Size(75, 23),
                DialogResult = DialogResult.OK
            };
            Controls.Add(_okButton);

            _cancelButton = new Button
            {
                Text = "Cancel",
                Location = new System.Drawing.Point(297, 70),
                Size = new System.Drawing.Size(75, 23),
                DialogResult = DialogResult.Cancel
            };
            Controls.Add(_cancelButton);

            AcceptButton = _okButton;
            CancelButton = _cancelButton;
            
            // Ensure the dialog is properly modal and visible
            _inputTextBox.Focus();
            _inputTextBox.SelectAll();
        }
    }
}

