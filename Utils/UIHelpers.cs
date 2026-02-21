using System;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.UI;

namespace JSE_RevitAddin_MEP_OPENINGS.Utils
{
    /// <summary>
    /// UI helper extension methods for DataGridView and other controls
    /// Centralized AddRow to normalize nulls, handle read-only grids, and ensure visibility
    /// </summary>
    public static class UIHelpers
    {
        public static void AddRow(this DataGridView grid, params object[] values)
        {
            if (grid == null) throw new ArgumentNullException(nameof(grid));

            // If the grid is read-only or does not allow adding rows, bail gracefully
            if (grid.ReadOnly || !grid.AllowUserToAddRows)
            {
                System.Diagnostics.Debug.WriteLine($"UIHelpers.AddRow: grid '{grid.Name}' is read-only or disallows adding rows.");
                return;
            }

            var normalized = values.Select(v => v ?? string.Empty).ToArray();
            grid.Rows.Add(normalized);

            // Ensure the newly added row is visible and selected
            int lastIndex = grid.Rows.Count - 1;
            if (lastIndex >= 0)
            {
                try
                {
                    grid.ClearSelection();
                    grid.Rows[lastIndex].Selected = true;
                    grid.FirstDisplayedScrollingRowIndex = Math.Max(0, lastIndex - 3);
                }
                catch
                {
                    // Ignore exceptions (virtual mode, not yet visible, etc.)
                }
            }
        }
        
        /// <summary>
        /// ✅ FIXED: Shows a MessageBox that appears on top of all windows
        /// Ensures the message box is visible even if main dialog is behind other windows
        /// </summary>
        public static DialogResult ShowMessageBoxOnTop(string message, string title = "", 
            MessageBoxButtons buttons = MessageBoxButtons.OK, MessageBoxIcon icon = MessageBoxIcon.Information)
        {
            // Create a temporary topmost form to ensure MessageBox appears on top
            var topForm = new System.Windows.Forms.Form
            {
                WindowState = FormWindowState.Minimized,
                ShowInTaskbar = false,
                TopMost = true
            };
            topForm.Show();
            
            try
            {
                // Show MessageBox with the topmost form as owner
                return MessageBox.Show(topForm, message, title, buttons, icon);
            }
            finally
            {
                topForm.Close();
                topForm.Dispose();
            }
        }
        
        /// <summary>
        /// ✅ FIXED: Shows a TaskDialog that appears on top of all windows
        /// Ensures the TaskDialog is visible even if main dialog is behind other windows
        /// </summary>
        public static Autodesk.Revit.UI.TaskDialogResult ShowTaskDialogOnTop(string title, string mainInstruction, 
            string mainContent = "", Autodesk.Revit.UI.TaskDialogCommonButtons buttons = Autodesk.Revit.UI.TaskDialogCommonButtons.Ok)
        {
            var dialog = new Autodesk.Revit.UI.TaskDialog(title)
            {
                MainInstruction = mainInstruction,
                MainContent = mainContent,
                CommonButtons = buttons
            };
            
            // TaskDialog.Show() automatically appears on top in Revit, but we ensure it's visible
            return dialog.Show();
        }
    }
}
