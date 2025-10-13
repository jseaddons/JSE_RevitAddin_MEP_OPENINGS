using System;
using System.Linq;
using System.Windows.Forms;

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
    }
}
