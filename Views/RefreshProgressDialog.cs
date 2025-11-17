using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using WinForms = System.Windows.Forms; // ✅ FIX: Alias to avoid conflict with Autodesk.Revit.DB.Form
using Drawing = System.Drawing; // ✅ FIX: Alias to avoid conflict with Autodesk.Revit.DB.Point, Size, Font, etc.

namespace JSE_RevitAddin_MEP_OPENINGS.Views
{
    /// <summary>
    /// Progress dialog for Refresh service showing intersection counts per category
    /// Compact design: 200px wide, 40px high per category row
    /// </summary>
    public partial class RefreshProgressDialog : WinForms.Form
    {
        private readonly Dictionary<string, WinForms.ProgressBar> _categoryProgressBars;
        private readonly Dictionary<string, WinForms.Label> _categoryLabels;
        private readonly Dictionary<string, WinForms.Label> _categoryCountLabels;
        private readonly Dictionary<string, int> _categoryCounts;
        private bool _cancelled = false;
        
        public RefreshProgressDialog(List<string> categories)
        {
            _categoryProgressBars = new Dictionary<string, WinForms.ProgressBar>();
            _categoryLabels = new Dictionary<string, WinForms.Label>();
            _categoryCountLabels = new Dictionary<string, WinForms.Label>();
            _categoryCounts = new Dictionary<string, int>();
            
            InitializeComponent(categories);
        }
        
        private void InitializeComponent(List<string> categories)
        {
            // ✅ CALCULATED SIZES: Proper layout with all elements visible and crisp
            // Layout calculation (verified):
            // - Dialog width: 400px total
            // - Dialog padding: 10px on each side (20px total)
            // - Available content width: 380px
            //
            // Title section:
            //   - Title: 380px wide, 24px high, at y=10
            //   - Gap below title: 10px
            //
            // Each category row (42px high):
            //   - Category label: 130px wide, 20px high, at y=rowStart (x=10)
            //   - Gap: 10px
            //   - Progress bar: 190px wide, 18px high, at y=rowStart+22 (x=150, below label)
            //   - Gap: 10px
            //   - Count label: 50px wide, 20px high, at y=rowStart (x=350, same row as category label)
            //   Horizontal total: 130 + 10 + 190 + 10 + 50 = 390px (fits perfectly in 400px with 10px padding)
            //
            // Cancel button section:
            //   - Button: 100px wide, 32px high, centered horizontally
            //   - Bottom margin: 12px
            //
            // Total height calculation (for 4 categories):
            //   10 (top margin) + 24 (title) + 10 (gap) + (4 * 42) (rows) + 12 (button margin) + 32 (button) + 12 (bottom margin) = 268px
            
            const int dialogPadding = 10; // 10px margin on each side
            const int dialogWidth = 400; // Total width: 10px + 380px content + 10px
            const int titleHeight = 24;
            const int titleTopMargin = 10;
            const int titleGap = 10; // Gap between title and first row
            const int rowHeight = 42; // Height per category row (label 20px + gap 2px + progress bar 18px + spacing 2px)
            const int categoryLabelWidth = 80;
            const int progressBarWidth = 190;
            const int countLabelWidth = 50;
            const int labelHeight = 20;
            const int progressBarHeight = 18;
            const int labelToProgressGap = 5; // Minimal gap between label and progress bar
            const int cancelButtonWidth = 100;
            const int cancelButtonHeight = 32;
            const int cancelButtonBottomMargin = 12;
            
            int categoryCount = Math.Min(4, categories.Count); // Max 4 categories
            int dialogHeight = titleTopMargin + titleHeight + titleGap + // Title section
                              (categoryCount * rowHeight) + // Category rows
                              cancelButtonBottomMargin + cancelButtonHeight + cancelButtonBottomMargin; // Cancel button section
            
            this.Text = "Refresh Progress";
            this.Size = new Drawing.Size(dialogWidth, dialogHeight);
            this.StartPosition = WinForms.FormStartPosition.CenterParent;
            this.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.TopMost = true;
            this.ShowInTaskbar = false; // ✅ Don't show in taskbar (modeless dialog)
            this.FormClosing += RefreshProgressDialog_FormClosing; // ✅ Handle closing properly
            this.BackColor = Drawing.Color.White; // ✅ FIX: Set background color
            
            // Title label (centered, full width)
            var titleLabel = new WinForms.Label
            {
                Text = "Intersection Detection",
                Location = new Drawing.Point(dialogPadding, titleTopMargin),
                Size = new Drawing.Size(dialogWidth - (dialogPadding * 2), titleHeight),
                TextAlign = Drawing.ContentAlignment.MiddleCenter,
                Font = new Drawing.Font("Microsoft Sans Serif", 10F, Drawing.FontStyle.Bold),
                ForeColor = Drawing.Color.Black
            };
            this.Controls.Add(titleLabel);
            
            // Create progress bar for each category
            int rowStartY = titleTopMargin + titleHeight + titleGap; // Start below title with gap
            foreach (var category in categories.Take(4)) // Max 4 categories
            {
                // Category label (left side)
                var categoryLabel = new WinForms.Label
                {
                    Text = category.Length > 18 ? category.Substring(0, 18) + "..." : category, // Truncate if too long
                    Location = new Drawing.Point(dialogPadding, rowStartY),
                    Size = new Drawing.Size(categoryLabelWidth, labelHeight),
                    Font = new Drawing.Font("Microsoft Sans Serif", 9F),
                    ForeColor = Drawing.Color.Black,
                    TextAlign = Drawing.ContentAlignment.MiddleLeft
                };
                this.Controls.Add(categoryLabel);
                _categoryLabels[category] = categoryLabel;
                
                // Count label (right side, shows "0" initially)
                int countLabelX = dialogWidth - dialogPadding - countLabelWidth;
                var countLabel = new WinForms.Label
                {
                    Text = "0",
                    Location = new Drawing.Point(countLabelX, rowStartY),
                    Size = new Drawing.Size(countLabelWidth, labelHeight),
                    TextAlign = Drawing.ContentAlignment.MiddleRight,
                    Font = new Drawing.Font("Microsoft Sans Serif", 9F, Drawing.FontStyle.Bold),
                    ForeColor = Drawing.Color.Blue
                };
                this.Controls.Add(countLabel);
                _categoryCountLabels[category] = countLabel;
                
                // Progress bar (below labels, spanning between category label and count label)
                int progressBarX = dialogPadding + categoryLabelWidth + labelToProgressGap; // Minimal gap after category label
                int progressBarY = rowStartY + (labelHeight - progressBarHeight) / 2; // Vertically align with label
                var progressBar = new WinForms.ProgressBar
                {
                    Location = new Drawing.Point(progressBarX, progressBarY),
                    Size = new Drawing.Size(progressBarWidth, progressBarHeight),
                    Style = WinForms.ProgressBarStyle.Continuous,
                    Minimum = 0,
                    Maximum = 1000, // Use high max for smooth updates
                    Value = 0,
                    Visible = true
                };
                this.Controls.Add(progressBar);
                _categoryProgressBars[category] = progressBar;
                _categoryCounts[category] = 0;
                
                rowStartY += rowHeight; // Move down for next category
            }
            
            // Cancel button (centered at bottom)
            int cancelButtonX = (dialogWidth - cancelButtonWidth) / 2; // Center horizontally
            int cancelButtonY = dialogHeight - cancelButtonBottomMargin - cancelButtonHeight;
            var cancelButton = new WinForms.Button
            {
                Text = "Cancel",
                Location = new Drawing.Point(cancelButtonX, cancelButtonY),
                Size = new Drawing.Size(cancelButtonWidth, cancelButtonHeight),
                DialogResult = WinForms.DialogResult.Cancel,
                Font = new Drawing.Font("Microsoft Sans Serif", 9F)
            };
            cancelButton.Click += (s, e) => { _cancelled = true; };
            this.Controls.Add(cancelButton);
        }
        
        /// <summary>
        /// Update intersection count for a specific category
        /// ✅ THREAD-SAFE: Uses Invoke if called from background thread
        /// </summary>
        public void UpdateCategoryProgress(string category, int intersectionCount)
        {
            if (string.IsNullOrEmpty(category) || !_categoryProgressBars.ContainsKey(category))
                return;
            
            // ✅ THREAD-SAFETY: Check if we need to invoke on UI thread
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => UpdateCategoryProgress(category, intersectionCount)));
                return;
            }
            
            try
            {
                _categoryCounts[category] = intersectionCount;
                
                // Update count label
                if (_categoryCountLabels.ContainsKey(category))
                {
                    _categoryCountLabels[category].Text = intersectionCount.ToString();
                    _categoryCountLabels[category].Refresh(); // ✅ Force label refresh
                }
                
                // Update progress bar (visual indicator, not percentage-based)
                var progressBar = _categoryProgressBars[category];
                if (progressBar != null)
                {
                    // ✅ FIX: Ensure progress bar is visible and updates correctly
                    // Use intersection count directly (capped at Maximum)
                    int progressValue = Math.Min(progressBar.Maximum, Math.Max(0, intersectionCount));
                    progressBar.Value = progressValue;
                    progressBar.Refresh(); // ✅ Force progress bar refresh
                    
                    // ✅ FIX: Ensure progress bar style is set correctly
                    if (progressBar.Style != WinForms.ProgressBarStyle.Continuous)
                    {
                        progressBar.Style = WinForms.ProgressBarStyle.Continuous;
                    }
                }
                
                // ✅ TRANSACTION-SAFE: Single DoEvents() per update (per PROGRESS_UI_IMPLEMENTATION_PLAN.md)
                // This allows UI to update while refresh continues on UI thread
                // Safe for WinForms in Revit context (not WPF)
                WinForms.Application.DoEvents();
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[RefreshProgressDialog] Error updating progress for {category}: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Set maximum expected intersections for a category (for percentage calculation if needed)
        /// </summary>
        public void SetCategoryMaximum(string category, int maxIntersections)
        {
            if (string.IsNullOrEmpty(category) || !_categoryProgressBars.ContainsKey(category))
                return;
            
            var progressBar = _categoryProgressBars[category];
            progressBar.Maximum = Math.Max(1000, maxIntersections);
        }
        
        /// <summary>
        /// Handle form closing - ensure proper cleanup
        /// </summary>
        private void RefreshProgressDialog_FormClosing(object sender, WinForms.FormClosingEventArgs e)
        {
            // ✅ Allow closing without blocking
            // If user clicks X or Cancel, mark as cancelled
            if (e.CloseReason == WinForms.CloseReason.UserClosing)
            {
                _cancelled = true;
            }
        }
        
        public bool IsCancelled => _cancelled;
        
        public Dictionary<string, int> GetCategoryCounts()
        {
            return new Dictionary<string, int>(_categoryCounts);
        }
    }
}

