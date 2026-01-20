using System;
using System.Drawing;
using System.Windows.Forms;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Views
{
    public partial class OpeningProgressDialog : System.Windows.Forms.Form
    {
        // Progress tracking (simplified for user's layout)
        private int _currentOpenings = 0;
        private int _totalOpenings = 0;
        
        // UI Controls (matching user's layout)
        private ProgressBar _progressBar;           // Middle row: Progress bar
        private Label _countLabel;                  // Bottom row: "15 / 20" count
        private Label _currentOperationLabel;       // Status text
        private ListBox _logListBox;                // Execution log
        private Button _cancelButton;               // Cancel button
        
        // State tracking
        private bool _cancelled = false;
        private bool _completed = false;
        private System.Windows.Forms.Timer _autoCloseTimer;
        
        public OpeningProgressDialog()
        {
            // Set logging context for progress dialog debugging
            DebugLogger.SetServiceContext("ProgressDialog");
            
            InitializeComponent();
            InitializeAutoCloseTimer();
        }
        
        private void InitializeAutoCloseTimer()
        {
            // Auto-close timer: 2 seconds after completion
            _autoCloseTimer = new System.Windows.Forms.Timer();
            _autoCloseTimer.Interval = 2000; // 2 seconds
            _autoCloseTimer.Tick += AutoCloseTimer_Tick;
        }
        
        private void AutoCloseTimer_Tick(object sender, EventArgs e)
        {
            _autoCloseTimer.Stop();
            if (_completed && !_cancelled)
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }
        
        private void InitializeComponent()
        {
            this.Text = "Opening Creation Progress";
            this.Size = new Size(500, 450); // Increased height to accommodate cancel button
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.TopMost = true; // Bring progress dialog to front
            
            // 1. Top Row: "Opening Created" label (centered)
            var openingCreatedLabel = new Label
            {
                Text = "Opening Created",
                Location = new System.Drawing.Point(50, 30),
                Size = new Size(400, 25),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft Sans Serif", 12F, FontStyle.Bold)
            };
            this.Controls.Add(openingCreatedLabel);
            
            // 2. Middle Row: Progress bar (centered, full width)
            _progressBar = new ProgressBar
            {
                Location = new System.Drawing.Point(50, 70),
                Size = new Size(400, 30),
                Style = ProgressBarStyle.Continuous,
                Minimum = 0,
                Maximum = 100,
                Value = 0
            };
            this.Controls.Add(_progressBar);
            
            // 3. Bottom Row: Count display "15 / 20" (centered)
            _countLabel = new Label
            {
                Text = "0 / 0",
                Location = new System.Drawing.Point(50, 110),
                Size = new Size(400, 25),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft Sans Serif", 14F, FontStyle.Bold)
            };
            this.Controls.Add(_countLabel);
            
            // 4. Status: Current operation text
            _currentOperationLabel = new Label
            {
                Text = "Initializing...",
                Location = new System.Drawing.Point(50, 150),
                Size = new Size(400, 20),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft Sans Serif", 9F)
            };
            this.Controls.Add(_currentOperationLabel);
            
            // 5. Log: Execution log with timestamps
            var logLabel = new Label
            {
                Text = "Execution Log:",
                Location = new System.Drawing.Point(50, 180),
                Size = new Size(100, 20)
            };
            this.Controls.Add(logLabel);
            
            _logListBox = new ListBox
            {
                Location = new System.Drawing.Point(50, 205),
                Size = new Size(400, 100), // Reduced height from 120 to 100 to make room for buttons
                Font = new Font("Consolas", 8F)
            };
            this.Controls.Add(_logListBox);
            
            // 6. Buttons: Cancel and Close buttons (moved higher for better visibility)
            _cancelButton = new Button
            {
                Text = "Cancel",
                Location = new System.Drawing.Point(300, 380), // Moved up from 340 to 380
                Size = new Size(80, 30)
            };
            _cancelButton.Click += CancelButton_Click;
            this.Controls.Add(_cancelButton);
            
            var closeButton = new Button
            {
                Text = "Close",
                Location = new System.Drawing.Point(390, 380), // Moved up from 340 to 380
                Size = new Size(80, 30),
                Enabled = false
            };
            closeButton.Click += CloseButton_Click;
            this.Controls.Add(closeButton);
        }
        
        // Event handlers for user control
        private void CancelButton_Click(object sender, EventArgs e)
        {
            _cancelled = true;
            _autoCloseTimer?.Stop();
            _currentOperationLabel.Text = "Operation cancelled by user";
            _logListBox.Items.Add($"[{DateTime.Now:HH:mm:ss}] Operation cancelled by user");
            _logListBox.TopIndex = _logListBox.Items.Count - 1;
            
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
        
        private void CloseButton_Click(object sender, EventArgs e)
        {
            _autoCloseTimer?.Stop();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
        
        public void UpdateProgress(
            int currentOpenings, int totalOpenings, string operation)
        {
            try
            {
                DebugLogger.Log($"UpdateProgress START: {currentOpenings}/{totalOpenings} - {operation}");
                
                // Update progress values
                _currentOpenings = currentOpenings;
                _totalOpenings = totalOpenings;
                
                // Calculate percentage
                double percentage = totalOpenings > 0 ? 100.0 * currentOpenings / totalOpenings : 0.0;
                
                // Update UI controls according to user's layout:
                // 1. Progress bar (middle row)
                _progressBar.Value = Math.Min(100, Math.Max(0, (int)percentage));
                
                // 2. Count display "15 / 20" (bottom row)
                _countLabel.Text = $"{currentOpenings} / {totalOpenings}";
                
                // 3. Current operation text
                _currentOperationLabel.Text = operation;
                
                // 4. Add to log
                _logListBox.Items.Add($"[{DateTime.Now:HH:mm:ss}] {operation}");
                _logListBox.TopIndex = _logListBox.Items.Count - 1; // Auto-scroll to bottom
                
                // Check if completed
                if (currentOpenings >= totalOpenings && totalOpenings > 0)
                {
                    _completed = true;
                    _currentOperationLabel.Text = "Opening creation completed successfully!";
                    _logListBox.Items.Add($"[{DateTime.Now:HH:mm:ss}] Opening creation completed successfully!");
                    _logListBox.TopIndex = _logListBox.Items.Count - 1;
                    
                    // Start auto-close timer (2 seconds delay)
                    _autoCloseTimer.Start();
                    
                    DebugLogger.Log("Opening creation completed - starting auto-close timer");
                }
                
                // Single UI pump (Jeremy Tammik's Building Coder pattern for WinForms)
                System.Windows.Forms.Application.DoEvents();
                
                DebugLogger.Log($"Progress updated: {currentOpenings}/{totalOpenings} ({percentage:F1}%)");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error updating progress: {ex.Message}");
                throw;
            }
        }
        
        public bool IsCancelled => _cancelled;
        public bool IsCompleted => _completed;
    }
}
