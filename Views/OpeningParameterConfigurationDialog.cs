using System;
using System.Drawing;
using System.Windows.Forms;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Views
{
    /// <summary>
    /// Dialog for configuring opening parameter names
    /// </summary>
    public partial class OpeningParameterConfigurationDialog : System.Windows.Forms.Form
    {
        private OpeningParameterConfiguration? _config;
        
        // Dimension parameter controls
        private TextBox _widthParameterTextBox;
        private TextBox _heightParameterTextBox;
        private TextBox _diameterParameterTextBox;
        
        // Level and elevation parameter controls
        private TextBox _levelParameterTextBox;
        private TextBox _centerFromFFLParameterTextBox;
        private TextBox _ceilingLevelFromFFLParameterTextBox;
        
        // Note: Mark values are handled by the existing MarkParameterAddValue command system
        
        // Action buttons
        private Button _okButton;
        private Button _cancelButton;
        private Button _resetDefaultsButton;
        private Button _loadFromProjectButton;
        
        public OpeningParameterConfigurationDialog(OpeningParameterConfiguration? config = null)
        {
            _config = config ?? new OpeningParameterConfiguration();
            InitializeComponent();
            LoadConfiguration();
        }
        
        private void InitializeComponent()
        {
            // Form properties
            this.Text = "Opening Parameter Configuration";
            this.Size = new Size(500, 320);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            
            // Dimension parameters group
            var dimensionGroupBox = new GroupBox
            {
                Text = "Dimension Parameters",
                Location = new System.Drawing.Point(10, 10),
                Size = new Size(460, 100),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(dimensionGroupBox);
            
            // Width parameter
            var widthLabel = new Label
            {
                Text = "Width Parameter:",
                Location = new System.Drawing.Point(10, 25),
                Size = new Size(120, 20)
            };
            dimensionGroupBox.Controls.Add(widthLabel);
            
            _widthParameterTextBox = new TextBox
            {
                Location = new System.Drawing.Point(140, 23),
                Size = new Size(200, 20),
                Text = "Width"
            };
            dimensionGroupBox.Controls.Add(_widthParameterTextBox);
            
            // Height parameter
            var heightLabel = new Label
            {
                Text = "Height Parameter:",
                Location = new System.Drawing.Point(10, 50),
                Size = new Size(120, 20)
            };
            dimensionGroupBox.Controls.Add(heightLabel);
            
            _heightParameterTextBox = new TextBox
            {
                Location = new System.Drawing.Point(140, 48),
                Size = new Size(200, 20),
                Text = "Height"
            };
            dimensionGroupBox.Controls.Add(_heightParameterTextBox);
            
            // Diameter parameter
            var diameterLabel = new Label
            {
                Text = "Diameter Parameter:",
                Location = new System.Drawing.Point(10, 75),
                Size = new Size(120, 20)
            };
            dimensionGroupBox.Controls.Add(diameterLabel);
            
            _diameterParameterTextBox = new TextBox
            {
                Location = new System.Drawing.Point(140, 73),
                Size = new Size(200, 20),
                Text = "Outside Diameter"
            };
            dimensionGroupBox.Controls.Add(_diameterParameterTextBox);
            
            // Level and elevation parameters group
            var levelGroupBox = new GroupBox
            {
                Text = "Level and Elevation Parameters",
                Location = new System.Drawing.Point(10, 120),
                Size = new Size(460, 100),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(levelGroupBox);
            
            // Level parameter
            var levelLabel = new Label
            {
                Text = "Level Parameter:",
                Location = new System.Drawing.Point(10, 25),
                Size = new Size(120, 20)
            };
            levelGroupBox.Controls.Add(levelLabel);
            
            _levelParameterTextBox = new TextBox
            {
                Location = new System.Drawing.Point(140, 23),
                Size = new Size(200, 20),
                Text = "Level"
            };
            levelGroupBox.Controls.Add(_levelParameterTextBox);
            
            // Center from FFL parameter
            var centerFromFFLLabel = new Label
            {
                Text = "Center From FFL:",
                Location = new System.Drawing.Point(10, 50),
                Size = new Size(120, 20)
            };
            levelGroupBox.Controls.Add(centerFromFFLLabel);
            
            _centerFromFFLParameterTextBox = new TextBox
            {
                Location = new System.Drawing.Point(140, 48),
                Size = new Size(200, 20),
                Text = "Center From FFL"
            };
            levelGroupBox.Controls.Add(_centerFromFFLParameterTextBox);
            
            // Ceiling level from FFL parameter
            var ceilingLevelLabel = new Label
            {
                Text = "Ceiling Level From FFL:",
                Location = new System.Drawing.Point(10, 75),
                Size = new Size(120, 20)
            };
            levelGroupBox.Controls.Add(ceilingLevelLabel);
            
            _ceilingLevelFromFFLParameterTextBox = new TextBox
            {
                Location = new System.Drawing.Point(140, 73),
                Size = new Size(200, 20),
                Text = "Ceiling Level From FFL"
            };
            levelGroupBox.Controls.Add(_ceilingLevelFromFFLParameterTextBox);
            
            // Note: Mark values are handled by the existing MarkParameterAddValue command system
            // No additional configuration needed for mark values
            
            // Action buttons
            _loadFromProjectButton = new Button
            {
                Text = "Load from Project",
                Location = new System.Drawing.Point(10, 230),
                Size = new Size(120, 25),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            _loadFromProjectButton.Click += LoadFromProjectButton_Click;
            this.Controls.Add(_loadFromProjectButton);
            
            _resetDefaultsButton = new Button
            {
                Text = "Reset to Defaults",
                Location = new System.Drawing.Point(140, 230),
                Size = new Size(120, 25),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            _resetDefaultsButton.Click += ResetDefaultsButton_Click;
            this.Controls.Add(_resetDefaultsButton);
            
            _okButton = new Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(320, 230),
                Size = new Size(80, 25),
                DialogResult = DialogResult.OK,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            _okButton.Click += OkButton_Click;
            this.Controls.Add(_okButton);
            
            _cancelButton = new Button
            {
                Text = "Cancel",
                Location = new System.Drawing.Point(410, 230),
                Size = new Size(80, 25),
                DialogResult = DialogResult.Cancel,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            this.Controls.Add(_cancelButton);
        }
        
        private void LoadConfiguration()
        {
            _widthParameterTextBox.Text = _config.WidthParameterName;
            _heightParameterTextBox.Text = _config.HeightParameterName;
            _diameterParameterTextBox.Text = _config.DiameterParameterName;
            _levelParameterTextBox.Text = _config.LevelParameterName;
            _centerFromFFLParameterTextBox.Text = _config.CenterFromFFLParameterName;
            _ceilingLevelFromFFLParameterTextBox.Text = _config.CeilingLevelFromFFLParameterName;
        }
        
        private void OkButton_Click(object sender, EventArgs e)
        {
            // Validate configuration
            if (string.IsNullOrEmpty(_widthParameterTextBox.Text) ||
                string.IsNullOrEmpty(_heightParameterTextBox.Text) ||
                string.IsNullOrEmpty(_levelParameterTextBox.Text))
            {
                MessageBox.Show("Please fill in all required parameter names.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            // Save configuration
            _config.WidthParameterName = _widthParameterTextBox.Text;
            _config.HeightParameterName = _heightParameterTextBox.Text;
            _config.DiameterParameterName = _diameterParameterTextBox.Text;
            _config.LevelParameterName = _levelParameterTextBox.Text;
            _config.CenterFromFFLParameterName = _centerFromFFLParameterTextBox.Text;
            _config.CeilingLevelFromFFLParameterName = _ceilingLevelFromFFLParameterTextBox.Text;
            
            this.DialogResult = DialogResult.OK;
        }
        
        private void ResetDefaultsButton_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("Are you sure you want to reset to default parameter names?", 
                "Confirm Reset", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            
            if (result == DialogResult.Yes)
            {
                _config = new OpeningParameterConfiguration();
                LoadConfiguration();
            }
        }
        
        private void LoadFromProjectButton_Click(object sender, EventArgs e)
        {
            // TODO: Implement loading parameter names from project
            MessageBox.Show("Load from Project functionality will be implemented.", "Info", 
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        public OpeningParameterConfiguration? GetConfiguration()
        {
            return _config;
        }
    }
}
