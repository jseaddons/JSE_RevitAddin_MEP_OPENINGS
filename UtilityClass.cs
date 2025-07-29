using System.Windows;
using System.Windows.Controls;

namespace JSE_RevitAddin_MEP_OPENINGS
{
    public static class UtilityClass
    {
        public static string GetPrefixFromUser()
        {
            // Create a simple WPF window for user input
            Window inputWindow = new Window
            {
                Title = "Enter Prefix",
                Width = 300,
                Height = 150,
                WindowStartupLocation = WindowStartupLocation.CenterScreen
            };

            // Create a stack panel to hold the input elements
            StackPanel stackPanel = new StackPanel { Margin = new Thickness(10) };

            // Add a label
            Label label = new Label { Content = "Prefix:" };
            stackPanel.Children.Add(label);

            // Add a text box for input
            TextBox textBox = new TextBox { Width = 250 };
            stackPanel.Children.Add(textBox);

            // Add OK and Cancel buttons
            StackPanel buttonPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center };
            Button okButton = new Button { Content = "OK", Width = 75, Margin = new Thickness(5) };
            Button cancelButton = new Button { Content = "Cancel", Width = 75, Margin = new Thickness(5) };
            buttonPanel.Children.Add(okButton);
            buttonPanel.Children.Add(cancelButton);
            stackPanel.Children.Add(buttonPanel);

            // Set the content of the window
            inputWindow.Content = stackPanel;

            // Handle button clicks
            string result = string.Empty;
            okButton.Click += (s, e) => { result = textBox.Text; inputWindow.DialogResult = true; inputWindow.Close(); };
            cancelButton.Click += (s, e) => { inputWindow.DialogResult = false; inputWindow.Close(); };

            // Show the window and return the result
            bool? dialogResult = inputWindow.ShowDialog();
            return dialogResult == true ? result : string.Empty;
        }
    }
}
