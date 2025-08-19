using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using System.Linq;
using System;

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

        /// <summary>
        /// Returns true if point is inside bounding box (inclusive).
        /// </summary>
        public static bool PointInBoundingBox(XYZ pt, BoundingBoxXYZ bbox)
        {
            return pt.X >= bbox.Min.X - 1e-6 && pt.X <= bbox.Max.X + 1e-6 &&
                   pt.Y >= bbox.Min.Y - 1e-6 && pt.Y <= bbox.Max.Y + 1e-6 &&
                   pt.Z >= bbox.Min.Z - 1e-6 && pt.Z <= bbox.Max.Z + 1e-6;
        }

        /// <summary>
        /// Converts side string to unit vector in WORLD coordinates (not damper local).
        /// </summary>
        public static XYZ OffsetVector4Way(string side, Transform damperT)
        {
            switch (side)
            {
                case "Right": return XYZ.BasisX;   // World +X direction
                case "Left": return -XYZ.BasisX;   // World -X direction  
                case "Top": return XYZ.BasisZ;   // World +Z direction
                case "Bottom": return -XYZ.BasisZ;   // World -Z direction
                default: return XYZ.BasisX;   // fallback to world +X
            }
        }

        // Helper to transform a bounding box by a transform
        public static BoundingBoxXYZ TransformBoundingBox(BoundingBoxXYZ bbox, Transform transform)
        {
            var corners = new[] {
                new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z),
                new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Max.Z),
                new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z),
                new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z),
                new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z),
                new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Max.Z),
                new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z),
                new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z)
            };
            var transformed = corners.Select(pt => transform.OfPoint(pt)).ToList();
            var min = new XYZ(transformed.Min(p => p.X), transformed.Min(p => p.Y), transformed.Min(p => p.Z));
            var max = new XYZ(transformed.Max(p => p.X), transformed.Max(p => p.Y), transformed.Max(p => p.Z));
            var newBox = new BoundingBoxXYZ { Min = min, Max = max };
            return newBox;
        }
    }
}