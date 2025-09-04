using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace JSE_RevitAddin_MEP_OPENINGS.Views.Converters
{
    /// <summary>
    /// Converts a string to Visibility. Returns Visible if string is not null or empty, Collapsed otherwise.
    /// </summary>
    public class StringToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string str)
            {
                return string.IsNullOrWhiteSpace(str) ? System.Windows.Visibility.Collapsed : System.Windows.Visibility.Visible;
            }
            return System.Windows.Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
