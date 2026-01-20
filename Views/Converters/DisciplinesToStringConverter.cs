using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Data;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Views.Converters
{
    /// <summary>
    /// Converts a list of disciplines to a formatted string
    /// </summary>
    public class DisciplinesToStringConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is IEnumerable<Discipline> disciplines)
            {
                var disciplineNames = disciplines.Select(d => d.Name).ToList();
                if (disciplineNames.Any())
                {
                    return $"MEP: {string.Join(", ", disciplineNames)}";
                }
            }
            return "No MEP disciplines";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
