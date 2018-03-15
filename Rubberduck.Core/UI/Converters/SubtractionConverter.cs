using System;
using System.Globalization;
using System.Windows.Data;

namespace Rubberduck.UI.Converters
{
    public class SubtractionConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var typedValue = (double)value;
            if (!Double.TryParse((string)parameter, out var typedParam))
            {
                return (double)value;
            }
            return typedValue - typedParam;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var typedValue = (double)value;
            if (!Double.TryParse((string)parameter, out var typedParam))
            {
                return (double)value;
            }
            return typedValue + typedParam;
        }
    }
}
