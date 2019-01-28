using System;
using System.Globalization;
using System.Windows.Data;

namespace Rubberduck.UI.Converters
{
    public class BooleanToNullableDoubleConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(value is double size) ||
                !(parameter is IConvertible input))
            {
                return false;
            }

            try
            {
                // ReSharper disable once CompareOfFloatsByEqualityOperator - these are hard coded values.
                return System.Convert.ToDouble(input) == size;
            }
            catch
            {
                return false;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {           
            if (!(value is bool toggle) ||
                !toggle ||
                !(parameter is IConvertible output))
            {
                return null;
            }

            try
            {
                return System.Convert.ToDouble(output);
            }
            catch
            {
                return null;
            }
        }
    }
}
