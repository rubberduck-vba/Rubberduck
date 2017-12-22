using System;
using System.Windows.Data;
using System.Globalization;

namespace Rubberduck.UI.Settings.Converters
{
    public class BooleanToInvalidConfigurationIconConverter :IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (bool)value ? string.Empty : "IsNotValid";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value;
        }
    }
}