using System;
using System.Windows.Data;

namespace Rubberduck.UI.Settings.Converters
{
    //Thanks to https://stackoverflow.com/a/2908885/4088852
    public class EnumToBooleanConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return value != null && value.Equals(parameter);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return value != null && value.Equals(true) ? parameter : Binding.DoNothing;
        }
    }
}
