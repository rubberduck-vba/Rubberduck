using System;
using System.Globalization;
using System.Linq;
using System.Windows.Data;

namespace Rubberduck.UI.GeneralConverters
{
    public class EqualWidthConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            return Math.Abs(values.Cast<double>().Max()) < .1 ? -1 : values.Cast<double>().Max();
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            return new object[] { };
        }
    }
}