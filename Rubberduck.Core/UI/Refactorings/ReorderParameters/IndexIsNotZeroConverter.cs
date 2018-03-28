using System;
using System.Globalization;
using System.Windows.Data;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    public class IndexIsNotZeroConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (int)value != 0;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
