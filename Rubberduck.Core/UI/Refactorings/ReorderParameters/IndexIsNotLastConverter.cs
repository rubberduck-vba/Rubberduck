using System;
using System.Globalization;
using System.Windows.Data;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    public class IndexIsNotLastConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            var selectedIndex = (int)values[0];
            var indexCount = (int)values[1] - 1;
            return selectedIndex < indexCount;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
