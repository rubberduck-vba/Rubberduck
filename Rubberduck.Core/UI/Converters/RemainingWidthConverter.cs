using System;
using System.Globalization;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Data;

namespace Rubberduck.UI.Converters
{
    public class RemainingWidthConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(value is ListView listview) || !(listview?.View is GridView grid))
            {
                return null;
            }

            return listview.Width - grid.Columns.Where(column => !double.IsNaN(column.Width)).Sum(column => column.Width) - grid.Columns.Count;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return null;
        }
    }
}
