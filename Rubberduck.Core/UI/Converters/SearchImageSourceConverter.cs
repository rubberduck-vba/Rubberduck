using System;
using System.Globalization;
using System.Windows.Media;

namespace Rubberduck.UI.Converters
{
    public class SearchImageSourceConverter : ImageSourceConverter
    {
        private readonly ImageSource _search = ToImageSource(Resources.RubberduckUI.magnifier_medium);
        private readonly ImageSource _clear = ToImageSource(Resources.RubberduckUI.cross_script);

        public override object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return string.IsNullOrEmpty(value?.ToString()) ? _search : _clear;
        }
    }
}
