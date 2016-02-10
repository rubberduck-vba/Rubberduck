using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Data;

namespace Rubberduck.UI.Settings.Converters
{
    public class LanguagesEnumToTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var languages = (IEnumerable<Languages>)value;
            return languages.Select(l => RubberduckUI.ResourceManager.GetString("Language_" + l)).ToArray();
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value;
        }
    }
}