using System;
using System.Globalization;
using System.Linq;
using System.Windows.Data;

namespace Rubberduck.UI.Settings.Converters
{
    public class LanguagesEnumValueToTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var language = (Languages)value;
            return RubberduckUI.ResourceManager.GetString("Language_" + language);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var languageDisplayText = (string)value;
            var languages = Enum.GetValues(typeof(Languages));

            return languages.Cast<Languages>()
                    .First(lang => RubberduckUI.ResourceManager.GetString("Language_" + lang) == languageDisplayText);
        }
    }
}