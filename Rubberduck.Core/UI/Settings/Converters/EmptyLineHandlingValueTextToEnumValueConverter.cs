using System;
using System.Globalization;
using System.Linq;
using System.Windows.Data;
using Rubberduck.Resources;
using Rubberduck.SmartIndenter;

namespace Rubberduck.UI.Settings.Converters
{
    public class EmptyLineHandlingValueTextToEnumValueConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var enumValue = (EmptyLineHandling)value;
            return RubberduckUI.ResourceManager.GetString("EmptyLineHandling_" + enumValue, CultureInfo.CurrentUICulture);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var selectedString = (string)value;

            var values = Enum.GetValues(typeof(EmptyLineHandling)).OfType<EmptyLineHandling>();

            foreach (var v in values.Where(v =>
                RubberduckUI.ResourceManager.GetString("EmptyLineHandling_" + v, CultureInfo.CurrentUICulture) == selectedString))
            {
                return v;
            }

            return value;
        }
    }
}
