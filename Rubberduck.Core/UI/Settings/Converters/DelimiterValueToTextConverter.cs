using System;
using System.Globalization;
using System.Windows.Data;
using Rubberduck.Resources.Settings;

namespace Rubberduck.UI.Settings.Converters
{
    public class DelimiterValueToTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var mode = (DelimiterOptions)value;
            switch (mode)
            {
                case DelimiterOptions.Period:
                    return GeneralSettingsPage.PeriodDelimiter;
                case DelimiterOptions.Slash:
                    return GeneralSettingsPage.SlashDelimiter;
                default:
                    return value;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var text = (string)value;
            return text == GeneralSettingsPage.PeriodDelimiter
                ? DelimiterOptions.Period
                : DelimiterOptions.Slash;
        }
    }
}
