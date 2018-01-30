using System;
using System.Globalization;
using System.Windows.Data;

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
                    return RubberduckUI.GeneralSettings_PeriodDelimiter;
                case DelimiterOptions.Slash:
                    return RubberduckUI.GeneralSettings_SlashDelimiter;
                default:
                    return value;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var text = (string)value;
            return text == RubberduckUI.GeneralSettings_PeriodDelimiter
                ? DelimiterOptions.Period
                : DelimiterOptions.Slash;
        }
    }
}
