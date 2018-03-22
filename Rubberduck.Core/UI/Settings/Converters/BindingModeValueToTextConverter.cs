using System;
using System.Globalization;
using System.Windows.Data;

namespace Rubberduck.UI.Settings.Converters
{
    public class BindingModeValueToTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var mode = (Rubberduck.Settings.BindingMode)value;
            switch (mode)
            {
                case Rubberduck.Settings.BindingMode.EarlyBinding:
                    return RubberduckUI.UnitTestSettings_EarlyBinding;
                case Rubberduck.Settings.BindingMode.LateBinding:
                    return RubberduckUI.UnitTestSettings_LateBinding;
                default:
                    return value;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var text = (string)value;
            return text == RubberduckUI.UnitTestSettings_EarlyBinding
                ? Rubberduck.Settings.BindingMode.EarlyBinding
                : Rubberduck.Settings.BindingMode.LateBinding;
        }
    }
}
