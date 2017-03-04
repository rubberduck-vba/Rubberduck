using System;
using System.Globalization;
using System.Windows.Data;

namespace Rubberduck.UI.Settings.Converters
{
    public class AssertModeValueToTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var mode = (Rubberduck.Settings.AssertMode)value;
            switch (mode)
            {
                case Rubberduck.Settings.AssertMode.StrictAssert:
                    return RubberduckUI.UnitTestSettings_StrictAssert;
                case Rubberduck.Settings.AssertMode.PermissiveAssert:
                    return RubberduckUI.UnitTestSettings_PermissiveAssert;
                default:
                    return value;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var text = (string)value;
            return text == RubberduckUI.UnitTestSettings_StrictAssert
                ? Rubberduck.Settings.AssertMode.StrictAssert
                : Rubberduck.Settings.AssertMode.PermissiveAssert;
        }
    }
}
