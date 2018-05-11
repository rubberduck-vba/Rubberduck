using System;
using System.Globalization;
using System.Windows.Data;
using Rubberduck.Settings;
using Rubberduck.Resources.Settings;

namespace Rubberduck.UI.Settings.Converters
{
    public class AssertModeValueToTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var mode = (Rubberduck.Settings.AssertMode)value;
            switch (mode)
            {
                case AssertMode.StrictAssert:
                    return UnitTestingPage.UnitTestSettings_StrictAssert;
                case AssertMode.PermissiveAssert:
                    return UnitTestingPage.UnitTestSettings_PermissiveAssert;
                default:
                    return value;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var text = (string)value;
            return text == UnitTestingPage.UnitTestSettings_StrictAssert
                ? AssertMode.StrictAssert
                : AssertMode.PermissiveAssert;
        }
    }
}
