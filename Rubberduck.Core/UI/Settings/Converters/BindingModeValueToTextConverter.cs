using System;
using System.Globalization;
using System.Windows.Data;
using Rubberduck.Resources.Settings;
using BindingMode = Rubberduck.UnitTesting.Settings.BindingMode;

namespace Rubberduck.UI.Settings.Converters
{
    public class BindingModeValueToTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var mode = (BindingMode)value;
            switch (mode)
            {
                case BindingMode.EarlyBinding:
                    return UnitTestingPage.UnitTestSettings_EarlyBinding;
                case BindingMode.LateBinding:
                    return UnitTestingPage.UnitTestSettings_LateBinding;
                case BindingMode.DualBinding:
                    return UnitTestingPage.UnitTestSettings_DualBinding;
                default:
                    return value;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var text = (string)value;

            if (UnitTestingPage.UnitTestSettings_EarlyBinding.Equals(text))
            {
                return BindingMode.EarlyBinding;
            }

            return UnitTestingPage.UnitTestSettings_LateBinding.Equals(text)
                ? BindingMode.LateBinding
                : BindingMode.DualBinding;
        }
    }
}
