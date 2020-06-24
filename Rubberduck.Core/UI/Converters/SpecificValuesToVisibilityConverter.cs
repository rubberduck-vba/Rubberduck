using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace Rubberduck.UI.Converters
{
    public class SpecificValueToVisibilityConverter : IValueConverter
    {
        public object SpecialValue { get; set; }
        public bool CollapseSpecialValue { get; set; }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null)
            {
                return null;
            }

            return value.Equals(SpecialValue)
                ? SpecialValueVisibility
                : OtherValueVisibility;
        }

        private Visibility SpecialValueVisibility => CollapseSpecialValue ? Visibility.Collapsed : Visibility.Visible;
        private Visibility OtherValueVisibility => CollapseSpecialValue ? Visibility.Visible : Visibility.Collapsed;


        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return DependencyProperty.UnsetValue;
        }
    }
}