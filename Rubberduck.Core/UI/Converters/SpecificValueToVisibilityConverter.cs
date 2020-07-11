using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Data;

namespace Rubberduck.UI.Converters
{
    public class SpecificValuesToVisibilityConverter : IValueConverter
    {
        public IReadOnlyCollection<object> SpecialValues { get; set; }
        public bool CollapseSpecialValues { get; set; }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null)
            {
                return null;
            }

            return SpecialValues.Contains(value)
                ? SpecialValueVisibility
                : OtherValueVisibility;
        }

        private Visibility SpecialValueVisibility => CollapseSpecialValues ? Visibility.Collapsed : Visibility.Visible;
        private Visibility OtherValueVisibility => CollapseSpecialValues ? Visibility.Visible : Visibility.Collapsed;


        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return DependencyProperty.UnsetValue;
        }
    }
}