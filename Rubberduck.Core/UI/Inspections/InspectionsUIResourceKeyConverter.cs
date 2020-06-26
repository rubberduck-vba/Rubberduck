using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.UI.Inspections
{
    public class InspectionsUIResourceKeyConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var key = value as string;

            if (key == null)
            {
                throw new ArgumentException("The value must be a string containing a key in the InspectionsUI resource.", "value");
            }

            return InspectionsUI.ResourceManager.GetString(key, culture);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return DependencyProperty.UnsetValue;
        }
    }
}