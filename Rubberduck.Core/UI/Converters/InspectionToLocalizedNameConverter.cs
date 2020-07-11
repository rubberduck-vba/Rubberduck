using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.UI.Converters
{
    public class InspectionToLocalizedNameConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var inspectionName = value is IInspection inspection
                ? inspection.AnnotationName
                : value as string;

            if (inspectionName == null)
            {
                throw new ArgumentException("The value must be an instance of IInspection or a string containing the programmatic name of an inspection.", "value");
            }

            return InspectionNames.ResourceManager.GetString($"{inspectionName}Inspection", culture);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return DependencyProperty.UnsetValue;
        }
    }
}