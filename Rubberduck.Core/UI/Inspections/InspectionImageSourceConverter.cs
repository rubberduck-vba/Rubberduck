using System;
using System.Globalization;
using System.Windows.Data;
using Rubberduck.CodeAnalysis.Inspections;

namespace Rubberduck.UI.Inspections
{
    public class InspectionImageSourceConverter : IValueConverter
    {
        private static readonly InspectionSeverityImageSourceConverter SeverityConverter = new InspectionSeverityImageSourceConverter();

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(value is IInspection inspection))
            {
                return null;
            }

            return SeverityConverter.Convert(inspection.Severity, targetType, parameter, culture);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return SeverityConverter.ConvertBack(value, targetType, parameter, culture);
        }
    }
}
