using System;
using System.Globalization;
using System.Windows.Data;
using Rubberduck.Parsing.Inspections.Abstract;

namespace Rubberduck.UI.Inspections
{
    public class InspectionImageSourceConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var inspection = value as IInspection;
            if (inspection == null)
            {
                return null;
            }

            var converter = new InspectionSeverityImageSourceConverter();
            return converter.Convert(inspection.Severity, targetType, parameter, culture);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var converter = new InspectionSeverityImageSourceConverter();
            return converter.ConvertBack(value, targetType, parameter, culture);
        }
    }
}
