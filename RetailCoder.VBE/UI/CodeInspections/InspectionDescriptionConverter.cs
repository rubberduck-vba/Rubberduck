using System;
using System.Globalization;
using System.Windows.Data;
using Rubberduck.Inspections;

namespace Rubberduck.UI.CodeInspections
{
    public class InspectionDescriptionConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var inspection = value as IInspection;
            if (inspection == null)
            {
                return null;
            }

            return inspection.Name;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}