using System;
using System.Globalization;
using System.Windows.Data;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.UI.Inspections
{
    public class InspectionTypeConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var inspection = value as IInspection;
            if (inspection == null)
            {
                return null;
            }
            return InspectionsUI.ResourceManager.GetString("CodeInspectionSettings_" + inspection.InspectionType.ToString(), CultureInfo.CurrentUICulture);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
