using System;
using System.Globalization;
using System.Windows.Data;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;

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

            return RubberduckUI.ResourceManager.GetString("CodeInspectionSettings_" + inspection.InspectionType, CultureInfo.CurrentUICulture);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class InspectionDescriptionConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var inspection = value as IInspection;
            if (inspection == null)
            {
                return null;
            }

            return InspectionsUI.ResourceManager.GetString(inspection.Name + "Name", CultureInfo.CurrentUICulture);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
