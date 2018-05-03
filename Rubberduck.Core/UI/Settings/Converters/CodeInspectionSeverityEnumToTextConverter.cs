using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Data;
using Rubberduck.Parsing.Inspections;

namespace Rubberduck.UI.Settings.Converters
{
    public class CodeInspectionSeverityEnumToTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (value as IEnumerable<CodeInspectionSeverity>)?
                .Select(s => Resources.Inspections.InspectionsUI.ResourceManager.GetString("CodeInspectionSeverity_" + s, CultureInfo.CurrentUICulture))
                .ToArray();
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value;
        }
    }
}
