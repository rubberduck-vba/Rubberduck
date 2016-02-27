using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Data;
using Rubberduck.Inspections;

namespace Rubberduck.UI.Settings.Converters
{
    public class CodeInspectionSeverityEnumToTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var severities = (IEnumerable<CodeInspectionSeverity>)value;
            return severities.Select(s => RubberduckUI.ResourceManager.GetString("CodeInspectionSeverity_" + s)).ToArray();
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value;
        }
    }
}