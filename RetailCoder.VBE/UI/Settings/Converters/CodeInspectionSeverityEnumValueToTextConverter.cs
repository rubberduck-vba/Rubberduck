using System;
using System.Globalization;
using System.Linq;
using System.Windows.Data;
using Rubberduck.Inspections;

namespace Rubberduck.UI.Settings.Converters
{
    public class CodeInspectionSeverityEnumValueToTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var severity = (CodeInspectionSeverity)value;
            return RubberduckUI.ResourceManager.GetString("CodeInspectionSeverity_" + severity);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var severityDisplayText = (string)value;
            var severities = Enum.GetValues(typeof(CodeInspectionSeverity));

            return severities.Cast<CodeInspectionSeverity>()
                    .First(lang => RubberduckUI.ResourceManager.GetString("CodeInspectionSeverity_" + lang) == severityDisplayText);
        }
    }
}