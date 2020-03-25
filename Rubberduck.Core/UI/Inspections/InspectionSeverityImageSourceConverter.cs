using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Media;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.Resources.Inspections;
using ImageSourceConverter = Rubberduck.UI.Converters.ImageSourceConverter;

namespace Rubberduck.UI.Inspections
{
    public class InspectionSeverityImageSourceConverter : ImageSourceConverter
    {
        private static readonly IDictionary<CodeInspectionSeverity,ImageSource> Icons = 
            new Dictionary<CodeInspectionSeverity, ImageSource>
            {
                { CodeInspectionSeverity.DoNotShow, null },
                { CodeInspectionSeverity.Hint, ToImageSource(InspectionsUI.information_white) },
                { CodeInspectionSeverity.Suggestion, ToImageSource(InspectionsUI.information) },
                { CodeInspectionSeverity.Warning, ToImageSource(InspectionsUI.exclamation) },
                { CodeInspectionSeverity.Error, ToImageSource(InspectionsUI.cross_circle) },
            };

        public override object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value is CodeInspectionSeverity severity ? Icons[severity] : null;
        }

        public override object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return Icons.First(f => Equals(f.Value, value)).Key;
        }
    }
}
