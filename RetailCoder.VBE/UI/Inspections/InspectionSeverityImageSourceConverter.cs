using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Media;
using Rubberduck.Parsing.Inspections.Resources;
using ImageSourceConverter = Rubberduck.UI.Converters.ImageSourceConverter;

namespace Rubberduck.UI.Inspections
{
    public class InspectionSeverityImageSourceConverter : ImageSourceConverter
    {
        private readonly IDictionary<CodeInspectionSeverity,ImageSource> _icons = 
            new Dictionary<CodeInspectionSeverity, ImageSource>
            {
                { CodeInspectionSeverity.Hint, ToImageSource(Properties.Resources.information_white) },
                { CodeInspectionSeverity.Suggestion, ToImageSource(Properties.Resources.information) },
                { CodeInspectionSeverity.Warning, ToImageSource(Properties.Resources.exclamation) },
                { CodeInspectionSeverity.Error, ToImageSource(Properties.Resources.cross_circle) },
            };

        public override object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value?.GetType() != typeof(CodeInspectionSeverity))
            {
                throw new ArgumentException("value must be a CodeInspectionSeverity");
            }

            var severity = (CodeInspectionSeverity)value;
            return _icons[severity];
        }

        public override object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return _icons.First(f => Equals(f.Value, value)).Key;
        }
    }
}
