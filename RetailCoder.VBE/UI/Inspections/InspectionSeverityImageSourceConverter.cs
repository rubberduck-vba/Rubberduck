using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Rubberduck.Parsing.Inspections.Resources;

namespace Rubberduck.UI.Inspections
{
    public class InspectionSeverityImageSourceConverter : IValueConverter
    {
        private static readonly IDictionary<CodeInspectionSeverity,ImageSource> Icons = 
            new Dictionary<CodeInspectionSeverity, ImageSource>
            {
                { CodeInspectionSeverity.Hint, ToImageSource(Properties.Resources.information_white) },
                { CodeInspectionSeverity.Suggestion, ToImageSource(Properties.Resources.information) },
                { CodeInspectionSeverity.Warning, ToImageSource(Properties.Resources.exclamation) },
                { CodeInspectionSeverity.Error, ToImageSource(Properties.Resources.cross_circle) },
            };

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value.GetType() != typeof(CodeInspectionSeverity))
            {
                throw new ArgumentException("value must be a CodeInspectionSeverity");
            }

            var severity = (CodeInspectionSeverity)value;
            return Icons[severity];
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return Icons.First(f => Equals(f.Value, value)).Key;
        }

        private static ImageSource ToImageSource(Image source)
        {
            var ms = new MemoryStream();
            ((Bitmap)source).Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            var image = new BitmapImage();
            image.BeginInit();
            ms.Seek(0, SeekOrigin.Begin);
            image.StreamSource = ms;
            image.EndInit();

            return image;
        }
    }
}
