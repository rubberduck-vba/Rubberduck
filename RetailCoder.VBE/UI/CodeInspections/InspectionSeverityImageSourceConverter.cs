using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Rubberduck.Inspections;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.CodeInspections
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

            return inspection.Name;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    public class InspectionSeverityImageSourceConverter : IValueConverter
    {
        private static readonly IDictionary<CodeInspectionSeverity,ImageSource> Icons = 
            new Dictionary<CodeInspectionSeverity, ImageSource>
            {
                { CodeInspectionSeverity.DoNotShow, null },
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
            throw new NotImplementedException();
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