using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using Rubberduck.Parsing.Annotations;

namespace Rubberduck.UI.Converters
{
    public class AnnotationToCodeStringConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null)
            {
                return null;
            }

            if (!(value is IAnnotation annotation))
            {
                throw new ArgumentException("The value must be an instance of IAnnotation.", "value");
            }

            return $"@{annotation.Name}";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return DependencyProperty.UnsetValue;
        }
    }
}