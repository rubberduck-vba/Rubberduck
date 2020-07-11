using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.Converters
{
    public class DeclarationToQualifiedNameConverter : IValueConverter
    {
        private readonly IValueConverter _declarationTypeConverter;

        public DeclarationToQualifiedNameConverter()
        {
            _declarationTypeConverter = new EnumToLocalizedStringConverter
            {
                ResourcePrefix = "DeclarationType_"
            };

        }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(value is Declaration declaration))
            {
                throw new ArgumentException("The value must be an instance of Declaration.", "value");
            }

            var qualifiedNameText = declaration.QualifiedName.ToString();
            var declarationTypeText = _declarationTypeConverter.Convert(declaration.DeclarationType, targetType, null, culture);

            return $"{qualifiedNameText} ({declarationTypeText})";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return DependencyProperty.UnsetValue;
        }
    }
}