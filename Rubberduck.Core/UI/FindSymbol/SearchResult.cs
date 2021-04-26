using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.FindSymbol
{
    public class SearchResult
    {
        public SearchResult(Declaration declaration)
        {
            Declaration = declaration;
        }

        public Declaration Declaration { get; }

        public string IdentifierName => Declaration.IdentifierName;
        public string Location => Declaration.Scope;

    }

    public class SearchBoxMultiBindingConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            return values[0];
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            return value is Declaration declaration ? new[] { declaration.IdentifierName , value } : new[] { value, null };
        }
    }
}
