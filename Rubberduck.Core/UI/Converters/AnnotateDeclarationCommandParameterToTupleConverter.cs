using System;
using System.Globalization;
using System.Windows.Data;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Annotations;

namespace Rubberduck.UI.Converters
{
    public class AnnotateDeclarationCommandParameterToTupleConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            (IAnnotation annotation, ICodeExplorerNode model) data = (
                values[0] as IAnnotation,
                values[1] as ICodeExplorerNode);
            return data;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            var data = ((IAnnotation annotation, ICodeExplorerNode model))value;
            return new[] { (object)data.annotation, data.model };
        }
    }
}