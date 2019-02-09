using System;
using System.Globalization;
using System.Windows.Data;
using Rubberduck.Navigation.CodeExplorer;

namespace Rubberduck.UI.Converters
{
    public class TemplateCommandParameterToTupleConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            (string templateName, ICodeExplorerNode model) data = (
                values[0] as string,
                values[1] as ICodeExplorerNode);
            return data;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            var data = ((string templateName, ICodeExplorerNode model))value;
            return new[] {(object) data.templateName, data.model};
        }
    }
}
