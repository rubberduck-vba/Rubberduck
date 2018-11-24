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
            (string templateName, CodeExplorerItemViewModel model) data = (
                values[0] as string,
                values[1] as CodeExplorerItemViewModel);
            return data;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            var data = ((string templateName, CodeExplorerItemViewModel model))value;
            return new[] {(object) data.templateName, data.model};
        }
    }
}
