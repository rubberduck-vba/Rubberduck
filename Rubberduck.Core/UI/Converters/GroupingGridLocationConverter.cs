using System;
using System.Globalization;
using System.Windows.Data;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.Converters
{
    class GroupingGridLocationConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is QualifiedModuleName qualifiedModuleName)
            {
                var componentTypeConverter = new ComponentTypeConverter();
                var localizedComponentType = (string)componentTypeConverter.Convert(qualifiedModuleName.ComponentType, typeof(ComponentType), parameter, culture);
                return $"{qualifiedModuleName} - {localizedComponentType}";
            }

            return Binding.DoNothing;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
