using System;
using System.Globalization;
using System.Windows.Data;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Inspections
{
    class GroupingGridLocationConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is QualifiedModuleName qualifiedModuleName)
            {
                return $"{qualifiedModuleName} - {qualifiedModuleName.ComponentType}";
            }

            return Binding.DoNothing;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
