using System;
using System.Globalization;
using System.Windows.Data;
using Rubberduck.AddRemoveReferences;

namespace Rubberduck.UI.AddRemoveReferences
{
    internal class PriorityButtonVisibilityConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values is null || 
                values.Length != 4 || 
                !(bool)values[0] ||                                 //IsSelected
                !(values[2] is ReferenceModel reference) ||         //DataContext
                reference.IsBuiltIn ||
                !(parameter is string direction))
            {
                return false;
            }

            var position = reference.Priority;                      //ProjectSelect.SelectedIndex
            var items = (int)values[1];                             //ProjectSelect.Items.Count
            var builtIn = (int)values[3];                           //AddRemoveReferencesWindow.DataContext.BuiltInReferenceCount

            if (direction.Equals("Up"))
            {
                return position > builtIn + 1;
            }

            return position != items;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new InvalidOperationException();
        }
    }
}
