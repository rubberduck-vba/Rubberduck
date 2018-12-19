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
                values.Length != 3 || 
                !(values[1] is ReferenceModel reference) ||         //DataContext
                reference.IsBuiltIn ||
                !(parameter is string direction))
            {
                return false;
            }

            var position = reference.Priority;                      
            var items = (int)values[0];                             //ProjectSelect.Items.Count
            var builtIn = (int)values[2];                           //AddRemoveReferencesWindow.DataContext.BuiltInReferenceCount

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
