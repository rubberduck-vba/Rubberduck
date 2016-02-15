using System;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using System.Windows.Data;

namespace Rubberduck.UI.Settings.Converters
{
    public class ChildTextBoxSelectionToParentSelectionConverter : IMultiValueConverter
    {
        public object Convert(object[] value, Type targetType, object parameter, CultureInfo culture)
        {
            var childCheckBoxes = new List<bool?>();
            for (var i = 1; i < value.Length; i++)
            {
                childCheckBoxes.Add((bool?)value[i]);
            }

            if ((bool?) value[0] == false)
            {
                return false;
            }
            if (childCheckBoxes.All(v => v == true))
            {
                return true;
            }

            return null;
        }

        public object[] ConvertBack(object value, Type[] targetType, object parameter, CultureInfo culture)
        {
            return null;
        }
    }
}