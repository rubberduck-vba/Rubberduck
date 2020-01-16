using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using Rubberduck.UI.Refactorings;

namespace Rubberduck.UI.Converters
{
    class ClassInstancingToBooleanConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (ClassInstancing)value == ClassInstancing.Private;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (bool)value
                ? ClassInstancing.Private
                : ClassInstancing.PublicNotCreatable;
        }
    }
}
