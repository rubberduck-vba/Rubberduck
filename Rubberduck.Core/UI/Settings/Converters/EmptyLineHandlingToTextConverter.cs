using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Data;
using Rubberduck.Resources;
using Rubberduck.SmartIndenter;

namespace Rubberduck.UI.Settings.Converters
{
    public class EmptyLineHandlingToTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var styles = (IEnumerable<EmptyLineHandling>)value;
            return styles.Select(s => RubberduckUI.ResourceManager.GetString("EmptyLineHandling_" + s, CultureInfo.CurrentUICulture)).ToArray();
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value;
        }
    }
}
